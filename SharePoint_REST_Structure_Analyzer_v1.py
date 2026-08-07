#!/usr/bin/env python3
"""
SharePoint REST Structure Analyzer v1

Accepts up to five SharePoint REST API URLs, analyzes all responses together,
and produces ONE self-contained HTML diagnostic report.

Dependencies:
    pip install requests beautifulsoup4
"""

from __future__ import annotations

import html as html_lib
import json
import re
import sys
import traceback
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple
from urllib.parse import urlparse

import requests
import urllib3
from bs4 import BeautifulSoup, Tag

APP_NAME = "SharePoint REST Structure Analyzer"
VERSION = "1.0"
MAX_SAMPLES = 5
OUTPUT_PREFIX = "SharePoint_REST_Structure_Analysis_Report"
INSECURE_SSL_HOSTS = {"w3.td.com"}
MAX_MATRIX_ROWS = 600
MAX_RAW_CHARS_PER_SAMPLE = 2_000_000
MAX_TEXT_PREVIEW = 5000


@dataclass
class HtmlAnalysis:
    source_field: str = ""
    length: int = 0
    text_length: int = 0
    max_depth: int = 0
    tag_counts: Dict[str, int] = field(default_factory=dict)
    classes: Set[str] = field(default_factory=set)
    ids: Set[str] = field(default_factory=set)
    data_attributes: Set[str] = field(default_factory=set)
    attributes: Set[str] = field(default_factory=set)
    structural_signals: Set[str] = field(default_factory=set)
    webpart_markers: Set[str] = field(default_factory=set)
    heading_texts: List[str] = field(default_factory=list)
    link_count: int = 0
    image_count: int = 0
    table_count: int = 0
    list_count: int = 0
    text_preview: str = ""


@dataclass
class SampleAnalysis:
    sample_no: int
    url: str
    final_url: str = ""
    ok: bool = False
    http_status: Optional[int] = None
    content_type: str = ""
    response_size: int = 0
    response_kind: str = "Unknown"
    elapsed_ms: Optional[int] = None
    error: str = ""
    page_title: str = ""
    top_level_fields: Set[str] = field(default_factory=set)
    flattened_paths: Set[str] = field(default_factory=set)
    path_types: Dict[str, str] = field(default_factory=dict)
    interesting_fields: Dict[str, str] = field(default_factory=dict)
    html_analyses: List[HtmlAnalysis] = field(default_factory=list)
    json_key_counts: Dict[str, int] = field(default_factory=dict)
    raw_text: str = ""


def esc(value: Any) -> str:
    return html_lib.escape("" if value is None else str(value))


def normalize_ws(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def shorten(value: str, limit: int = 220) -> str:
    value = normalize_ws(value)
    return value if len(value) <= limit else value[: limit - 3] + "..."


def typename(value: Any) -> str:
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, int) and not isinstance(value, bool):
        return "integer"
    if isinstance(value, float):
        return "number"
    if isinstance(value, dict):
        return "object"
    if isinstance(value, list):
        return "array"
    if isinstance(value, str):
        return "string"
    return type(value).__name__


def flatten_json(value: Any, prefix: str = "$", paths=None, path_types=None, key_counts=None,
                 depth: int = 0, max_depth: int = 40):
    if paths is None:
        paths = set()
    if path_types is None:
        path_types = {}
    if key_counts is None:
        key_counts = Counter()
    if depth > max_depth:
        return paths, path_types, key_counts

    paths.add(prefix)
    path_types[prefix] = typename(value)

    if isinstance(value, dict):
        for key, child in value.items():
            key_str = str(key)
            key_counts[key_str] += 1
            flatten_json(child, f"{prefix}.{key_str}", paths, path_types, key_counts, depth + 1, max_depth)
    elif isinstance(value, list):
        array_path = f"{prefix}[*]"
        paths.add(array_path)
        if value:
            for child in value[:200]:
                flatten_json(child, array_path, paths, path_types, key_counts, depth + 1, max_depth)
        else:
            path_types[array_path] = "empty-array"
    return paths, path_types, key_counts


def walk_json(value: Any, path: str = "$") -> Iterable[Tuple[str, str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}"
            yield child_path, str(key), child
            yield from walk_json(child, child_path)
    elif isinstance(value, list):
        for child in value[:500]:
            yield from walk_json(child, f"{path}[*]")


def looks_like_html_string(value: str) -> bool:
    if not isinstance(value, str) or len(value) < 10:
        return False
    low = value.lower()
    return any(x in low for x in ("<div", "<p", "<section", "<table", "<span", "<a ", "data-sp-", "<article"))


def dom_depth(tag: Tag) -> int:
    depth = 0
    parent = tag.parent
    while isinstance(parent, Tag):
        depth += 1
        parent = parent.parent
    return depth


def analyze_html_fragment(source_field: str, html_text: str) -> HtmlAnalysis:
    soup = BeautifulSoup(html_text or "", "html.parser")
    tag_counter = Counter()
    class_counter = Counter()
    classes, ids, data_attrs, attrs, webpart_markers, structural = set(), set(), set(), set(), set(), set()
    headings: List[str] = []
    max_depth = 0

    for tag in soup.find_all(True):
        name = (tag.name or "").lower()
        tag_counter[name] += 1
        max_depth = max(max_depth, dom_depth(tag))

        for attr_name, attr_value in tag.attrs.items():
            attr_name_l = str(attr_name).lower()
            attrs.add(attr_name_l)
            if attr_name_l.startswith("data-"):
                data_attrs.add(attr_name_l)
            if attr_name_l.startswith("data-sp-"):
                webpart_markers.add(attr_name_l)
            if attr_name_l == "class":
                vals = attr_value if isinstance(attr_value, list) else [attr_value]
                for c in vals:
                    c = str(c).strip()
                    if c:
                        classes.add(c)
                        class_counter[c] += 1
            if attr_name_l == "id":
                i = str(attr_value).strip()
                if i:
                    ids.add(i)

        if name in {"table", "ul", "ol", "section", "article", "details", "summary"}:
            structural.add(f"tag:{name}")

        role = str(tag.attrs.get("role", "")).strip()
        if role:
            structural.add(f"role:{role}")

        tag_classes = tag.attrs.get("class", [])
        if isinstance(tag_classes, str):
            tag_classes = [tag_classes]
        for c in tag_classes:
            c = str(c).strip()
            lc = c.lower()
            if c and any(token in lc for token in ("tab", "accordion", "expand", "collapse", "decision", "step", "branch", "section", "webpart", "web-part", "control", "canvas")):
                structural.add(f"class:{c}")

        if name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            txt = shorten(tag.get_text(" ", strip=True), 180)
            if txt:
                headings.append(txt)

    all_text = normalize_ws(soup.get_text(" ", strip=True))

    # Detect nested relationships between structurally interesting classes.
    interesting_tokens = ("tab", "accordion", "decision", "step", "section", "webpart", "canvas")
    for tag in soup.find_all(True):
        child_classes = tag.attrs.get("class", [])
        if isinstance(child_classes, str):
            child_classes = [child_classes]
        child_blob = " ".join(str(x).lower() for x in child_classes)
        if not any(tok in child_blob for tok in interesting_tokens):
            continue
        parent = tag.parent
        while isinstance(parent, Tag):
            parent_classes = parent.attrs.get("class", [])
            if isinstance(parent_classes, str):
                parent_classes = [parent_classes]
            parent_blob = " ".join(str(x).lower() for x in parent_classes)
            if any(tok in parent_blob for tok in interesting_tokens):
                structural.add(f"nested:{shorten(parent_blob, 80)} -> {shorten(child_blob, 80)}")
                break
            parent = parent.parent

    return HtmlAnalysis(
        source_field=source_field,
        length=len(html_text or ""),
        text_length=len(all_text),
        max_depth=max_depth,
        tag_counts=dict(tag_counter),
        classes=classes,
        ids=ids,
        data_attributes=data_attrs,
        attributes=attrs,
        structural_signals=structural,
        webpart_markers=webpart_markers,
        heading_texts=headings[:200],
        link_count=len(soup.find_all("a")),
        image_count=len(soup.find_all("img")),
        table_count=len(soup.find_all("table")),
        list_count=len(soup.find_all(["ul", "ol"])),
        text_preview=all_text[:MAX_TEXT_PREVIEW],
    )


def extract_interesting_json_fields(data: Any):
    interesting: Dict[str, str] = {}
    html_analyses: List[HtmlAnalysis] = []
    title = ""
    special_names = {
        "canvascontent1", "layoutwebpartscontent", "title", "pagelayouttype", "promotedstate",
        "topicheader", "description", "bannerimageurl", "authorbyline", "webpartdata",
        "controltype", "id", "uniqueid", "fileleafref", "fileref"
    }
    seen_html_hashes = set()

    for path, key, value in walk_json(data):
        kl = key.lower()
        if kl in special_names:
            preview = shorten(json.dumps(value, ensure_ascii=False), 500) if isinstance(value, (dict, list)) else shorten(str(value), 500)
            interesting[path] = preview
            if kl == "title" and not title and isinstance(value, str):
                title = value

        if isinstance(value, str) and (kl in {"canvascontent1", "layoutwebpartscontent"} or looks_like_html_string(value)):
            h = hash(value)
            if h not in seen_html_hashes:
                seen_html_hashes.add(h)
                html_analyses.append(analyze_html_fragment(path, value))

    return interesting, html_analyses, title


def detect_response_kind(response: requests.Response, raw_text: str):
    content_type = (response.headers.get("content-type") or "").lower()
    try:
        return "JSON", response.json()
    except Exception:
        pass
    stripped = raw_text.lstrip()
    if "xml" in content_type or stripped.startswith("<?xml") or stripped.startswith("<feed") or stripped.startswith("<entry"):
        return "XML", None
    if "html" in content_type or re.search(r"<(html|body|div|span|p|section|article)\b", stripped[:5000], re.I):
        return "HTML", None
    return "Text", None


def fetch_sample(sample_no: int, url: str) -> SampleAnalysis:
    result = SampleAnalysis(sample_no=sample_no, url=url)
    if not url:
        result.error = "No URL supplied."
        return result

    headers = {
        "Accept": "application/json;odata=verbose, application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36",
    }
    host = (urlparse(url).hostname or "").lower()
    bypass_ssl = host in INSECURE_SSL_HOSTS
    if bypass_ssl:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    try:
        response = requests.get(url, headers=headers, timeout=60, allow_redirects=True, verify=not bypass_ssl)
        result.final_url = response.url
        result.http_status = response.status_code
        result.content_type = response.headers.get("content-type", "")
        result.response_size = len(response.content or b"")
        result.elapsed_ms = int(response.elapsed.total_seconds() * 1000)
        raw = response.text or ""
        result.raw_text = raw[:MAX_RAW_CHARS_PER_SAMPLE]

        if response.status_code >= 400:
            result.error = f"HTTP {response.status_code}. Endpoint may require authentication/permissions or the URL may be invalid."
            return result

        result.ok = True
        kind, parsed_data = detect_response_kind(response, raw)
        result.response_kind = kind

        if kind == "JSON":
            data = parsed_data
            if isinstance(data, dict):
                result.top_level_fields = {str(k) for k in data.keys()}
            paths, path_types, key_counts = flatten_json(data)
            result.flattened_paths = paths
            result.path_types = path_types
            result.json_key_counts = dict(key_counts)
            interesting, html_analyses, title = extract_interesting_json_fields(data)
            result.interesting_fields = interesting
            result.html_analyses = html_analyses
            result.page_title = title
        elif kind in {"HTML", "XML"}:
            result.html_analyses = [analyze_html_fragment("$response", raw)]
            soup = BeautifulSoup(raw, "html.parser")
            if soup.title:
                result.page_title = normalize_ws(soup.title.get_text(" ", strip=True))

        return result
    except requests.exceptions.SSLError as exc:
        result.error = f"SSL error: {exc}"
    except requests.exceptions.RequestException as exc:
        result.error = f"Request error: {exc}"
    except Exception as exc:
        result.error = f"Unexpected error: {exc}"
    return result


def merge_html_sets(sample: SampleAnalysis, attr: str) -> Set[str]:
    merged = set()
    for analysis in sample.html_analyses:
        merged.update(getattr(analysis, attr, set()))
    return merged


def merged_tag_set(sample: SampleAnalysis) -> Set[str]:
    out = set()
    for h in sample.html_analyses:
        out.update(h.tag_counts.keys())
    return out


def build_presence_rows(samples: List[SampleAnalysis], item_sets: List[Set[str]], category: str):
    universe = set()
    for s in item_sets:
        universe.update(s)
    rows = []
    for item in sorted(universe, key=lambda x: x.lower()):
        presence = [item in s for s in item_sets]
        freq = sum(presence)
        finding = "Consistent" if freq == len(samples) else ("Unique" if freq == 1 else "Variant")
        rows.append({"category": category, "item": item, "presence": presence, "frequency": freq, "finding": finding})
    return rows


def prioritize_rows(rows: List[Dict[str, Any]]):
    rank = {"Consistent": 0, "Unique": 1, "Variant": 2}
    return sorted(rows, key=lambda r: (rank.get(r["finding"], 9), r["category"], r["item"].lower()))[:MAX_MATRIX_ROWS]


def count_totals(sample: SampleAnalysis):
    return {
        "html_fragments": len(sample.html_analyses),
        "classes": len(merge_html_sets(sample, "classes")),
        "data_attrs": len(merge_html_sets(sample, "data_attributes")),
        "links": sum(h.link_count for h in sample.html_analyses),
        "images": sum(h.image_count for h in sample.html_analyses),
        "tables": sum(h.table_count for h in sample.html_analyses),
        "lists": sum(h.list_count for h in sample.html_analyses),
        "max_depth": max([h.max_depth for h in sample.html_analyses] or [0]),
    }


def matrix_table(rows, sample_count, title):
    if not rows:
        return f"<h3>{esc(title)}</h3><p>No comparable items detected.</p>"
    ths = "".join(f"<th>S{i}</th>" for i in range(1, sample_count + 1))
    body = []
    for r in rows:
        checks = "".join(f"<td class='center'>{'✓' if p else '—'}</td>" for p in r["presence"])
        body.append(
            f"<tr class='{r['finding'].lower()}'><td>{esc(r['category'])}</td><td><code>{esc(r['item'])}</code></td>"
            f"{checks}<td class='center'>{r['frequency']}/{sample_count}</td><td><strong>{esc(r['finding'])}</strong></td></tr>"
        )
    return f"""
    <h3>{esc(title)}</h3>
    <div class='table-wrap'><table><thead><tr><th>Category</th><th>Attribute / Pattern</th>{ths}<th>Frequency</th><th>Finding</th></tr></thead>
    <tbody>{''.join(body)}</tbody></table></div>
    """


def sample_cards(samples):
    cards = []
    for s in samples:
        t = count_totals(s)
        cards.append(f"""
        <div class='card'>
          <div class='card-title'>Sample {s.sample_no}</div>
          <div class='{'ok' if s.ok else 'bad'}'><strong>{esc('HTTP '+str(s.http_status) if s.http_status is not None else 'No response')}</strong> · {esc(s.response_kind)}</div>
          <div><strong>Title:</strong> {esc(s.page_title or 'Not detected')}</div>
          <div><strong>Response:</strong> {s.response_size:,} bytes</div>
          <div><strong>HTML fragments:</strong> {t['html_fragments']}</div>
          <div><strong>Classes:</strong> {t['classes']} · <strong>data-*:</strong> {t['data_attrs']}</div>
          <div><strong>Max HTML depth:</strong> {t['max_depth']}</div>
          <div><strong>Links:</strong> {t['links']} · <strong>Tables:</strong> {t['tables']} · <strong>Images:</strong> {t['images']}</div>
          {f"<div class='bad'><strong>Error:</strong> {esc(s.error)}</div>" if s.error else ''}
        </div>""")
    return "<div class='cards'>" + "".join(cards) + "</div>"


def sample_detail(s: SampleAnalysis):
    targeted = "".join(f"<tr><td><code>{esc(p)}</code></td><td>{esc(v)}</td></tr>" for p, v in sorted(s.interesting_fields.items()))
    if not targeted:
        targeted = "<tr><td colspan='2'>No targeted fields detected.</td></tr>"

    fragments = []
    for i, h in enumerate(s.html_analyses, 1):
        fragments.append(f"""
        <details><summary>HTML Fragment {i}: <code>{esc(h.source_field)}</code></summary><div class='detail-body'>
        <table>
          <tr><th>HTML length</th><td>{h.length:,}</td></tr>
          <tr><th>Text length</th><td>{h.text_length:,}</td></tr>
          <tr><th>Maximum DOM depth</th><td>{h.max_depth}</td></tr>
          <tr><th>Links / Images / Tables / Lists</th><td>{h.link_count} / {h.image_count} / {h.table_count} / {h.list_count}</td></tr>
          <tr><th>CSS classes</th><td><code>{esc(', '.join(sorted(h.classes)[:100]))}</code></td></tr>
          <tr><th>data-* attributes</th><td><code>{esc(', '.join(sorted(h.data_attributes)[:100]))}</code></td></tr>
          <tr><th>Structural signals</th><td>{'<br>'.join('<code>'+esc(x)+'</code>' for x in sorted(h.structural_signals)[:100]) or 'None'}</td></tr>
          <tr><th>Headings</th><td>{'<br>'.join(esc(x) for x in h.heading_texts[:50]) or 'None'}</td></tr>
        </table>
        <h4>Clean text preview</h4><pre>{esc(h.text_preview)}</pre>
        </div></details>""")

    return f"""
    <section id='sample-{s.sample_no}'>
      <h2>Sample {s.sample_no}</h2>
      <table>
        <tr><th>Input URL</th><td class='break'><code>{esc(s.url)}</code></td></tr>
        <tr><th>Final URL</th><td class='break'><code>{esc(s.final_url)}</code></td></tr>
        <tr><th>HTTP status</th><td>{esc(s.http_status)}</td></tr>
        <tr><th>Content type</th><td>{esc(s.content_type)}</td></tr>
        <tr><th>Response kind</th><td>{esc(s.response_kind)}</td></tr>
        <tr><th>Response size</th><td>{s.response_size:,} bytes</td></tr>
        <tr><th>Elapsed</th><td>{esc(s.elapsed_ms)} ms</td></tr>
        <tr><th>Page title</th><td>{esc(s.page_title or 'Not detected')}</td></tr>
        <tr><th>Error</th><td>{esc(s.error or 'None')}</td></tr>
      </table>
      <h3>Targeted REST fields</h3>
      <div class='table-wrap'><table><thead><tr><th>JSON Path</th><th>Value Preview</th></tr></thead><tbody>{targeted}</tbody></table></div>
      <h3>HTML / Canvas diagnostics</h3>
      {''.join(fragments) if fragments else '<p>No embedded HTML/Canvas fragments detected.</p>'}
      <details><summary>Raw REST Response — Sample {s.sample_no}</summary><div class='detail-body'><pre>{esc(s.raw_text)}</pre></div></details>
    </section>"""


def build_report(samples: List[SampleAnalysis]) -> str:
    n = len(samples)
    rows = []
    rows += build_presence_rows(samples, [s.flattened_paths for s in samples], "REST / JSON Path")
    rows += build_presence_rows(samples, [merge_html_sets(s, "classes") for s in samples], "CSS Class")
    rows += build_presence_rows(samples, [merge_html_sets(s, "data_attributes") for s in samples], "data-* Attribute")
    rows += build_presence_rows(samples, [merge_html_sets(s, "attributes") for s in samples], "HTML Attribute")
    rows += build_presence_rows(samples, [merged_tag_set(s) for s in samples], "HTML Tag")
    rows += build_presence_rows(samples, [merge_html_sets(s, "structural_signals") for s in samples], "Structural Signal")
    rows += build_presence_rows(samples, [merge_html_sets(s, "webpart_markers") for s in samples], "SharePoint Marker")

    prioritized = prioritize_rows(rows)
    consistent = [r for r in prioritized if r["finding"] == "Consistent"]
    unique = [r for r in prioritized if r["finding"] == "Unique"]
    variants = [r for r in prioritized if r["finding"] == "Variant"]

    all_consistent = sum(1 for r in rows if r["finding"] == "Consistent")
    all_unique = sum(1 for r in rows if r["finding"] == "Unique")
    all_variants = sum(1 for r in rows if r["finding"] == "Variant")
    toc = "".join(f"<a href='#sample-{s.sample_no}'>Sample {s.sample_no}</a>" for s in samples)

    return f"""<!doctype html>
<html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>
<title>{esc(APP_NAME)} Report</title>
<style>
:root{{--border:#d0d7de;--soft:#f6f8fa;--text:#1f2328;--muted:#59636e;--ok:#1a7f37;--bad:#cf222e;--consistent:#eef8f0;--unique:#fff8c5;--variant:#f6f8fa}}
*{{box-sizing:border-box}} body{{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--text);line-height:1.45}} header{{padding:28px 34px;border-bottom:1px solid var(--border);background:var(--soft)}} main{{max-width:1500px;margin:0 auto;padding:28px 34px 60px}} h1{{margin:0 0 8px;font-size:30px}} h2{{margin-top:38px;padding-top:8px;border-top:1px solid var(--border)}} .subtitle{{color:var(--muted)}}
.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:12px;margin:16px 0 24px}} .card{{border:1px solid var(--border);border-radius:8px;padding:14px}} .card-title{{font-size:18px;font-weight:700;margin-bottom:8px}} .ok{{color:var(--ok)}} .bad{{color:var(--bad)}} nav{{display:flex;gap:12px;flex-wrap:wrap;margin-top:14px}} nav a{{color:#0969da;text-decoration:none}}
table{{width:100%;border-collapse:collapse;font-size:13px}} th,td{{border:1px solid var(--border);padding:7px 8px;text-align:left;vertical-align:top}} th{{background:var(--soft);position:sticky;top:0}} td.center{{text-align:center}} tr.consistent td{{background:var(--consistent)}} tr.unique td{{background:var(--unique)}} tr.variant td{{background:var(--variant)}} .table-wrap{{overflow:auto;max-height:650px;border:1px solid var(--border)}} code,pre{{font-family:Consolas,'Courier New',monospace}} code{{word-break:break-word}} pre{{white-space:pre-wrap;word-break:break-word;background:#0d1117;color:#e6edf3;padding:14px;border-radius:6px;max-height:650px;overflow:auto;font-size:12px}} details{{margin:10px 0;border:1px solid var(--border);border-radius:6px}} summary{{cursor:pointer;padding:10px 12px;font-weight:600;background:var(--soft)}} .detail-body{{padding:12px}} .break{{word-break:break-all}} .pill{{display:inline-block;border:1px solid var(--border);border-radius:999px;padding:3px 8px;margin:2px 4px 2px 0;font-size:12px}}
</style></head><body>
<header><h1>{esc(APP_NAME)}</h1><div class='subtitle'>Version {VERSION} · Consolidated 5-sample diagnostic report · Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div><nav><a href='#comparison'>Comparison Matrix</a><a href='#consistent'>Consistent</a><a href='#unique'>Unique</a><a href='#variants'>Variants</a>{toc}</nav></header>
<main>
<section><h2>Executive Diagnostic Summary</h2><p>This report analyzes all supplied SharePoint REST responses together. It does not assign business meaning or Level 1/2/3 classification. Its purpose is to expose structural patterns that can support future SharePoint programs.</p>{sample_cards(samples)}
<div><span class='pill'><strong>{len(rows):,}</strong> comparable attributes/patterns</span><span class='pill'><strong>{all_consistent:,}</strong> consistent</span><span class='pill'><strong>{all_variants:,}</strong> shared variants</span><span class='pill'><strong>{all_unique:,}</strong> unique variants</span></div></section>
<section id='comparison'><h2>Consolidated Attribute Comparison Matrix</h2><p><strong>Consistent</strong> = present in every sample. <strong>Variant</strong> = present in more than one but not all. <strong>Unique</strong> = present in only one.</p>{matrix_table(prioritized,n,'All Prioritized Structural Attributes')}</section>
<section id='consistent'><h2>Consistent Attributes Across All Samples</h2>{matrix_table(consistent,n,'Consistent Attributes')}</section>
<section id='unique'><h2>Unique Variants</h2>{matrix_table(unique,n,'Unique Attributes')}</section>
<section id='variants'><h2>Shared Variants</h2>{matrix_table(variants,n,'Shared Variants')}</section>
<section><h2>Individual Sample Evidence</h2><p>The sections below preserve targeted REST fields, Canvas/HTML analysis, and the raw REST response used in the comparison.</p></section>
{''.join(sample_detail(s) for s in samples)}
</main></body></html>"""


def prompt_urls():
    print("=" * 78)
    print(f"{APP_NAME} v{VERSION}")
    print("=" * 78)
    print("\nPaste up to five SharePoint REST API URLs.\nPress Enter to leave a slot blank.\n")
    urls = []
    for i in range(1, MAX_SAMPLES + 1):
        urls.append(input(f"REST API URL {i}: ").strip())
    return urls if any(urls) else []


def main():
    urls = prompt_urls()
    if not urls:
        print("\nERROR: No REST API URLs supplied.")
        return 1

    samples = []
    print("\nAnalyzing SharePoint REST responses...\n")
    for i, url in enumerate(urls, 1):
        if not url:
            samples.append(SampleAnalysis(sample_no=i, url="", error="No URL supplied."))
            print(f"Sample {i}: skipped")
            continue
        print(f"Sample {i}: requesting REST endpoint...")
        s = fetch_sample(i, url)
        samples.append(s)
        print(f"  {'OK' if s.ok else 'FAILED'} - {('HTTP '+str(s.http_status)) if s.http_status is not None else s.error}")

    report = build_report(samples)
    out = Path(f"{OUTPUT_PREFIX}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    out.write_text(report, encoding="utf-8")
    print("\nREPORT CREATED")
    print(out.resolve())
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(130)
    except Exception:
        print("\nUNEXPECTED ERROR")
        traceback.print_exc()
        sys.exit(2)
