#!/usr/bin/env python3
"""
SharePoint REST Structure Analyzer v6 - Native Edge + Clipboard Capture

Up to five SharePoint REST API URLs in -> one consolidated HTML report out.

This version uses the user's normal authenticated Microsoft Edge session.
It does NOT use Selenium and does NOT use remote debugging.

For each URL it:
1. Opens the REST URL in normal Edge.
2. Waits for the response to load.
3. Activates Edge.
4. Sends Ctrl+A and Ctrl+C.
5. Reads the response from the Windows clipboard.
6. Analyzes JSON/XML/text and compares all samples.

Dependency:
    pip install beautifulsoup4
"""

from __future__ import annotations

import html as html_lib
import json
import os
import re
import subprocess
import sys
import time
import traceback
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from bs4 import BeautifulSoup, Tag

APP_NAME = "SharePoint REST Structure Analyzer"
VERSION = "6.0 Native Edge + Clipboard Capture"
MAX_SAMPLES = 5
OUTPUT_PREFIX = "SharePoint_REST_Structure_Analysis_Report"
MAX_MATRIX_ROWS = 800
MAX_RAW_CHARS = 2_000_000
PAGE_LOAD_SECONDS = 5
CLIPBOARD_WAIT_SECONDS = 2


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
    headings: List[str] = field(default_factory=list)
    link_count: int = 0
    image_count: int = 0
    table_count: int = 0
    list_count: int = 0
    text_preview: str = ""


@dataclass
class Sample:
    number: int
    url: str
    ok: bool = False
    browser: str = "Microsoft Edge - Existing Session"
    response_kind: str = "Unknown"
    page_title: str = ""
    raw_text: str = ""
    error: str = ""
    top_fields: Set[str] = field(default_factory=set)
    json_paths: Set[str] = field(default_factory=set)
    path_types: Dict[str, str] = field(default_factory=dict)
    html_analyses: List[HtmlAnalysis] = field(default_factory=list)
    interesting_fields: Dict[str, str] = field(default_factory=dict)


def esc(value: Any) -> str:
    return html_lib.escape("" if value is None else str(value))


def norm(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def short(value: str, limit: int = 250) -> str:
    value = norm(value)
    return value if len(value) <= limit else value[:limit - 3] + "..."


def type_name(value: Any) -> str:
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


def flatten_json(
    value: Any,
    path: str = "$",
    paths: Optional[Set[str]] = None,
    types: Optional[Dict[str, str]] = None,
    depth: int = 0,
) -> Tuple[Set[str], Dict[str, str]]:
    if paths is None:
        paths = set()
    if types is None:
        types = {}

    if depth > 50:
        return paths, types

    paths.add(path)
    types[path] = type_name(value)

    if isinstance(value, dict):
        for key, child in value.items():
            flatten_json(child, f"{path}.{key}", paths, types, depth + 1)
    elif isinstance(value, list):
        array_path = f"{path}[*]"
        paths.add(array_path)
        if not value:
            types[array_path] = "empty-array"
        else:
            for child in value[:250]:
                flatten_json(child, array_path, paths, types, depth + 1)

    return paths, types


def walk_json(value: Any, path: str = "$") -> Iterable[Tuple[str, str, Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = f"{path}.{key}"
            yield child_path, str(key), child
            yield from walk_json(child, child_path)
    elif isinstance(value, list):
        for child in value[:500]:
            yield from walk_json(child, f"{path}[*]")


def looks_like_html(value: str) -> bool:
    if not isinstance(value, str) or len(value) < 10:
        return False
    low = value.lower()
    return any(
        marker in low
        for marker in (
            "<div", "<p", "<section", "<article", "<table",
            "<span", "<a ", "data-sp-", "<img"
        )
    )


def dom_depth(tag: Tag) -> int:
    result = 0
    parent = tag.parent
    while isinstance(parent, Tag):
        result += 1
        parent = parent.parent
    return result


def analyze_html(source: str, text: str) -> HtmlAnalysis:
    soup = BeautifulSoup(text or "", "html.parser")
    tags = Counter()
    classes: Set[str] = set()
    ids: Set[str] = set()
    data_attrs: Set[str] = set()
    attrs: Set[str] = set()
    signals: Set[str] = set()
    headings: List[str] = []
    max_depth = 0

    for tag in soup.find_all(True):
        name = (tag.name or "").lower()
        tags[name] += 1
        max_depth = max(max_depth, dom_depth(tag))

        for attr_name, attr_value in tag.attrs.items():
            attr_name = str(attr_name).lower()
            attrs.add(attr_name)

            if attr_name.startswith("data-"):
                data_attrs.add(attr_name)

            if attr_name == "id":
                value = str(attr_value).strip()
                if value:
                    ids.add(value)

            if attr_name == "class":
                values = attr_value if isinstance(attr_value, list) else [attr_value]
                for css_class in values:
                    css_class = str(css_class).strip()
                    if not css_class:
                        continue
                    classes.add(css_class)
                    low = css_class.lower()
                    if any(
                        token in low
                        for token in (
                            "tab", "accordion", "decision", "branch", "step",
                            "expand", "collapse", "section", "canvas",
                            "webpart", "web-part", "control"
                        )
                    ):
                        signals.add("class:" + css_class)

        role = str(tag.attrs.get("role", "")).strip()
        if role:
            signals.add("role:" + role)

        if name in {"table", "ul", "ol", "section", "article", "details", "summary"}:
            signals.add("tag:" + name)

        if name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            heading = short(tag.get_text(" ", strip=True), 180)
            if heading:
                headings.append(heading)

    clean_text = norm(soup.get_text(" ", strip=True))

    return HtmlAnalysis(
        source_field=source,
        length=len(text or ""),
        text_length=len(clean_text),
        max_depth=max_depth,
        tag_counts=dict(tags),
        classes=classes,
        ids=ids,
        data_attributes=data_attrs,
        attributes=attrs,
        structural_signals=signals,
        headings=headings[:200],
        link_count=len(soup.find_all("a")),
        image_count=len(soup.find_all("img")),
        table_count=len(soup.find_all("table")),
        list_count=len(soup.find_all(["ul", "ol"])),
        text_preview=clean_text[:5000],
    )


def extract_json(
    data: Any,
) -> Tuple[Set[str], Dict[str, str], List[HtmlAnalysis], Dict[str, str], str]:
    paths, types = flatten_json(data)
    analyses: List[HtmlAnalysis] = []
    interesting: Dict[str, str] = {}
    title = ""
    important_names = {
        "canvascontent1", "layoutwebpartscontent", "title",
        "pagelayouttype", "promotedstate", "description",
        "bannerimageurl", "webpartdata", "controltype",
        "id", "uniqueid", "fileleafref", "fileref"
    }
    seen_html = set()

    if isinstance(data, dict):
        for key, value in data.items():
            if str(key).lower() == "title" and isinstance(value, str):
                title = value

    for path, key, value in walk_json(data):
        key_low = key.lower()

        if key_low in important_names:
            if isinstance(value, (dict, list)):
                preview = short(json.dumps(value, ensure_ascii=False), 600)
            else:
                preview = short(str(value), 600)
            interesting[path] = preview

            if key_low == "title" and not title and isinstance(value, str):
                title = value

        if isinstance(value, str) and (
            key_low in {"canvascontent1", "layoutwebpartscontent"}
            or looks_like_html(value)
        ):
            fingerprint = hash(value)
            if fingerprint not in seen_html:
                seen_html.add(fingerprint)
                analyses.append(analyze_html(path, value))

    return paths, types, analyses, interesting, title


def find_edge_executable() -> str:
    candidates = [
        os.path.join(
            os.environ.get("PROGRAMFILES(X86)", ""),
            "Microsoft", "Edge", "Application", "msedge.exe"
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES", ""),
            "Microsoft", "Edge", "Application", "msedge.exe"
        ),
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Microsoft", "Edge", "Application", "msedge.exe"
        ),
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    raise RuntimeError("Microsoft Edge executable could not be found.")


def open_in_existing_edge(url: str) -> None:
    edge = find_edge_executable()
    subprocess.Popen(
        [edge, "--new-tab", url],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def powershell(script: str, timeout: int = 20) -> subprocess.CompletedProcess:
    return subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-Command", script
        ],
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def clear_clipboard() -> None:
    powershell("Set-Clipboard -Value ''")


def get_clipboard() -> str:
    result = powershell("Get-Clipboard -Raw", timeout=20)
    if result.returncode != 0:
        return ""
    return result.stdout or ""


def activate_edge_and_copy() -> None:
    ps = r"""
$wshell = New-Object -ComObject WScript.Shell
$activated = $false

for ($i = 0; $i -lt 20; $i++) {
    if ($wshell.AppActivate("Microsoft Edge")) {
        $activated = $true
        break
    }
    Start-Sleep -Milliseconds 250
}

if (-not $activated) {
    $p = Get-Process msedge -ErrorAction SilentlyContinue |
         Where-Object { $_.MainWindowHandle -ne 0 } |
         Select-Object -First 1

    if ($p) {
        $activated = $wshell.AppActivate($p.Id)
    }
}

if (-not $activated) {
    Write-Error "Could not activate Microsoft Edge."
    exit 1
}

Start-Sleep -Milliseconds 500
$wshell.SendKeys("^a")
Start-Sleep -Milliseconds 300
$wshell.SendKeys("^c")
Start-Sleep -Milliseconds 800
"""
    result = powershell(ps, timeout=20)

    if result.returncode != 0:
        raise RuntimeError(
            "Could not activate Microsoft Edge and copy the page. "
            + (result.stderr or "").strip()
        )


def looks_like_error_or_signin(raw: str) -> bool:
    low = (raw or "").lower()
    bad_markers = (
        "this page isn't working",
        "this page isn’t working",
        "can't reach this page",
        "can’t reach this page",
        "sign in",
        "signin",
        "login",
        "log in",
        "unauthorized",
        "access denied",
        "error-code",
    )
    return any(marker in low for marker in bad_markers)


def parse_json_text(raw: str) -> Tuple[Optional[Any], str]:
    raw = (raw or "").strip()
    if not raw:
        return None, raw

    try:
        return json.loads(raw), raw
    except Exception:
        pass

    starts = []
    object_index = raw.find("{")
    array_index = raw.find("[")

    if object_index >= 0:
        starts.append(object_index)
    if array_index >= 0:
        starts.append(array_index)

    for start in sorted(starts):
        candidate = raw[start:]
        try:
            return json.loads(candidate), candidate
        except Exception:
            continue

    return None, raw


def capture_sample(number: int, url: str) -> Sample:
    sample = Sample(number=number, url=url)

    try:
        clear_clipboard()

        print(f"Sample {number}: opening REST API URL in normal Edge...")
        open_in_existing_edge(url)

        print(f"  Waiting {PAGE_LOAD_SECONDS} seconds for the REST response...")
        time.sleep(PAGE_LOAD_SECONDS)

        print("  Capturing page content from Edge...")
        activate_edge_and_copy()
        time.sleep(CLIPBOARD_WAIT_SECONDS)

        raw = get_clipboard().strip()

        if not raw:
            sample.error = (
                "No content was captured from Edge. "
                "The REST tab may not have been active."
            )
            return sample

        sample.raw_text = raw[:MAX_RAW_CHARS]

        if looks_like_error_or_signin(raw):
            sample.response_kind = "Browser/Auth Page"
            sample.error = (
                "Edge copied an error/sign-in/access page instead of the REST payload. "
                f"Preview: {short(raw, 300)}"
            )
            return sample

        parsed, cleaned = parse_json_text(raw)

        if parsed is not None:
            sample.response_kind = "JSON"
            sample.raw_text = cleaned[:MAX_RAW_CHARS]

            if isinstance(parsed, dict):
                sample.top_fields = {str(k) for k in parsed.keys()}

            (
                sample.json_paths,
                sample.path_types,
                sample.html_analyses,
                sample.interesting_fields,
                sample.page_title,
            ) = extract_json(parsed)

            sample.ok = True
            return sample

        low = raw.lstrip().lower()

        if (
            low.startswith("<?xml")
            or low.startswith("<feed")
            or low.startswith("<entry")
        ):
            sample.response_kind = "XML"
            sample.html_analyses = [analyze_html("$xml_response", raw)]
            sample.ok = True
            return sample

        sample.response_kind = "Text"

        if len(raw) < 500:
            sample.error = (
                "Captured content is too small to be a useful REST payload. "
                f"Only {len(raw)} characters were captured. "
                f"Preview: {short(raw, 300)}"
            )
            return sample

        sample.html_analyses = [analyze_html("$captured_browser_text", raw)]
        sample.ok = True
        return sample

    except Exception as exc:
        sample.error = str(exc)
        return sample


def merged_set(sample: Sample, attribute: str) -> Set[str]:
    output: Set[str] = set()
    for analysis in sample.html_analyses:
        output.update(getattr(analysis, attribute, set()))
    return output


def tag_set(sample: Sample) -> Set[str]:
    output: Set[str] = set()
    for analysis in sample.html_analyses:
        output.update(analysis.tag_counts.keys())
    return output


def build_presence_rows(
    samples: List[Sample],
    sets: List[Set[str]],
    category: str,
) -> List[Dict[str, Any]]:
    universe: Set[str] = set()
    for values in sets:
        universe.update(values)

    rows: List[Dict[str, Any]] = []
    sample_count = len(samples)

    for item in sorted(universe, key=str.lower):
        presence = [item in values for values in sets]
        frequency = sum(presence)

        if frequency == sample_count:
            finding = "Consistent"
        elif frequency == 1:
            finding = "Unique"
        else:
            finding = "Variant"

        rows.append(
            {
                "category": category,
                "item": item,
                "presence": presence,
                "frequency": frequency,
                "finding": finding,
            }
        )
    return rows


def matrix(rows: List[Dict[str, Any]], sample_count: int, heading: str) -> str:
    if not rows:
        return f"<h3>{esc(heading)}</h3><p>No comparable structures detected.</p>"

    sample_headers = "".join(
        f"<th>S{i}</th>" for i in range(1, sample_count + 1)
    )

    body = []

    for row in rows:
        checks = "".join(
            f"<td class='center'>{'✓' if present else '—'}</td>"
            for present in row["presence"]
        )

        body.append(
            f"<tr class='{row['finding'].lower()}'>"
            f"<td>{esc(row['category'])}</td>"
            f"<td><code>{esc(row['item'])}</code></td>"
            f"{checks}"
            f"<td class='center'>{row['frequency']}/{sample_count}</td>"
            f"<td><strong>{esc(row['finding'])}</strong></td>"
            "</tr>"
        )

    return f"""
    <h3>{esc(heading)}</h3>
    <div class="scroll">
      <table>
        <thead>
          <tr>
            <th>Category</th>
            <th>Attribute / Pattern</th>
            {sample_headers}
            <th>Frequency</th>
            <th>Finding</th>
          </tr>
        </thead>
        <tbody>{''.join(body)}</tbody>
      </table>
    </div>
    """


def sample_details(sample: Sample) -> str:
    interesting_rows = "".join(
        f"<tr><td><code>{esc(path)}</code></td><td>{esc(value)}</td></tr>"
        for path, value in sorted(sample.interesting_fields.items())
    ) or "<tr><td colspan='2'>No targeted fields detected.</td></tr>"

    html_blocks = []

    for index, analysis in enumerate(sample.html_analyses, start=1):
        structural = "<br>".join(
            "<code>" + esc(value) + "</code>"
            for value in sorted(analysis.structural_signals)[:200]
        ) or "None"

        html_blocks.append(
            f"""
            <details>
              <summary>HTML / Canvas Fragment {index}: <code>{esc(analysis.source_field)}</code></summary>
              <div class="pad">
                <table>
                  <tr><th>HTML length</th><td>{analysis.length:,}</td></tr>
                  <tr><th>Maximum DOM depth</th><td>{analysis.max_depth}</td></tr>
                  <tr><th>Links / Images / Tables / Lists</th>
                      <td>{analysis.link_count} / {analysis.image_count} / {analysis.table_count} / {analysis.list_count}</td></tr>
                  <tr><th>CSS classes</th><td><code>{esc(", ".join(sorted(analysis.classes)[:200]))}</code></td></tr>
                  <tr><th>data-* attributes</th><td><code>{esc(", ".join(sorted(analysis.data_attributes)[:200]))}</code></td></tr>
                  <tr><th>Structural signals</th><td>{structural}</td></tr>
                </table>
                <h4>Clean text preview</h4>
                <pre>{esc(analysis.text_preview)}</pre>
              </div>
            </details>
            """
        )

    return f"""
    <section>
      <h2>Sample {sample.number}</h2>
      <table>
        <tr><th>REST API URL</th><td class="break"><code>{esc(sample.url)}</code></td></tr>
        <tr><th>Browser</th><td>{esc(sample.browser)}</td></tr>
        <tr><th>Status</th><td>{'Success' if sample.ok else 'Failed'}</td></tr>
        <tr><th>Response type</th><td>{esc(sample.response_kind)}</td></tr>
        <tr><th>Page title</th><td>{esc(sample.page_title or 'Not detected')}</td></tr>
        <tr><th>Error</th><td>{esc(sample.error or 'None')}</td></tr>
      </table>

      <h3>Targeted REST Fields</h3>
      <div class="scroll">
        <table>
          <thead><tr><th>JSON Path</th><th>Value Preview</th></tr></thead>
          <tbody>{interesting_rows}</tbody>
        </table>
      </div>

      {''.join(html_blocks) if html_blocks else '<p>No embedded HTML / Canvas fragments detected.</p>'}

      <details>
        <summary>Raw REST Response — Sample {sample.number}</summary>
        <div class="pad"><pre>{esc(sample.raw_text)}</pre></div>
      </details>
    </section>
    """


def build_report(samples: List[Sample]) -> str:
    sample_count = len(samples)
    rows: List[Dict[str, Any]] = []

    rows += build_presence_rows(samples, [s.json_paths for s in samples], "REST / JSON Path")
    rows += build_presence_rows(samples, [merged_set(s, "classes") for s in samples], "CSS Class")
    rows += build_presence_rows(samples, [merged_set(s, "ids") for s in samples], "HTML ID")
    rows += build_presence_rows(samples, [merged_set(s, "data_attributes") for s in samples], "data-* Attribute")
    rows += build_presence_rows(samples, [merged_set(s, "attributes") for s in samples], "HTML Attribute")
    rows += build_presence_rows(samples, [tag_set(s) for s in samples], "HTML Tag")
    rows += build_presence_rows(samples, [merged_set(s, "structural_signals") for s in samples], "Structural Signal")

    rank = {"Consistent": 0, "Unique": 1, "Variant": 2}

    ordered = sorted(
        rows,
        key=lambda row: (
            rank[row["finding"]],
            row["category"],
            row["item"].lower(),
        ),
    )[:MAX_MATRIX_ROWS]

    consistent = [r for r in ordered if r["finding"] == "Consistent"]
    unique = [r for r in ordered if r["finding"] == "Unique"]
    variants = [r for r in ordered if r["finding"] == "Variant"]
    success_count = sum(1 for s in samples if s.ok)

    cards = []

    for sample in samples:
        status_class = "ok" if sample.ok else "bad"
        status_text = "Success" if sample.ok else "Failed"

        cards.append(
            f"""
            <div class="card">
              <strong>Sample {sample.number}</strong><br>
              <span class="{status_class}">{status_text}</span><br>
              {esc(sample.response_kind)}<br>
              {esc(sample.page_title or 'Title not detected')}
              {'<br><span class="bad">' + esc(sample.error) + '</span>' if sample.error else ''}
            </div>
            """
        )

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(APP_NAME)} Report</title>
<style>
body{{font-family:Segoe UI,Arial,sans-serif;margin:0;color:#1f2328;line-height:1.45}}
header{{background:#f6f8fa;border-bottom:1px solid #d0d7de;padding:26px 32px}}
main{{max-width:1500px;margin:auto;padding:26px 32px 60px}}
h1{{margin:0}} h2{{margin-top:38px;border-top:1px solid #d0d7de;padding-top:14px}}
.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px;margin:18px 0}}
.card{{border:1px solid #d0d7de;border-radius:8px;padding:12px}}
.ok{{color:#1a7f37}} .bad{{color:#cf222e}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th,td{{border:1px solid #d0d7de;padding:7px;vertical-align:top;text-align:left}}
th{{background:#f6f8fa;position:sticky;top:0}}
.center{{text-align:center}} .consistent td{{background:#eef8f0}}
.unique td{{background:#fff8c5}} .variant td{{background:#f6f8fa}}
.scroll{{overflow:auto;max-height:650px;border:1px solid #d0d7de}}
pre{{background:#0d1117;color:#e6edf3;padding:14px;border-radius:6px;white-space:pre-wrap;word-break:break-word;max-height:650px;overflow:auto;font-size:12px}}
code{{font-family:Consolas,monospace;word-break:break-word}}
details{{border:1px solid #d0d7de;border-radius:6px;margin:10px 0}}
summary{{cursor:pointer;background:#f6f8fa;padding:10px;font-weight:600}}
.pad{{padding:12px}} .break{{word-break:break-all}}
</style>
</head>
<body>
<header>
<h1>{esc(APP_NAME)}</h1>
<div>Version {esc(VERSION)} · Generated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>
</header>

<main>
<h2>Consolidated Diagnostic Summary</h2>
<p><strong>{success_count}/{sample_count}</strong> samples were successfully captured from the normal authenticated Edge session.</p>
<p>The analyzer compares the SharePoint REST structures objectively. It does not yet assign Level 1 / Level 2 / Level 3 procedure meaning.</p>

<div class="cards">{''.join(cards)}</div>

<h2>Consolidated Attribute Comparison Matrix</h2>
<p><strong>Consistent</strong> = present in every sample.
<strong>Variant</strong> = present in more than one but not all samples.
<strong>Unique</strong> = present in one sample only.</p>

{matrix(ordered, sample_count, "All Prioritized Attributes and Patterns")}

<h2>Consistent Attributes</h2>
{matrix(consistent, sample_count, "Present Across All Samples")}

<h2>Unique Variants</h2>
{matrix(unique, sample_count, "Present in One Sample Only")}

<h2>Shared Variants</h2>
{matrix(variants, sample_count, "Present in Multiple But Not All Samples")}

<h2>Individual Sample Evidence</h2>
{''.join(sample_details(s) for s in samples)}
</main>
</body>
</html>
"""


def prompt_urls() -> List[str]:
    print("=" * 78)
    print(f"{APP_NAME} v{VERSION}")
    print("=" * 78)
    print()
    print("Paste up to five SharePoint REST API URLs.")
    print("Press Enter to leave a sample slot blank.")
    print()

    urls: List[str] = []

    for index in range(1, MAX_SAMPLES + 1):
        urls.append(input(f"REST API URL {index}: ").strip())

    if not any(urls):
        print("\nERROR: No URLs supplied.")
        return []

    return urls


def main() -> int:
    if os.name != "nt":
        print("ERROR: This version is designed for Windows.")
        return 2

    urls = prompt_urls()
    if not urls:
        return 1

    print()
    print("IMPORTANT")
    print("Keep your normal Microsoft Edge session open and signed in to SharePoint.")
    print("Do not use the keyboard or mouse while each sample is being captured.")
    print()
    input("Press Enter when Edge is open and authenticated: ")

    samples: List[Sample] = []

    print("\nStarting capture...\n")

    for index, url in enumerate(urls, start=1):
        if not url:
            samples.append(Sample(number=index, url="", error="No URL supplied."))
            print(f"Sample {index}: skipped\n")
            continue

        sample = capture_sample(index, url)
        samples.append(sample)

        if sample.ok:
            print(
                f"  SUCCESS - {sample.response_kind}, "
                f"{len(sample.raw_text):,} characters"
            )
        else:
            print(f"  FAILED - {sample.error}")

        print()

    print("Building consolidated HTML report...")

    report = build_report(samples)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = Path(f"{OUTPUT_PREFIX}_{timestamp}.html")
    output.write_text(report, encoding="utf-8")

    print("\n" + "=" * 78)
    print("REPORT CREATED")
    print("=" * 78)
    print(output.resolve())
    print()
    print("A valid SharePoint REST capture should normally show JSON.")
    print()

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
