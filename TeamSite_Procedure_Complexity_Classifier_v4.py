#!/usr/bin/env python3
"""
TeamSite Procedure Complexity Classifier v4

Purpose
-------
Reads a Legacy / TeamSite procedure URL directly and classifies the page as:

    Level 1 - Straightforward
    Level 2 - Decision Points
    Level 3 - Nested Decision Points

The program does NOT require SharePoint.  It classifies the original Legacy
structure, which is the safest place to identify decision-point nesting before
the migration transformation flattens the content.

Input
-----
A Legacy URL, for example:
https://w3.td.com/td/intranet/legacy/content?DOCID=...&LOCALE=en_ca

Output
------
Creates a folder named:
    teamsite_complexity_output

and writes:
    <DOCID>_complexity.json
    <DOCID>_complexity.csv
    <DOCID>_evidence.html
    <DOCID>_source.html

Dependencies
------------
pip install requests beautifulsoup4

No username/password is stored or requested.
"""

from __future__ import annotations

import csv
import json
import re
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Tuple
from urllib.parse import urlparse, parse_qs

import requests
import urllib3
from bs4 import BeautifulSoup, Tag


APP_NAME = "TeamSite Procedure Complexity Classifier"
VERSION = "4.0"

OUTPUT_DIR = Path("teamsite_complexity_output")

# Internal TeamSite host that may use a corporate/self-signed certificate.
# SSL verification is bypassed ONLY for this exact host.
INSECURE_SSL_HOST = "w3.td.com"

# Strong signals taken from the TeamSite / legacy structures already seen in
# the comparison tool and procedure markup.
DECISION_CLASS_PATTERNS = (
    "decision-point",
    "decisionpoint",
    "decision_point",
    "decision point",
    "procedure-decision",
    "proceduredecision",
    "decision-container",
    "decisioncontainer",

    # TeamSite legacy procedures frequently render decision branches as tabs.
    # Example observed markup:
    #   div.custom-wcm-wysiwyg-tabs-wrapper
    #   a.custom-wcm-wysiwyg-tabs-button
    #   span.td-tabs-label
    #   div.td-tabs-content-wrapper
    "custom-wcm-wysiwyg-tabs-wrapper",
)

# UI-only elements.  These can contain the words "Decision Point" but are not
# themselves the structural container that should be counted.
UI_ONLY_PATTERNS = (
    "decision-point-header",
    "decision-point-title",
    "decision-point-toggle",
    "decision-point-label",
    "custom-wcm-wysiwyg-tabs-button",
    "td-tabs-label",
    "td-tabs-active",
    "td-tabs-afteractive",
    "triggericon",
    "details-button",
    "toggle-button",
    "expand-button",
    "collapse-button",
)

DECISION_TEXT_RE = re.compile(r"\bdecision\s*point\b", re.I)
YES_NO_RE = re.compile(r"^\s*(yes|no)\s*$", re.I)


@dataclass
class DecisionEvidence:
    index: int
    depth: int
    tag: str
    element_id: str
    classes: str
    text_preview: str
    dom_path: str


@dataclass
class AnalysisResult:
    application: str
    version: str
    source_url: str
    docid: str
    locale: str
    page_title: str
    complexity_level: int
    complexity_name: str
    decision_points: int
    nested_decision_points: int
    maximum_decision_depth: int
    reason: str
    evidence: List[DecisionEvidence]
    nested_relationships: List[dict]


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def get_docid(url: str) -> str:
    query = parse_qs(urlparse(url).query)
    for key in ("DOCID", "docid", "DocID", "documentId", "documentid"):
        if key in query and query[key]:
            return query[key][0]
    return "unknown_docid"


def get_locale(url: str) -> str:
    query = parse_qs(urlparse(url).query)
    for key in ("LOCALE", "locale", "Locale"):
        if key in query and query[key]:
            return query[key][0]
    return ""


def safe_filename(value: str) -> str:
    value = normalize_spaces(value)
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("._") or "unknown"


def fetch_page(url: str) -> str:
    """
    Direct HTTP read only.
    No username/password is requested or saved.

    IMPORTANT:
    SSL verification is bypassed ONLY for the exact internal host w3.td.com.
    All other hosts continue to use normal certificate validation.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 Chrome/124 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    }

    parsed = urlparse(url)
    hostname = (parsed.hostname or "").lower()

    # Narrow bypass: only the known internal TeamSite host.
    bypass_ssl = hostname == INSECURE_SSL_HOST

    if bypass_ssl:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        print(f"SSL note: certificate verification is bypassed for {INSECURE_SSL_HOST} only.")

    session = requests.Session()
    response = session.get(
        url,
        headers=headers,
        timeout=45,
        allow_redirects=True,
        verify=not bypass_ssl,
    )

    if response.status_code in (401, 403):
        raise RuntimeError(
            f"Legacy site returned HTTP {response.status_code}. "
            "The SSL issue was bypassed, but the site is still rejecting the request "
            "because of authentication or access permissions."
        )

    response.raise_for_status()

    content_type = (response.headers.get("content-type") or "").lower()
    html = response.text

    # Defensive check against a common login redirect.
    low = html.lower()
    if (
        "text/html" in content_type
        and len(html) < 250000
        and any(x in low for x in ("sign in", "login", "log in"))
        and "password" in low
    ):
        raise RuntimeError(
            "The URL appears to have returned a login page instead of the procedure."
        )

    return html


def class_blob(tag: Tag) -> str:
    classes = tag.get("class", [])
    if isinstance(classes, str):
        classes = [classes]
    return " ".join(str(x) for x in classes)


def marker_blob(tag: Tag) -> str:
    parts = [
        tag.name or "",
        tag.get("id", "") or "",
        class_blob(tag),
        tag.get("role", "") or "",
        tag.get("data-type", "") or "",
        tag.get("data-component", "") or "",
        tag.get("data-content-type", "") or "",
    ]
    return normalize_spaces(" ".join(parts)).lower()


def is_ui_only(tag: Tag) -> bool:
    blob = marker_blob(tag)
    return any(p in blob for p in UI_ONLY_PATTERNS)


def has_decision_marker(tag: Tag) -> bool:
    blob = marker_blob(tag)
    if any(p in blob for p in DECISION_CLASS_PATTERNS):
        return True

    # We only use visible "Decision Point" text as a secondary indicator.
    # It is intentionally restricted to container-like elements so ordinary
    # paragraph text mentioning a decision point does not become a false hit.
    if tag.name in ("div", "section", "article", "li", "td"):
        own_text = normalize_spaces(
            " ".join(
                str(x)
                for x in tag.find_all(string=True, recursive=False)
                if normalize_spaces(str(x))
            )
        )
        if own_text and DECISION_TEXT_RE.search(own_text):
            return True

    return False


def meaningful_child_count(tag: Tag) -> int:
    count = 0
    for child in tag.find_all(recursive=False):
        if not isinstance(child, Tag):
            continue
        text = normalize_spaces(child.get_text(" ", strip=True))
        if text or child.find(["ul", "ol", "table", "a", "img"]):
            count += 1
    return count


def looks_like_structural_decision_container(tag: Tag) -> bool:
    """
    Separate real decision containers from headings/toggles/labels.

    A candidate is considered structural if:
      - it carries a decision marker,
      - it is not an obvious UI-only element,
      - and it has either meaningful child content or another decision element.
    """
    if not has_decision_marker(tag) or is_ui_only(tag):
        return False

    if tag.name not in ("div", "section", "article", "li", "td", "fieldset"):
        return False

    nested_marker = any(
        isinstance(desc, Tag) and desc is not tag and has_decision_marker(desc)
        for desc in tag.find_all(True)
    )

    text = normalize_spaces(tag.get_text(" ", strip=True))
    child_count = meaningful_child_count(tag)

    return nested_marker or child_count >= 2 or len(text) >= 40



def is_teamsite_tab_decision(tag: Tag) -> bool:
    """
    Detect the legacy TeamSite tab widget used to represent a decision point.

    The screenshot supplied for the procedure shows markup such as:
        div.custom-wcm-wysiwyg-tabs-wrapper
        ul
          li.td-tabs-active
            a.custom-wcm-wysiwyg-tabs-button
              span.td-tabs-label
        div.td-tabs-content-wrapper

    A whole tab wrapper is ONE decision point. Individual tabs are branches,
    not separate decision points.
    """
    if not isinstance(tag, Tag):
        return False

    classes = tag.get("class", [])
    if isinstance(classes, str):
        classes = [classes]
    classes_lower = {str(c).lower() for c in classes}

    if "custom-wcm-wysiwyg-tabs-wrapper" in classes_lower:
        return True

    # Fallback for pages where wrapper naming varies slightly:
    # require both tab buttons/labels and a tab content wrapper.
    if tag.name in ("div", "section"):
        has_button = bool(tag.select_one(
            "a.custom-wcm-wysiwyg-tabs-button, "
            ".td-tabs-label"
        ))
        has_content = bool(tag.select_one(
            ".td-tabs-content-wrapper, "
            "[id^='tabwrapper_']"
        ))
        if has_button and has_content:
            return True

    return False


def teamsite_tab_branch_count(tag: Tag) -> int:
    """Return the number of visible decision branches in a TeamSite tab widget."""
    if not is_teamsite_tab_decision(tag):
        return 0

    buttons = tag.select(
        "a.custom-wcm-wysiwyg-tabs-button, "
        "li > span.td-tabs-label"
    )

    # Deduplicate by normalized branch text because some TeamSite markup
    # exposes the same label through multiple nested elements.
    labels = []
    seen = set()
    for b in buttons:
        label = normalize_spaces(b.get_text(" ", strip=True))
        if not label:
            continue
        key = label.lower()
        if key not in seen:
            seen.add(key)
            labels.append(label)

    return len(labels)


def is_nested_tab_decision(tag: Tag) -> bool:
    """
    True when this tab decision is contained inside another TeamSite tab
    decision. This is the key Level 3 signal.
    """
    parent = tag.parent
    while isinstance(parent, Tag):
        if is_teamsite_tab_decision(parent):
            return True
        parent = parent.parent
    return False


def select_decision_containers(soup: BeautifulSoup) -> List[Tag]:
    """
    Find structural decision points.

    Priority 1:
      TeamSite tab wrappers. These are known legacy decision-point widgets.

    Priority 2:
      Other explicit decision-point containers from class/id/text markers.
    """
    candidates: List[Tag] = []

    # --- Known TeamSite tab decision points ---
    tab_wrappers = [
        tag for tag in soup.find_all(True)
        if isinstance(tag, Tag) and is_teamsite_tab_decision(tag)
    ]

    # Keep only the actual outer wrapper for each widget. The fallback detector
    # can also match inner containers, so explicit wrapper class wins.
    for tag in tab_wrappers:
        classes = tag.get("class", [])
        if isinstance(classes, str):
            classes = [classes]
        lower = {str(c).lower() for c in classes}

        if "custom-wcm-wysiwyg-tabs-wrapper" in lower:
            candidates.append(tag)
            continue

        # Fallback candidate is kept only if no ancestor is an explicit wrapper.
        ancestor_has_explicit_wrapper = False
        parent = tag.parent
        while isinstance(parent, Tag):
            pclasses = parent.get("class", [])
            if isinstance(pclasses, str):
                pclasses = [pclasses]
            if "custom-wcm-wysiwyg-tabs-wrapper" in {str(c).lower() for c in pclasses}:
                ancestor_has_explicit_wrapper = True
                break
            parent = parent.parent

        if not ancestor_has_explicit_wrapper:
            candidates.append(tag)

    # --- Other decision-point structures ---
    for tag in soup.find_all(True):
        if not isinstance(tag, Tag):
            continue
        if id(tag) in {id(x) for x in candidates}:
            continue
        if looks_like_structural_decision_container(tag):
            candidates.append(tag)

    # Deduplicate while preserving document order.
    seen = set()
    result: List[Tag] = []
    for tag in candidates:
        key = id(tag)
        if key not in seen:
            seen.add(key)
            result.append(tag)

    return result

def dom_path(tag: Tag) -> str:
    parts = []
    current: Optional[Tag] = tag

    while current and isinstance(current, Tag) and current.name not in ("html", "[document]"):
        label = current.name

        if current.get("id"):
            label += f"#{current.get('id')}"
            parts.append(label)
            break

        classes = current.get("class", [])
        if isinstance(classes, str):
            classes = [classes]
        if classes:
            useful = [c for c in classes if c][:2]
            if useful:
                label += "." + ".".join(useful)

        parts.append(label)
        current = current.parent if isinstance(current.parent, Tag) else None

    return " > ".join(reversed(parts))



def is_explicit_teamsite_decision_wrapper(tag: Tag) -> bool:
    """Return True only for the actual TeamSite decision wrapper element."""
    if not isinstance(tag, Tag):
        return False
    classes = tag.get("class", [])
    if isinstance(classes, str):
        classes = [classes]
    classes_lower = {str(c).lower() for c in classes}
    return "custom-wcm-wysiwyg-tabs-wrapper" in classes_lower


def is_tab_content_wrapper(tag: Tag) -> bool:
    """Return True for the TeamSite container holding branch content."""
    if not isinstance(tag, Tag):
        return False
    classes = tag.get("class", [])
    if isinstance(classes, str):
        classes = [classes]
    classes_lower = {str(c).lower() for c in classes}
    return "td-tabs-content-wrapper" in classes_lower


def nearest_ancestor_decision(tag: Tag) -> Optional[Tag]:
    """
    Find the nearest ancestor TeamSite decision wrapper.

    For a wrapper to be considered nested, the route from the inner wrapper
    to the outer wrapper must pass through a td-tabs-content-wrapper.  This
    prevents two independent tab groups on the same page from being treated
    as nested.
    """
    parent = tag.parent
    saw_tab_content_wrapper = False

    while isinstance(parent, Tag):
        if is_tab_content_wrapper(parent):
            saw_tab_content_wrapper = True

        if is_explicit_teamsite_decision_wrapper(parent):
            return parent if saw_tab_content_wrapper else None

        parent = parent.parent

    return None


def strict_decision_depth(tag: Tag) -> int:
    """
    Calculate true TeamSite nesting depth.

    Depth 1 = top-level decision
    Depth 2 = decision inside the content of another decision
    Depth 3+ = deeper nested decisions
    """
    if not is_explicit_teamsite_decision_wrapper(tag):
        return 1

    depth = 1
    current = tag

    while True:
        ancestor = nearest_ancestor_decision(current)
        if ancestor is None:
            break
        depth += 1
        current = ancestor

    return depth


def get_branch_labels(tag: Tag) -> List[str]:
    """Extract unique branch labels from one TeamSite decision wrapper."""
    labels: List[str] = []
    seen = set()

    for btn in tag.select("a.custom-wcm-wysiwyg-tabs-button, .td-tabs-label"):
        label = normalize_spaces(btn.get_text(" ", strip=True))
        if not label:
            continue
        key = label.lower()
        if key not in seen:
            seen.add(key)
            labels.append(label)

    return labels


def build_nested_relationships(candidates: List[Tag]) -> List[dict]:
    """
    Build explicit outer -> inner decision relationships for the report.
    """
    relationships = []

    for inner in candidates:
        if not is_explicit_teamsite_decision_wrapper(inner):
            continue

        outer = nearest_ancestor_decision(inner)
        if outer is None:
            continue

        relationships.append(
            {
                "outer_branches": get_branch_labels(outer),
                "inner_branches": get_branch_labels(inner),
                "outer_path": dom_path(outer),
                "inner_path": dom_path(inner),
            }
        )

    return relationships


def decision_depth(tag: Tag, candidates: List[Tag]) -> int:
    """
    Determine decision nesting depth.

    For TeamSite tab-based decisions, use the strict rule:
      inner custom-wcm-wysiwyg-tabs-wrapper
      must sit under a td-tabs-content-wrapper
      of an outer custom-wcm-wysiwyg-tabs-wrapper.

    For any non-TeamSite fallback decision markers, use normal DOM ancestry.
    """
    if is_explicit_teamsite_decision_wrapper(tag):
        return strict_decision_depth(tag)

    candidate_ids = {id(x) for x in candidates}
    depth = 1
    parent = tag.parent

    while isinstance(parent, Tag):
        if id(parent) in candidate_ids:
            depth += 1
        parent = parent.parent

    return depth

def analyze_html(url: str, html: str) -> AnalysisResult:
    soup = BeautifulSoup(html, "html.parser")

    title = ""
    if soup.title:
        title = normalize_spaces(soup.title.get_text(" ", strip=True))

    if not title:
        heading = soup.find(["h1", "h2"])
        if heading:
            title = normalize_spaces(heading.get_text(" ", strip=True))

    candidates = select_decision_containers(soup)

    evidence: List[DecisionEvidence] = []
    max_depth = 0

    for i, tag in enumerate(candidates, start=1):
        depth = decision_depth(tag, candidates)
        max_depth = max(max_depth, depth)

        if is_teamsite_tab_decision(tag):
            branch_labels = []
            seen_labels = set()
            for btn in tag.select("a.custom-wcm-wysiwyg-tabs-button, .td-tabs-label"):
                label = normalize_spaces(btn.get_text(" ", strip=True))
                if label and label.lower() not in seen_labels:
                    seen_labels.add(label.lower())
                    branch_labels.append(label)
            branch_text = " | ".join(branch_labels)
            text = (
                f"TeamSite tab decision; branches={len(branch_labels)}"
                + (f"; labels: {branch_text}" if branch_text else "")
            )
        else:
            text = normalize_spaces(tag.get_text(" ", strip=True))

        if len(text) > 220:
            text = text[:217] + "..."

        evidence.append(
            DecisionEvidence(
                index=i,
                depth=depth,
                tag=tag.name or "",
                element_id=tag.get("id", "") or "",
                classes=class_blob(tag),
                text_preview=text,
                dom_path=dom_path(tag),
            )
        )

    decision_count = len(evidence)
    nested_relationships = build_nested_relationships(candidates)
    nested_count = len(nested_relationships)

    if decision_count == 0:
        level = 1
        name = "Straightforward"
        reason = "No structural decision-point containers were detected in the Legacy page."
    elif max_depth <= 1:
        level = 2
        name = "Decision Points"
        reason = (
            f"{decision_count} decision-point structure(s) detected, "
            "with no decision point nested inside another."
        )
    else:
        level = 3
        name = "Nested Decision Points"
        reason = (
            f"Nested TeamSite decision structure detected: a decision wrapper exists "
            f"inside the tab-content area of another decision wrapper. "
            f"Maximum decision depth: {max_depth}."
        )

    return AnalysisResult(
        application=APP_NAME,
        version=VERSION,
        source_url=url,
        docid=get_docid(url),
        locale=get_locale(url),
        page_title=title,
        complexity_level=level,
        complexity_name=name,
        decision_points=decision_count,
        nested_decision_points=nested_count,
        maximum_decision_depth=max_depth,
        reason=reason,
        evidence=evidence,
        nested_relationships=nested_relationships,
    )


def write_outputs(result: AnalysisResult, html: str) -> Tuple[Path, Path, Path, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    base = safe_filename(result.docid)
    json_path = OUTPUT_DIR / f"{base}_complexity.json"
    csv_path = OUTPUT_DIR / f"{base}_complexity.csv"
    evidence_path = OUTPUT_DIR / f"{base}_evidence.html"
    source_path = OUTPUT_DIR / f"{base}_source.html"

    payload = asdict(result)
    json_path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "DOCID",
            "Locale",
            "Page Title",
            "Legacy URL",
            "Complexity Level",
            "Complexity Name",
            "Decision Points",
            "Nested Decision Points",
            "Maximum Decision Depth",
            "Reason",
        ])
        writer.writerow([
            result.docid,
            result.locale,
            result.page_title,
            result.source_url,
            result.complexity_level,
            result.complexity_name,
            result.decision_points,
            result.nested_decision_points,
            result.maximum_decision_depth,
            result.reason,
        ])

    rows = []
    for e in result.evidence:
        rows.append(
            "<tr>"
            f"<td>{e.index}</td>"
            f"<td>{e.depth}</td>"
            f"<td>{escape_html(e.tag)}</td>"
            f"<td>{escape_html(e.element_id)}</td>"
            f"<td>{escape_html(e.classes)}</td>"
            f"<td>{escape_html(e.dom_path)}</td>"
            f"<td>{escape_html(e.text_preview)}</td>"
            "</tr>"
        )

    relationship_rows_list = []
    for idx, rel in enumerate(result.nested_relationships, start=1):
        relationship_rows_list.append(
            "<tr>"
            f"<td>{idx}</td>"
            f"<td>{escape_html(' | '.join(rel.get('outer_branches', [])))}</td>"
            f"<td>{escape_html(' | '.join(rel.get('inner_branches', [])))}</td>"
            f"<td>{escape_html(rel.get('outer_path', ''))}</td>"
            f"<td>{escape_html(rel.get('inner_path', ''))}</td>"
            "</tr>"
        )

    relationship_rows = (
        "".join(relationship_rows_list)
        if relationship_rows_list
        else '<tr><td colspan="5">No nested parent-child decision relationship detected.</td></tr>'
    )

    evidence_html = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>{escape_html(result.docid)} - Procedure Complexity</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 28px; color: #222; }}
.summary {{ border: 1px solid #bbb; padding: 16px; border-radius: 8px; }}
.level {{ font-size: 28px; font-weight: 700; margin: 8px 0; }}
table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; vertical-align: top; }}
th {{ background: #f3f3f3; }}
code {{ white-space: pre-wrap; word-break: break-word; }}
</style>
</head>
<body>
<h1>TeamSite Procedure Complexity Classification</h1>
<div class="summary">
<div><strong>DOCID:</strong> {escape_html(result.docid)}</div>
<div><strong>Page:</strong> {escape_html(result.page_title)}</div>
<div><strong>URL:</strong> {escape_html(result.source_url)}</div>
<div class="level">Level {result.complexity_level} - {escape_html(result.complexity_name)}</div>
<div><strong>Decision points:</strong> {result.decision_points}</div>
<div><strong>Nested decision points:</strong> {result.nested_decision_points}</div>
<div><strong>Maximum decision depth:</strong> {result.maximum_decision_depth}</div>
<div><strong>Reason:</strong> {escape_html(result.reason)}</div>
</div>

<h2>Detected decision-point evidence</h2>
<table>
<thead>
<tr>
<th>#</th><th>Depth</th><th>Tag</th><th>ID</th><th>Classes</th>
<th>DOM Path</th><th>Text Preview</th>
</tr>
</thead>
<tbody>
{''.join(rows) if rows else '<tr><td colspan="7">No decision-point structure detected.</td></tr>'}
</tbody>
</table>

<h2>Nested decision relationships</h2>
<table>
<thead>
<tr>
<th>#</th><th>Outer Decision Branches</th><th>Nested Decision Branches</th>
<th>Outer DOM Path</th><th>Inner DOM Path</th>
</tr>
</thead>
<tbody>
{relationship_rows}
</tbody>
</table>
</body>
</html>
"""
    evidence_path.write_text(evidence_html, encoding="utf-8")
    source_path.write_text(html, encoding="utf-8")

    return json_path, csv_path, evidence_path, source_path


def escape_html(value: str) -> str:
    return (
        str(value or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def print_result(result: AnalysisResult) -> None:
    print()
    print("=" * 72)
    print("PROCEDURE COMPLEXITY RESULT")
    print("=" * 72)
    print(f"DOCID                  : {result.docid}")
    print(f"Page title             : {result.page_title or '(not found)'}")
    print(f"Complexity             : LEVEL {result.complexity_level} - {result.complexity_name}")
    print(f"Decision points        : {result.decision_points}")
    print(f"Nested decision points : {result.nested_decision_points}")
    print(f"Maximum decision depth : {result.maximum_decision_depth}")
    print(f"Reason                  : {result.reason}")

    if result.nested_relationships:
        print()
        print("Nested decision relationships:")
        for idx, rel in enumerate(result.nested_relationships, start=1):
            outer = " | ".join(rel.get("outer_branches", [])) or "(outer branches not labeled)"
            inner = " | ".join(rel.get("inner_branches", [])) or "(inner branches not labeled)"
            print(f"  {idx}. OUTER : {outer}")
            print(f"     NESTED: {inner}")

    print("=" * 72)


def main() -> int:
    print(f"{APP_NAME} v{VERSION}")
    print()
    print("Paste the LEGACY / TeamSite procedure URL.")
    print("Example:")
    print("https://w3.td.com/td/intranet/legacy/content?DOCID=...&LOCALE=en_ca")
    print()

    url = input("Legacy URL: ").strip()

    if not url:
        print("ERROR: No URL entered.")
        return 1

    if not re.match(r"^https?://", url, re.I):
        print("ERROR: Please enter the full http:// or https:// URL.")
        return 1

    try:
        print("\nReading Legacy page...")
        html = fetch_page(url)

        print("Analyzing TeamSite procedure structure...")
        result = analyze_html(url, html)

        print_result(result)

        json_path, csv_path, evidence_path, source_path = write_outputs(result, html)

        print("\nOutput files:")
        print(f"  JSON     : {json_path}")
        print(f"  CSV      : {csv_path}")
        print(f"  Evidence : {evidence_path}")
        print(f"  Source   : {source_path}")

        print("\nClassification rules:")
        print("  Level 1 = no decision-point structure")
        print("  Level 2 = decision point(s), no nesting")
        print("  Level 3 = decision point nested within another decision point")
        print()
        return 0

    except requests.exceptions.SSLError as exc:
        print("\nERROR: SSL certificate validation failed.")
        print("SSL bypass is allowed only for the configured internal TeamSite host.")
        print(str(exc))
        return 2

    except requests.exceptions.RequestException as exc:
        print("\nERROR: Could not read the Legacy URL.")
        print(str(exc))
        return 2

    except Exception as exc:
        print("\nERROR:")
        print(str(exc))
        return 3


if __name__ == "__main__":
    sys.exit(main())
