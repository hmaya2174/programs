#!/usr/bin/env python3
"""
SharePoint API Reader Baseline v1
===============================================

Purpose
-------
Analyze up to five NORMAL SharePoint page URLs and produce one consolidated
HTML diagnostic report focused on CanvasContent1 structure.

Input
-----
Normal SharePoint page URLs, for example:

https://myteam.td.com/sites/tdcentral-dept-wealth/SitePages/Authenticate-Client.aspx

The program automatically builds the REST endpoint:

https://myteam.td.com/sites/tdcentral-dept-wealth/_api/web/
GetFileByServerRelativeUrl('/sites/tdcentral-dept-wealth/SitePages/Authenticate-Client.aspx')/
ListItemAllFields?$select=Title,CanvasContent1,LayoutWebpartsContent

Retrieval approach
------------------
This version is designed for a managed Windows environment where:
- the REST URL works in the user's normal authenticated Microsoft Edge session,
- direct Python requests may return HTTP 401,
- Selenium/remote debugging may be blocked.

For each sample the program:
1. Converts the normal SharePoint page URL into the exact REST endpoint.
2. Opens ONLY the generated REST URL through Windows in the user's normal browser session.
3. Brings Edge to the foreground.
4. Sends Ctrl+A / Ctrl+C.
5. Reads the visible REST XML/JSON from the Windows clipboard.
6. Extracts and decodes CanvasContent1.
7. Analyzes structural patterns.
8. Compares all samples in one HTML report.

Output
------
One HTML file:
SharePoint_API_Reader_Baseline_Report_YYYYMMDD_HHMMSS.html

Dependencies
------------
pip install beautifulsoup4

No Selenium is required.
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
import urllib.parse
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

from bs4 import BeautifulSoup, Tag


APP_NAME = "SharePoint API Reader Baseline"
VERSION = "1.0"
MAX_SAMPLES = 5
OUTPUT_PREFIX = "SharePoint_CanvasContent1_Structure_Report"

PAGE_WARMUP_SECONDS = 4
REST_LOAD_SECONDS = 5
CLIPBOARD_WAIT_SECONDS = 2
MAX_RAW_CHARS = 2_000_000
MAX_MATRIX_ROWS = 1000


# ---------------------------------------------------------------------------
# DATA MODELS
# ---------------------------------------------------------------------------

@dataclass
class CanvasAnalysis:
    raw_length: int = 0
    decoded_length: int = 0
    text_length: int = 0
    max_dom_depth: int = 0

    tag_counts: Dict[str, int] = field(default_factory=dict)
    classes: Set[str] = field(default_factory=set)
    ids: Set[str] = field(default_factory=set)
    data_attributes: Set[str] = field(default_factory=set)
    all_attributes: Set[str] = field(default_factory=set)

    canvas_control_count: int = 0
    webpart_count: int = 0
    text_control_count: int = 0

    section_indexes: Set[str] = field(default_factory=set)
    zone_indexes: Set[str] = field(default_factory=set)
    control_types: Set[str] = field(default_factory=set)

    webpart_titles: Set[str] = field(default_factory=set)
    webpart_ids: Set[str] = field(default_factory=set)

    structural_signals: Set[str] = field(default_factory=set)
    nested_signals: Set[str] = field(default_factory=set)

    headings: List[str] = field(default_factory=list)
    links: int = 0
    images: int = 0
    tables: int = 0
    lists: int = 0

    text_preview: str = ""


@dataclass
class Sample:
    number: int
    page_url: str
    rest_url: str = ""

    ok: bool = False
    response_kind: str = "Unknown"
    error: str = ""

    title: str = ""
    canvas_found: bool = False
    layout_found: bool = False

    raw_response: str = ""
    canvas_raw: str = ""
    canvas_decoded: str = ""
    layout_raw: str = ""

    analysis: CanvasAnalysis = field(default_factory=CanvasAnalysis)


# ---------------------------------------------------------------------------
# BASIC HELPERS
# ---------------------------------------------------------------------------

def esc(value: Any) -> str:
    return html_lib.escape("" if value is None else str(value))


def norm(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def short(value: str, limit: int = 300) -> str:
    value = norm(value)
    return value if len(value) <= limit else value[: limit - 3] + "..."


def decode_html_entities(value: str) -> str:
    """
    Decode repeatedly because Atom/XML payloads can contain nested entity
    encoding such as &amp;lt; or &amp;quot;.
    """
    current = value or ""

    for _ in range(5):
        decoded = html_lib.unescape(current)

        if decoded == current:
            break

        current = decoded

    return current


# ---------------------------------------------------------------------------
# SHAREPOINT URL -> REST URL
# ---------------------------------------------------------------------------

def build_rest_url(page_url: str) -> str:
    parsed = urllib.parse.urlparse(page_url)

    if not parsed.scheme or not parsed.netloc:
        raise ValueError("Please enter a full SharePoint URL beginning with https://")

    path = parsed.path

    sitepages_marker = "/SitePages/"
    marker_index = path.lower().find(sitepages_marker.lower())

    if marker_index < 0:
        raise ValueError(
            "The URL does not appear to be a SharePoint SitePages URL. "
            "Expected a path containing /SitePages/."
        )

    site_path = path[:marker_index]
    server_relative_page = path

    # OData string literal escaping.
    server_relative_odata = server_relative_page.replace("'", "''")

    rest_path = (
        f"{site_path}/_api/web/"
        f"GetFileByServerRelativeUrl('{server_relative_odata}')/"
        "ListItemAllFields"
    )

    query = urllib.parse.urlencode(
        {
            "$select": "Title,CanvasContent1,LayoutWebpartsContent"
        },
        safe="$,"
    )

    return urllib.parse.urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            rest_path,
            "",
            query,
            "",
        )
    )


# ---------------------------------------------------------------------------
# WINDOWS / EDGE / CLIPBOARD
# ---------------------------------------------------------------------------

def find_edge_executable() -> str:
    candidates = [
        os.path.join(
            os.environ.get("PROGRAMFILES(X86)", ""),
            "Microsoft",
            "Edge",
            "Application",
            "msedge.exe",
        ),
        os.path.join(
            os.environ.get("PROGRAMFILES", ""),
            "Microsoft",
            "Edge",
            "Application",
            "msedge.exe",
        ),
        os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Microsoft",
            "Edge",
            "Application",
            "msedge.exe",
        ),
    ]

    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate

    raise RuntimeError("Microsoft Edge could not be found in the standard locations.")


def open_in_edge(url: str) -> None:
    """
    Open the URL through Windows itself, using the user's normal browser session.

    This intentionally does NOT start msedge.exe directly. In the TD-managed
    environment, opening the URL through Windows is closer to manually pasting
    the same URL into the already-authenticated browser.
    """
    if os.name != "nt":
        raise RuntimeError("Native browser opening is supported on Windows only.")

    os.startfile(url)


def run_powershell(script: str, timeout: int = 30) -> subprocess.CompletedProcess:
    return subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def clear_clipboard() -> None:
    run_powershell("Set-Clipboard -Value ''")


def get_clipboard() -> str:
    result = run_powershell("Get-Clipboard -Raw", timeout=30)

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

Start-Sleep -Milliseconds 600
$wshell.SendKeys("^a")
Start-Sleep -Milliseconds 350
$wshell.SendKeys("^c")
Start-Sleep -Milliseconds 900
"""

    result = run_powershell(ps, timeout=30)

    if result.returncode != 0:
        raise RuntimeError(
            "Could not activate Edge and copy the REST page. "
            + (result.stderr or "").strip()
        )


# ---------------------------------------------------------------------------
# REST RESPONSE PARSING
# ---------------------------------------------------------------------------

def parse_json_response(raw: str) -> Tuple[str, str, str]:
    """
    Returns: title, canvas, layout
    """
    data = json.loads(raw)

    def find_key(obj: Any, target: str) -> Optional[Any]:
        if isinstance(obj, dict):
            for key, value in obj.items():
                if str(key).lower() == target.lower():
                    return value

            for value in obj.values():
                found = find_key(value, target)

                if found is not None:
                    return found

        elif isinstance(obj, list):
            for value in obj:
                found = find_key(value, target)

                if found is not None:
                    return found

        return None

    title = find_key(data, "Title")
    canvas = find_key(data, "CanvasContent1")
    layout = find_key(data, "LayoutWebpartsContent")

    return (
        "" if title is None else str(title),
        "" if canvas is None else str(canvas),
        "" if layout is None else str(layout),
    )


def parse_xml_response(raw: str) -> Tuple[str, str, str]:
    """
    Parse SharePoint Atom XML.

    The REST response observed in Edge is typically:
      <entry ...>
        ...
        <content type="application/xml">
          <m:properties>
            <d:Title>...</d:Title>
            <d:CanvasContent1>...</d:CanvasContent1>
            <d:LayoutWebpartsContent>...</d:LayoutWebpartsContent>
          </m:properties>
        </content>
      </entry>
    """
    title = ""
    canvas = ""
    layout = ""

    root = ET.fromstring(raw)

    for element in root.iter():
        local_name = element.tag.split("}")[-1].lower()

        if local_name == "title":
            # Avoid the Atom entry title if it is not the SharePoint Title field.
            parent_text = element.text or ""

            if parent_text and not title:
                title = parent_text

        elif local_name == "canvascontent1":
            canvas = element.text or ""

        elif local_name == "layoutwebpartscontent":
            layout = element.text or ""

    # BeautifulSoup fallback is helpful when namespace parsing is unusual.
    if not canvas or not layout:
        soup = BeautifulSoup(raw, "xml")

        if not canvas:
            node = soup.find(
                lambda tag: tag.name
                and tag.name.split(":")[-1].lower() == "canvascontent1"
            )

            if node:
                canvas = node.get_text("", strip=False)

        if not layout:
            node = soup.find(
                lambda tag: tag.name
                and tag.name.split(":")[-1].lower() == "layoutwebpartscontent"
            )

            if node:
                layout = node.get_text("", strip=False)

        # Prefer d:Title / property title when available.
        title_nodes = soup.find_all(
            lambda tag: tag.name
            and tag.name.split(":")[-1].lower() == "title"
        )

        for node in reversed(title_nodes):
            value = norm(node.get_text(" ", strip=True))

            if value:
                title = value
                break

    return title, canvas, layout


def detect_and_parse_response(raw: str) -> Tuple[str, str, str, str]:
    stripped = (raw or "").strip()

    if not stripped:
        raise ValueError("The captured REST response was empty.")

    # JSON
    try:
        title, canvas, layout = parse_json_response(stripped)
        return "JSON", title, canvas, layout
    except Exception:
        pass

    # XML
    if (
        stripped.startswith("<?xml")
        or stripped.startswith("<entry")
        or stripped.startswith("<feed")
    ):
        title, canvas, layout = parse_xml_response(stripped)
        return "XML", title, canvas, layout

    # Sometimes Edge's XML viewer may omit the XML declaration in copied text.
    if "<entry" in stripped[:1000] and "CanvasContent1" in stripped:
        start = stripped.find("<entry")

        try:
            candidate = stripped[start:]
            title, canvas, layout = parse_xml_response(candidate)
            return "XML", title, canvas, layout
        except Exception:
            pass

    raise ValueError(
        "The captured browser content is not recognized as SharePoint REST JSON/XML. "
        f"Preview: {short(stripped, 400)}"
    )


# ---------------------------------------------------------------------------
# CANVASCONTENT1 ANALYSIS
# ---------------------------------------------------------------------------

def dom_depth(tag: Tag) -> int:
    depth = 0
    parent = tag.parent

    while isinstance(parent, Tag):
        depth += 1
        parent = parent.parent

    return depth


def parse_json_attribute(value: str) -> Optional[Any]:
    if not value:
        return None

    decoded = decode_html_entities(str(value))

    try:
        return json.loads(decoded)
    except Exception:
        return None


def walk_any_json(obj: Any) -> Iterable[Tuple[str, Any]]:
    if isinstance(obj, dict):
        for key, value in obj.items():
            yield str(key), value
            yield from walk_any_json(value)

    elif isinstance(obj, list):
        for value in obj:
            yield from walk_any_json(value)


def analyze_canvas(canvas_raw: str) -> Tuple[str, CanvasAnalysis]:
    decoded = decode_html_entities(canvas_raw)

    soup = BeautifulSoup(decoded, "html.parser")

    analysis = CanvasAnalysis(
        raw_length=len(canvas_raw or ""),
        decoded_length=len(decoded),
    )

    tag_counts = Counter()

    for tag in soup.find_all(True):
        name = (tag.name or "").lower()
        tag_counts[name] += 1
        analysis.max_dom_depth = max(
            analysis.max_dom_depth,
            dom_depth(tag),
        )

        for attr_name, attr_value in tag.attrs.items():
            attr_name_low = str(attr_name).lower()

            analysis.all_attributes.add(attr_name_low)

            if attr_name_low.startswith("data-"):
                analysis.data_attributes.add(attr_name_low)

            if attr_name_low == "id":
                analysis.ids.add(str(attr_value))

            if attr_name_low == "class":
                values = attr_value if isinstance(attr_value, list) else [attr_value]

                for css_class in values:
                    css_class = str(css_class)

                    if css_class:
                        analysis.classes.add(css_class)

                        low = css_class.lower()

                        if any(
                            token in low
                            for token in (
                                "tab",
                                "accordion",
                                "decision",
                                "branch",
                                "step",
                                "expand",
                                "collapse",
                                "section",
                                "canvas",
                                "webpart",
                                "control",
                            )
                        ):
                            analysis.structural_signals.add(
                                "class:" + css_class
                            )

        if tag.get("data-sp-canvascontrol") is not None:
            analysis.canvas_control_count += 1

        control_data = tag.get("data-sp-controldata")

        if control_data:
            parsed = parse_json_attribute(control_data)

            if isinstance(parsed, dict):
                position = parsed.get("position", {})

                if isinstance(position, dict):
                    if "sectionIndex" in position:
                        analysis.section_indexes.add(
                            str(position.get("sectionIndex"))
                        )

                    if "zoneIndex" in position:
                        analysis.zone_indexes.add(
                            str(position.get("zoneIndex"))
                        )

                    if "controlIndex" in position:
                        analysis.structural_signals.add(
                            "controlIndex:" + str(position.get("controlIndex"))
                        )

                if "controlType" in parsed:
                    analysis.control_types.add(
                        str(parsed.get("controlType"))
                    )

        webpart_data = tag.get("data-sp-webpartdata")

        if webpart_data:
            analysis.webpart_count += 1

            parsed = parse_json_attribute(webpart_data)

            if parsed is not None:
                for key, value in walk_any_json(parsed):
                    key_low = key.lower()

                    if key_low in {"id", "instanceid", "webpartid"}:
                        if value:
                            analysis.webpart_ids.add(str(value))

                    if key_low == "title":
                        if value:
                            analysis.webpart_titles.add(str(value))

                    if key_low in {
                        "controltype",
                        "layouttype",
                        "webparttype",
                        "componentname",
                        "serverprocessedcontent",
                    }:
                        analysis.structural_signals.add(
                            f"webpart:{key}={short(str(value), 120)}"
                        )

        # Likely text controls / rich text.
        if name in {"p", "div", "span"}:
            if tag.get("data-sp-feature-tag") or tag.get("data-sp-rte"):
                analysis.text_control_count += 1

        if name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            heading = norm(tag.get_text(" ", strip=True))

            if heading:
                analysis.headings.append(heading[:250])

    # Nested structural relationships.
    interesting_tokens = (
        "tab",
        "accordion",
        "decision",
        "branch",
        "step",
        "section",
        "canvas",
        "webpart",
        "control",
    )

    for tag in soup.find_all(True):
        child_classes = tag.get("class", [])

        if isinstance(child_classes, str):
            child_classes = [child_classes]

        child_blob = " ".join(str(x).lower() for x in child_classes)

        child_signal = any(
            token in child_blob
            for token in interesting_tokens
        )

        if not child_signal and not tag.get("data-sp-webpartdata"):
            continue

        parent = tag.parent

        while isinstance(parent, Tag):
            parent_classes = parent.get("class", [])

            if isinstance(parent_classes, str):
                parent_classes = [parent_classes]

            parent_blob = " ".join(
                str(x).lower()
                for x in parent_classes
            )

            parent_signal = (
                any(
                    token in parent_blob
                    for token in interesting_tokens
                )
                or parent.get("data-sp-webpartdata") is not None
                or parent.get("data-sp-canvascontrol") is not None
            )

            if parent_signal:
                relation = (
                    f"{short(parent_blob or parent.name, 100)}"
                    " -> "
                    f"{short(child_blob or tag.name, 100)}"
                )

                analysis.nested_signals.add(relation)
                break

            parent = parent.parent

    clean_text = norm(soup.get_text(" ", strip=True))

    analysis.text_length = len(clean_text)
    analysis.tag_counts = dict(tag_counts)

    analysis.links = len(soup.find_all("a"))
    analysis.images = len(soup.find_all("img"))
    analysis.tables = len(soup.find_all("table"))
    analysis.lists = len(soup.find_all(["ul", "ol"]))

    analysis.text_preview = clean_text[:7000]

    return decoded, analysis


# ---------------------------------------------------------------------------
# CAPTURE ONE SAMPLE
# ---------------------------------------------------------------------------

def capture_sample(number: int, page_url: str) -> Sample:
    sample = Sample(
        number=number,
        page_url=page_url,
    )

    try:
        sample.rest_url = build_rest_url(page_url)

        print(f"Sample {number}: generated REST endpoint:")
        print(f"  {sample.rest_url}")
        print()
        print(f"Sample {number}: opening REST endpoint in your normal authenticated browser...")
        clear_clipboard()
        open_in_edge(sample.rest_url)

        print()
        print("WAIT until the REST API response is fully visible in Edge.")
        input("Then return to this terminal and press Enter to capture it: ")

        print(f"Sample {number}: capturing REST response...")
        activate_edge_and_copy()

        time.sleep(CLIPBOARD_WAIT_SECONDS)

        raw = get_clipboard().strip()

        if not raw or len(raw) < 200:
            sample.error = (
                "REST endpoint opened successfully, but the API response was not "
                "captured. No analysis was performed for this sample."
            )
            print("  CAPTURE FAILED - no usable REST response was copied.")
            return sample

        sample.raw_response = raw[:MAX_RAW_CHARS]

        (
            sample.response_kind,
            sample.title,
            sample.canvas_raw,
            sample.layout_raw,
        ) = detect_and_parse_response(raw)

        sample.canvas_found = bool(sample.canvas_raw)
        sample.layout_found = bool(sample.layout_raw)

        if not sample.canvas_found:
            sample.error = (
                "REST response was captured successfully, but CanvasContent1 "
                "was not found or was empty."
            )
            return sample

        (
            sample.canvas_decoded,
            sample.analysis,
        ) = analyze_canvas(sample.canvas_raw)

        sample.ok = True
        return sample

    except Exception as exc:
        sample.error = str(exc)
        return sample


# ---------------------------------------------------------------------------
# COMPARISON
# ---------------------------------------------------------------------------

def sample_feature_sets(sample: Sample) -> Dict[str, Set[str]]:
    a = sample.analysis

    return {
        "CSS Class": set(a.classes),
        "HTML ID": set(a.ids),
        "data-* Attribute": set(a.data_attributes),
        "HTML Attribute": set(a.all_attributes),
        "HTML Tag": set(a.tag_counts.keys()),
        "Control Type": set(a.control_types),
        "Section Index": set(a.section_indexes),
        "Zone Index": set(a.zone_indexes),
        "Web Part Title": set(a.webpart_titles),
        "Web Part ID": set(a.webpart_ids),
        "Structural Signal": set(a.structural_signals),
        "Nested Signal": set(a.nested_signals),
    }


def build_comparison_rows(samples: List[Sample]) -> List[Dict[str, Any]]:
    sample_sets = [
        sample_feature_sets(sample)
        for sample in samples
    ]

    categories = sorted(
        {
            category
            for feature_set in sample_sets
            for category in feature_set.keys()
        }
    )

    rows: List[Dict[str, Any]] = []
    count = len(samples)

    for category in categories:
        universe: Set[str] = set()

        for feature_set in sample_sets:
            universe.update(
                feature_set.get(category, set())
            )

        for item in sorted(universe, key=str.lower):
            presence = [
                item in feature_set.get(category, set())
                for feature_set in sample_sets
            ]

            frequency = sum(presence)

            if frequency == count:
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

    rank = {
        "Consistent": 0,
        "Unique": 1,
        "Variant": 2,
    }

    return sorted(
        rows,
        key=lambda row: (
            rank[row["finding"]],
            row["category"],
            row["item"].lower(),
        ),
    )[:MAX_MATRIX_ROWS]


# ---------------------------------------------------------------------------
# HTML REPORT
# ---------------------------------------------------------------------------

def render_matrix(
    rows: List[Dict[str, Any]],
    sample_count: int,
    heading: str,
) -> str:
    if not rows:
        return (
            f"<h3>{esc(heading)}</h3>"
            "<p>No comparable structures detected.</p>"
        )

    sample_headers = "".join(
        f"<th>S{i}</th>"
        for i in range(1, sample_count + 1)
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
        <tbody>
          {''.join(body)}
        </tbody>
      </table>
    </div>
    """


def render_sample(sample: Sample) -> str:
    a = sample.analysis

    return f"""
    <section>

      <h2>Sample {sample.number}</h2>

      <table>

        <tr>
          <th>SharePoint Page URL</th>
          <td class="break">
            <code>{esc(sample.page_url)}</code>
          </td>
        </tr>

        <tr>
          <th>Generated REST URL</th>
          <td class="break">
            <code>{esc(sample.rest_url)}</code>
          </td>
        </tr>

        <tr>
          <th>Status</th>
          <td>{'Success' if sample.ok else 'Failed'}</td>
        </tr>

        <tr>
          <th>REST Response Type</th>
          <td>{esc(sample.response_kind)}</td>
        </tr>

        <tr>
          <th>Page Title</th>
          <td>{esc(sample.title or 'Not detected')}</td>
        </tr>

        <tr>
          <th>CanvasContent1</th>
          <td>{'Found' if sample.canvas_found else 'Not found'}</td>
        </tr>

        <tr>
          <th>LayoutWebpartsContent</th>
          <td>{'Found' if sample.layout_found else 'Not found'}</td>
        </tr>

        <tr>
          <th>Error</th>
          <td>{esc(sample.error or 'None')}</td>
        </tr>

      </table>

      <h3>CanvasContent1 Diagnostic Summary</h3>

      <table>

        <tr>
          <th>Raw CanvasContent1 length</th>
          <td>{a.raw_length:,}</td>
        </tr>

        <tr>
          <th>Decoded HTML length</th>
          <td>{a.decoded_length:,}</td>
        </tr>

        <tr>
          <th>Clean text length</th>
          <td>{a.text_length:,}</td>
        </tr>

        <tr>
          <th>Maximum DOM depth</th>
          <td>{a.max_dom_depth}</td>
        </tr>

        <tr>
          <th>Canvas controls</th>
          <td>{a.canvas_control_count}</td>
        </tr>

        <tr>
          <th>Web parts</th>
          <td>{a.webpart_count}</td>
        </tr>

        <tr>
          <th>Control types</th>
          <td><code>{esc(", ".join(sorted(a.control_types)))}</code></td>
        </tr>

        <tr>
          <th>Section indexes</th>
          <td><code>{esc(", ".join(sorted(a.section_indexes)))}</code></td>
        </tr>

        <tr>
          <th>Zone indexes</th>
          <td><code>{esc(", ".join(sorted(a.zone_indexes)))}</code></td>
        </tr>

        <tr>
          <th>Links / Images / Tables / Lists</th>
          <td>{a.links} / {a.images} / {a.tables} / {a.lists}</td>
        </tr>

      </table>

      <details>
        <summary>Structural Signals</summary>
        <div class="pad">
          <pre>{esc(chr(10).join(sorted(a.structural_signals)))}</pre>
        </div>
      </details>

      <details>
        <summary>Nested Signals</summary>
        <div class="pad">
          <pre>{esc(chr(10).join(sorted(a.nested_signals)))}</pre>
        </div>
      </details>

      <details>
        <summary>Clean Text Preview</summary>
        <div class="pad">
          <pre>{esc(a.text_preview)}</pre>
        </div>
      </details>

      <details>
        <summary>Decoded CanvasContent1</summary>
        <div class="pad">
          <pre>{esc(sample.canvas_decoded)}</pre>
        </div>
      </details>

      <details>
        <summary>Raw REST Response</summary>
        <div class="pad">
          <pre>{esc(sample.raw_response)}</pre>
        </div>
      </details>

    </section>
    """


def build_report(samples: List[Sample]) -> str:
    rows = build_comparison_rows(samples)

    consistent = [
        row
        for row in rows
        if row["finding"] == "Consistent"
    ]

    unique = [
        row
        for row in rows
        if row["finding"] == "Unique"
    ]

    variants = [
        row
        for row in rows
        if row["finding"] == "Variant"
    ]

    success_count = sum(
        1
        for sample in samples
        if sample.ok
    )

    cards = []

    for sample in samples:
        css_class = "ok" if sample.ok else "bad"

        cards.append(
            f"""
            <div class="card">
              <strong>Sample {sample.number}</strong><br>
              <span class="{css_class}">
                {'Success' if sample.ok else 'Failed'}
              </span><br>
              {esc(sample.response_kind)}<br>
              {esc(sample.title or 'Title not detected')}<br>
              Canvas: {'Yes' if sample.canvas_found else 'No'}<br>
              Controls: {sample.analysis.canvas_control_count}<br>
              Web Parts: {sample.analysis.webpart_count}
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

body {{
    font-family: Segoe UI, Arial, sans-serif;
    margin: 0;
    color: #1f2328;
    line-height: 1.45;
}}

header {{
    background: #f6f8fa;
    border-bottom: 1px solid #d0d7de;
    padding: 26px 32px;
}}

main {{
    max-width: 1500px;
    margin: auto;
    padding: 26px 32px 60px;
}}

h1 {{
    margin: 0;
}}

h2 {{
    margin-top: 38px;
    border-top: 1px solid #d0d7de;
    padding-top: 14px;
}}

.cards {{
    display: grid;
    grid-template-columns:
        repeat(auto-fit, minmax(220px, 1fr));
    gap: 10px;
    margin: 18px 0;
}}

.card {{
    border: 1px solid #d0d7de;
    border-radius: 8px;
    padding: 12px;
}}

.ok {{
    color: #1a7f37;
}}

.bad {{
    color: #cf222e;
}}

table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}}

th,
td {{
    border: 1px solid #d0d7de;
    padding: 7px;
    vertical-align: top;
    text-align: left;
}}

th {{
    background: #f6f8fa;
    position: sticky;
    top: 0;
}}

.center {{
    text-align: center;
}}

.consistent td {{
    background: #eef8f0;
}}

.unique td {{
    background: #fff8c5;
}}

.variant td {{
    background: #f6f8fa;
}}

.scroll {{
    overflow: auto;
    max-height: 650px;
    border: 1px solid #d0d7de;
}}

pre {{
    background: #0d1117;
    color: #e6edf3;
    padding: 14px;
    border-radius: 6px;
    white-space: pre-wrap;
    word-break: break-word;
    max-height: 700px;
    overflow: auto;
    font-size: 12px;
}}

code {{
    font-family: Consolas, monospace;
    word-break: break-word;
}}

details {{
    border: 1px solid #d0d7de;
    border-radius: 6px;
    margin: 10px 0;
}}

summary {{
    cursor: pointer;
    background: #f6f8fa;
    padding: 10px;
    font-weight: 600;
}}

.pad {{
    padding: 12px;
}}

.break {{
    word-break: break-all;
}}

</style>

</head>

<body>

<header>

<h1>{esc(APP_NAME)}</h1>

<div>
Version {esc(VERSION)}
· Generated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
</div>

</header>

<main>

<h2>Consolidated Diagnostic Summary</h2>

<p>
<strong>{success_count}/{len(samples)}</strong>
samples were successfully captured and CanvasContent1 was analyzed.
</p>

<p>
This diagnostic report compares CanvasContent1 structures across the supplied
SharePoint pages. It intentionally does not assign Level 1 / Level 2 / Level 3
procedure classifications yet.
</p>

<div class="cards">
{''.join(cards)}
</div>

<h2>Consolidated Structural Comparison Matrix</h2>

<p>
<strong>Consistent</strong> = found in every sample.
<strong>Variant</strong> = found in multiple but not all samples.
<strong>Unique</strong> = found in one sample only.
</p>

{render_matrix(
    rows,
    len(samples),
    "All Prioritized CanvasContent1 Attributes and Patterns"
)}

<h2>Consistent Attributes</h2>

{render_matrix(
    consistent,
    len(samples),
    "Present Across All Samples"
)}

<h2>Unique Variants</h2>

{render_matrix(
    unique,
    len(samples),
    "Present in One Sample Only"
)}

<h2>Shared Variants</h2>

{render_matrix(
    variants,
    len(samples),
    "Present in Multiple But Not All Samples"
)}

<h2>Individual Sample Evidence</h2>

{''.join(
    render_sample(sample)
    for sample in samples
)}

</main>

</body>

</html>
"""


# ---------------------------------------------------------------------------
# USER INPUT / MAIN
# ---------------------------------------------------------------------------

def prompt_urls() -> List[str]:
    print("=" * 78)
    print(f"{APP_NAME} v{VERSION}")
    print("=" * 78)
    print()
    print("Paste NORMAL SharePoint procedure page URLs.")
    print("Do NOT paste the REST API URL.")
    print("Press Enter to leave a sample slot blank.")
    print()

    urls: List[str] = []

    for index in range(1, MAX_SAMPLES + 1):
        urls.append(
            input(
                f"SharePoint Page URL {index}: "
            ).strip()
        )

    if not any(urls):
        print()
        print("ERROR: No SharePoint URLs were supplied.")
        return []

    return urls


def main() -> int:
    if os.name != "nt":
        print("ERROR: This program is designed for Windows.")
        return 2

    urls = prompt_urls()

    if not urls:
        return 1

    print()
    print("IMPORTANT")
    print("1. Open Microsoft Edge normally and confirm SharePoint is already working.")
    print("2. Leave that authenticated Edge window open.")
    print("3. The program will open ONLY the generated REST endpoints.")
    print("4. Do not use the keyboard/mouse while each sample is being captured.")
    print()

    input("Press Enter when Edge is open and authenticated: ")

    samples: List[Sample] = []

    print()
    print("Starting CanvasContent1 diagnostics...")
    print()

    for index, page_url in enumerate(
        urls,
        start=1,
    ):
        if not page_url:
            samples.append(
                Sample(
                    number=index,
                    page_url="",
                    error="No URL supplied.",
                )
            )

            print(f"Sample {index}: skipped")
            print()
            continue

        sample = capture_sample(
            index,
            page_url,
        )

        samples.append(sample)

        if sample.ok:
            print(
                f"  SUCCESS - {sample.response_kind}, "
                f"CanvasContent1 {sample.analysis.decoded_length:,} characters, "
                f"{sample.analysis.canvas_control_count} canvas controls, "
                f"{sample.analysis.webpart_count} web parts"
            )

        else:
            print(
                f"  FAILED - {sample.error}"
            )

        print()

    print(
        "Building consolidated CanvasContent1 HTML report..."
    )

    report = build_report(samples)

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    output = Path(
        f"{OUTPUT_PREFIX}_{timestamp}.html"
    )

    output.write_text(
        report,
        encoding="utf-8",
    )

    print()
    print("=" * 78)
    print("REPORT CREATED")
    print("=" * 78)
    print(output.resolve())
    print()
    print(
        "Upload this one HTML report for structural analysis "
        "and future Level 1/2/3 rule development."
    )
    print()

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())

    except KeyboardInterrupt:
        print()
        print("Cancelled.")
        sys.exit(130)

    except Exception:
        print()
        print("UNEXPECTED ERROR")
        traceback.print_exc()
        sys.exit(2)
