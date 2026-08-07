#!/usr/bin/env python3
"""
SharePoint REST Structure Analyzer v4
=====================================

Purpose
-------
Reusable diagnostic utility for SharePoint REST API responses.

INPUT:
    Up to 5 SharePoint REST API URLs.

RETRIEVAL:
    Opens the URLs through Microsoft Edge using the user's existing Edge profile
    so the same SharePoint/Microsoft authentication state can be reused.

OUTPUT:
    One self-contained HTML report comparing all samples.

The report identifies:
    - Consistent REST/JSON attributes across all samples
    - Shared variants
    - Unique variants
    - CanvasContent1 / LayoutWebpartsContent
    - Embedded HTML structures
    - CSS classes
    - IDs
    - data-* attributes
    - HTML tags
    - structural/nesting signals
    - raw REST responses

IMPORTANT
---------
1. Close all normal Microsoft Edge windows before running this program.
2. The program does NOT request or store usernames/passwords.
3. A REST response is considered successful only when actual JSON/XML/content
   is returned. Small Edge error/interstitial pages are NOT treated as success.

Dependencies
------------
pip install selenium beautifulsoup4
"""

from __future__ import annotations

import html as html_lib
import json
import os
import socket
import subprocess
import re
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
VERSION = "5.0 Remote Debug Attach"
MAX_SAMPLES = 5
OUTPUT_PREFIX = "SharePoint_REST_Structure_Analysis_Report"
MAX_MATRIX_ROWS = 800
MAX_RAW_CHARS = 2_000_000
MIN_VALID_TEXT_RESPONSE = 1000


# ---------------------------------------------------------------------------
# DATA MODELS
# ---------------------------------------------------------------------------

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
    browser: str = ""
    response_kind: str = "Unknown"
    page_title: str = ""
    raw_text: str = ""
    error: str = ""
    top_fields: Set[str] = field(default_factory=set)
    json_paths: Set[str] = field(default_factory=set)
    path_types: Dict[str, str] = field(default_factory=dict)
    html_analyses: List[HtmlAnalysis] = field(default_factory=list)
    interesting_fields: Dict[str, str] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# BASIC HELPERS
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# JSON STRUCTURE ANALYSIS
# ---------------------------------------------------------------------------

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
            "<div",
            "<p",
            "<section",
            "<article",
            "<table",
            "<span",
            "<a ",
            "data-sp-",
            "<img",
        )
    )


# ---------------------------------------------------------------------------
# HTML / CANVAS ANALYSIS
# ---------------------------------------------------------------------------

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
                            "web-part",
                            "control",
                        )
                    ):
                        signals.add("class:" + css_class)

        role = str(tag.attrs.get("role", "")).strip()

        if role:
            signals.add("role:" + role)

        if name in {
            "table",
            "ul",
            "ol",
            "section",
            "article",
            "details",
            "summary",
        }:
            signals.add("tag:" + name)

        if name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            heading = short(tag.get_text(" ", strip=True), 180)

            if heading:
                headings.append(heading)

    # Detect nested structural classes.
    tokens = (
        "tab",
        "accordion",
        "decision",
        "branch",
        "step",
        "section",
        "canvas",
        "webpart",
    )

    nested = Counter()

    for tag in soup.find_all(True):
        child_classes = tag.attrs.get("class", [])

        if isinstance(child_classes, str):
            child_classes = [child_classes]

        child_blob = " ".join(str(x).lower() for x in child_classes)

        if not any(token in child_blob for token in tokens):
            continue

        parent = tag.parent

        while isinstance(parent, Tag):
            parent_classes = parent.attrs.get("class", [])

            if isinstance(parent_classes, str):
                parent_classes = [parent_classes]

            parent_blob = " ".join(str(x).lower() for x in parent_classes)

            if any(token in parent_blob for token in tokens):
                nested[
                    f"{short(parent_blob, 80)} -> {short(child_blob, 80)}"
                ] += 1
                break

            parent = parent.parent

    for relationship, count in nested.items():
        signals.add(f"nested:{relationship} ({count})")

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
) -> Tuple[
    Set[str],
    Dict[str, str],
    List[HtmlAnalysis],
    Dict[str, str],
    str,
]:
    paths, types = flatten_json(data)

    analyses: List[HtmlAnalysis] = []
    interesting: Dict[str, str] = {}
    title = ""

    important_names = {
        "canvascontent1",
        "layoutwebpartscontent",
        "title",
        "pagelayouttype",
        "promotedstate",
        "description",
        "bannerimageurl",
        "webpartdata",
        "controltype",
        "id",
        "uniqueid",
        "fileleafref",
        "fileref",
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


# ---------------------------------------------------------------------------
# SELENIUM / EDGE
# ---------------------------------------------------------------------------

def import_selenium():
    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.common.by import By

        return webdriver, EdgeOptions, By

    except ImportError:
        print()
        print("ERROR: Selenium is not installed for this Python interpreter.")
        print()
        print("Install it using:")
        print("    python -m pip install selenium beautifulsoup4")
        print()
        raise


def detect_edge_profile() -> Tuple[str, str]:
    """
    Detect Edge user-data root and the most recently used profile.

    Edge stores the last used profile in:
        %LOCALAPPDATA%\\Microsoft\\Edge\\User Data\\Local State
    """
    local_appdata = os.environ.get("LOCALAPPDATA", "")

    user_data = os.path.join(
        local_appdata,
        "Microsoft",
        "Edge",
        "User Data",
    )

    profile = "Default"

    local_state = os.path.join(user_data, "Local State")

    try:
        if os.path.exists(local_state):
            with open(local_state, "r", encoding="utf-8") as handle:
                data = json.load(handle)

            profile_info = data.get("profile", {})

            last_used = profile_info.get("last_used")

            if last_used:
                profile = last_used

    except Exception:
        pass

    return user_data, profile


def find_edge_executable() -> str:
    """
    Locate Microsoft Edge in common Windows installation locations.
    """
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

    raise RuntimeError(
        "Microsoft Edge executable could not be found in the standard locations."
    )


def port_is_open(host: str, port: int) -> bool:
    try:
        with socket.create_connection((host, port), timeout=0.5):
            return True
    except OSError:
        return False


def start_browser():
    """
    Launch Edge explicitly with the user's existing profile and a local
    remote-debugging port, then attach Selenium to that already-open browser.

    This is more reliable than asking Selenium itself to create Edge with an
    existing corporate profile.
    """
    webdriver, EdgeOptions, By = import_selenium()

    user_data, profile = detect_edge_profile()
    edge_exe = find_edge_executable()
    debug_port = 9222

    print()
    print("Authentication mode : Existing Microsoft Edge profile")
    print("Edge executable     :", edge_exe)
    print("Edge user data      :", user_data)
    print("Edge profile        :", profile)
    print("Debug port          :", debug_port)
    print()
    print("IMPORTANT")
    print("Close ALL Microsoft Edge windows before continuing.")
    print()

    input("Press Enter after Edge is fully closed: ")

    if port_is_open("127.0.0.1", debug_port):
        raise RuntimeError(
            f"Port {debug_port} is already in use. Close any existing Edge "
            "remote-debugging session and run the program again."
        )

    cmd = [
        edge_exe,
        f"--remote-debugging-port={debug_port}",
        f"--user-data-dir={user_data}",
        f"--profile-directory={profile}",
        "--start-maximized",
        "--disable-notifications",
        "--disable-popup-blocking",
        "about:blank",
    ]

    print()
    print("Launching Microsoft Edge...")
    print("Waiting for Edge to become available...")

    try:
        edge_process = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        raise RuntimeError(
            f"Could not launch Microsoft Edge directly. Technical detail: {exc}"
        )

    ready = False
    for _ in range(40):
        if port_is_open("127.0.0.1", debug_port):
            ready = True
            break
        time.sleep(0.5)

    if not ready:
        try:
            edge_process.terminate()
        except Exception:
            pass

        raise RuntimeError(
            "Edge was launched but the remote-debugging connection never became "
            "available. Your corporate Edge policy may block remote debugging, "
            "or the selected Edge profile may still be locked by a background "
            "Edge process."
        )

    print("Edge opened successfully.")
    print("Attaching Selenium to the existing Edge window...")

    options = EdgeOptions()
    options.add_experimental_option(
        "debuggerAddress",
        f"127.0.0.1:{debug_port}"
    )

    try:
        driver = webdriver.Edge(options=options)
    except Exception as exc:
        try:
            edge_process.terminate()
        except Exception:
            pass

        raise RuntimeError(
            "Edge opened, but Selenium could not attach to it. "
            f"Technical detail: {exc}"
        )

    print("Selenium attached successfully.")
    return driver, f"Microsoft Edge - {profile}", By


# ---------------------------------------------------------------------------
# BROWSER RESPONSE EXTRACTION
# ---------------------------------------------------------------------------

def get_body_text(driver, By) -> str:
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        return body.text or ""
    except Exception:
        return ""


def get_pre_text(driver, By) -> str:
    """
    Browser JSON/XML viewers often render the response inside <pre>.
    """
    try:
        elements = driver.find_elements(By.TAG_NAME, "pre")

        candidates = [
            element.text
            for element in elements
            if (element.text or "").strip()
        ]

        if candidates:
            return max(candidates, key=len)

    except Exception:
        pass

    return ""


def get_source(driver) -> str:
    try:
        return driver.page_source or ""
    except Exception:
        return ""


def looks_like_browser_error_page(text: str, source: str) -> bool:
    blob = (text + "\n" + source[:50000]).lower()

    edge_error_markers = (
        "edge-logo",
        "edge-branding-text",
        "icon-elixir-page-error",
        "error-code",
        "interstitial-wrapper",
        "this page isn't working",
        "this page isn’t working",
        "can't reach this page",
        "can’t reach this page",
    )

    return any(marker in blob for marker in edge_error_markers)


def looks_like_signin(text: str, source: str) -> bool:
    blob = (text + "\n" + source[:30000]).lower()

    return (
        any(marker in blob for marker in ("sign in", "signin", "log in", "login"))
        and any(marker in blob for marker in ("microsoft", "account", "password"))
    )


def parse_json_from_browser_text(raw: str) -> Tuple[Optional[Any], str]:
    """
    Attempt to parse browser-visible REST content as JSON.
    """
    raw = (raw or "").strip()

    if not raw:
        return None, raw

    try:
        return json.loads(raw), raw

    except Exception:
        pass

    # Browser viewers sometimes add a short prefix before the JSON.
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


def fetch_via_browser(
    driver,
    browser_name: str,
    By,
    sample_number: int,
    url: str,
) -> Sample:

    result = Sample(
        number=sample_number,
        url=url,
        browser=browser_name,
    )

    try:
        driver.get(url)

        time.sleep(2.5)

        pre_text = get_pre_text(driver, By)
        body_text = get_body_text(driver, By)
        source = get_source(driver)

        raw = (pre_text or body_text or "").strip()

        # Handle authentication page if necessary.
        if looks_like_signin(raw, source):
            print()
            print(f"Sample {sample_number}: SharePoint sign-in page detected.")
            print("Complete your normal corporate sign-in in the Edge window.")
            print()

            input(
                "When the REST response is visible in Edge, "
                "press Enter here to continue: "
            )

            time.sleep(1)

            pre_text = get_pre_text(driver, By)
            body_text = get_body_text(driver, By)
            source = get_source(driver)

            raw = (pre_text or body_text or "").strip()

        # Reject Edge error/interstitial pages.
        if looks_like_browser_error_page(raw, source):
            result.response_kind = "Browser Error Page"
            result.raw_text = raw[:MAX_RAW_CHARS]
            result.error = (
                "Microsoft Edge returned an error/interstitial page instead "
                "of the SharePoint REST response."
            )
            return result

        if not raw:
            result.error = "The browser returned an empty response."
            return result

        parsed_json, cleaned_raw = parse_json_from_browser_text(raw)

        if parsed_json is not None:
            result.response_kind = "JSON"
            result.raw_text = cleaned_raw[:MAX_RAW_CHARS]

            if isinstance(parsed_json, dict):
                result.top_fields = {
                    str(key)
                    for key in parsed_json.keys()
                }

            (
                paths,
                types,
                analyses,
                interesting,
                title,
            ) = extract_json(parsed_json)

            result.json_paths = paths
            result.path_types = types
            result.html_analyses = analyses
            result.interesting_fields = interesting
            result.page_title = title
            result.ok = True

            return result

        low = raw.lstrip().lower()

        if (
            low.startswith("<?xml")
            or low.startswith("<feed")
            or low.startswith("<entry")
        ):
            result.response_kind = "XML"
            result.raw_text = raw[:MAX_RAW_CHARS]
            result.html_analyses = [
                analyze_html("$xml_response", raw)
            ]
            result.ok = True
            return result

        # Small HTML/Text responses are almost always browser/interstitial/auth
        # content and should not be considered valid REST payloads.
        if len(raw) < MIN_VALID_TEXT_RESPONSE:
            result.response_kind = "HTML/Text"
            result.raw_text = raw[:MAX_RAW_CHARS]
            result.error = (
                "The browser opened the URL but did not return usable REST data. "
                f"Only {len(raw)} characters were returned. "
                f"Preview: {short(raw, 300)}"
            )
            return result

        result.response_kind = "HTML/Text"
        result.raw_text = raw[:MAX_RAW_CHARS]
        result.html_analyses = [
            analyze_html("$browser_response", source or raw)
        ]

        soup = BeautifulSoup(source or raw, "html.parser")

        if soup.title:
            result.page_title = norm(
                soup.title.get_text(" ", strip=True)
            )

        result.ok = True

        return result

    except Exception as exc:
        result.error = str(exc)
        return result


# ---------------------------------------------------------------------------
# CROSS-SAMPLE COMPARISON
# ---------------------------------------------------------------------------

def merged_set(sample: Sample, attribute: str) -> Set[str]:
    output: Set[str] = set()

    for analysis in sample.html_analyses:
        output.update(
            getattr(analysis, attribute, set())
        )

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
        presence = [
            item in values
            for values in sets
        ]

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


def matrix(
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


# ---------------------------------------------------------------------------
# SAMPLE REPORT DETAIL
# ---------------------------------------------------------------------------

def sample_details(sample: Sample) -> str:
    interesting_rows = "".join(
        f"<tr>"
        f"<td><code>{esc(path)}</code></td>"
        f"<td>{esc(value)}</td>"
        f"</tr>"
        for path, value in sorted(
            sample.interesting_fields.items()
        )
    )

    if not interesting_rows:
        interesting_rows = (
            "<tr><td colspan='2'>"
            "No targeted fields detected."
            "</td></tr>"
        )

    html_blocks = []

    for index, analysis in enumerate(
        sample.html_analyses,
        start=1,
    ):
        html_blocks.append(
            f"""
            <details>
              <summary>
                HTML / Canvas Fragment {index}:
                <code>{esc(analysis.source_field)}</code>
              </summary>

              <div class="pad">
                <table>
                  <tr>
                    <th>HTML length</th>
                    <td>{analysis.length:,}</td>
                  </tr>

                  <tr>
                    <th>Maximum DOM depth</th>
                    <td>{analysis.max_depth}</td>
                  </tr>

                  <tr>
                    <th>Links / Images / Tables / Lists</th>
                    <td>
                      {analysis.link_count} /
                      {analysis.image_count} /
                      {analysis.table_count} /
                      {analysis.list_count}
                    </td>
                  </tr>

                  <tr>
                    <th>CSS classes</th>
                    <td>
                      <code>
                        {esc(", ".join(sorted(analysis.classes)[:200]))}
                      </code>
                    </td>
                  </tr>

                  <tr>
                    <th>data-* attributes</th>
                    <td>
                      <code>
                        {esc(", ".join(sorted(analysis.data_attributes)[:200]))}
                      </code>
                    </td>
                  </tr>

                  <tr>
                    <th>Structural signals</th>
                    <td>
                      {
                        "<br>".join(
                            "<code>" + esc(value) + "</code>"
                            for value in sorted(
                                analysis.structural_signals
                            )[:200]
                        )
                        or "None"
                      }
                    </td>
                  </tr>
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
        <tr>
          <th>REST API URL</th>
          <td class="break">
            <code>{esc(sample.url)}</code>
          </td>
        </tr>

        <tr>
          <th>Browser</th>
          <td>{esc(sample.browser)}</td>
        </tr>

        <tr>
          <th>Status</th>
          <td>
            {'Success' if sample.ok else 'Failed'}
          </td>
        </tr>

        <tr>
          <th>Response type</th>
          <td>{esc(sample.response_kind)}</td>
        </tr>

        <tr>
          <th>Page title</th>
          <td>{esc(sample.page_title or 'Not detected')}</td>
        </tr>

        <tr>
          <th>Error</th>
          <td>{esc(sample.error or 'None')}</td>
        </tr>
      </table>

      <h3>Targeted REST Fields</h3>

      <div class="scroll">
        <table>
          <thead>
            <tr>
              <th>JSON Path</th>
              <th>Value Preview</th>
            </tr>
          </thead>

          <tbody>
            {interesting_rows}
          </tbody>
        </table>
      </div>

      {
        ''.join(html_blocks)
        if html_blocks
        else '<p>No embedded HTML / Canvas fragments detected.</p>'
      }

      <details>
        <summary>
          Raw REST Response — Sample {sample.number}
        </summary>

        <div class="pad">
          <pre>{esc(sample.raw_text)}</pre>
        </div>
      </details>

    </section>
    """


# ---------------------------------------------------------------------------
# HTML REPORT
# ---------------------------------------------------------------------------

def build_report(samples: List[Sample]) -> str:
    sample_count = len(samples)

    rows: List[Dict[str, Any]] = []

    rows += build_presence_rows(
        samples,
        [sample.json_paths for sample in samples],
        "REST / JSON Path",
    )

    rows += build_presence_rows(
        samples,
        [merged_set(sample, "classes") for sample in samples],
        "CSS Class",
    )

    rows += build_presence_rows(
        samples,
        [merged_set(sample, "ids") for sample in samples],
        "HTML ID",
    )

    rows += build_presence_rows(
        samples,
        [
            merged_set(sample, "data_attributes")
            for sample in samples
        ],
        "data-* Attribute",
    )

    rows += build_presence_rows(
        samples,
        [merged_set(sample, "attributes") for sample in samples],
        "HTML Attribute",
    )

    rows += build_presence_rows(
        samples,
        [tag_set(sample) for sample in samples],
        "HTML Tag",
    )

    rows += build_presence_rows(
        samples,
        [
            merged_set(sample, "structural_signals")
            for sample in samples
        ],
        "Structural Signal",
    )

    rank = {
        "Consistent": 0,
        "Unique": 1,
        "Variant": 2,
    }

    ordered = sorted(
        rows,
        key=lambda row: (
            rank[row["finding"]],
            row["category"],
            row["item"].lower(),
        ),
    )

    ordered = ordered[:MAX_MATRIX_ROWS]

    consistent = [
        row
        for row in ordered
        if row["finding"] == "Consistent"
    ]

    unique = [
        row
        for row in ordered
        if row["finding"] == "Unique"
    ]

    variants = [
        row
        for row in ordered
        if row["finding"] == "Variant"
    ]

    success_count = sum(
        1
        for sample in samples
        if sample.ok
    )

    cards = []

    for sample in samples:
        status_class = "ok" if sample.ok else "bad"
        status_text = "Success" if sample.ok else "Failed"

        cards.append(
            f"""
            <div class="card">

              <strong>Sample {sample.number}</strong>
              <br>

              <span class="{status_class}">
                {status_text}
              </span>

              <br>

              {esc(sample.response_kind)}

              <br>

              {esc(sample.page_title or 'Title not detected')}

              {
                '<br><span class="bad">' + esc(sample.error) + '</span>'
                if sample.error
                else ''
              }

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
        repeat(auto-fit, minmax(210px, 1fr));
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
    max-height: 650px;
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
· Generated
{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
</div>

</header>

<main>

<h2>Consolidated Diagnostic Summary</h2>

<p>
<strong>{success_count}/{sample_count}</strong>
samples were successfully retrieved and analyzed.
</p>

<p>
The analyzer compares the SharePoint REST structures objectively.
It does not yet assign Level 1 / Level 2 / Level 3 procedure meaning.
</p>

<div class="cards">
{''.join(cards)}
</div>

<h2>Consolidated Attribute Comparison Matrix</h2>

<p>
<strong>Consistent</strong> = present in every sample.
<strong>Variant</strong> = present in more than one but not all samples.
<strong>Unique</strong> = present in one sample only.
</p>

{matrix(
    ordered,
    sample_count,
    "All Prioritized Attributes and Patterns"
)}

<h2>Consistent Attributes</h2>

{matrix(
    consistent,
    sample_count,
    "Present Across All Samples"
)}

<h2>Unique Variants</h2>

{matrix(
    unique,
    sample_count,
    "Present in One Sample Only"
)}

<h2>Shared Variants</h2>

{matrix(
    variants,
    sample_count,
    "Present in Multiple But Not All Samples"
)}

<h2>Individual Sample Evidence</h2>

{''.join(
    sample_details(sample)
    for sample in samples
)}

</main>

</body>

</html>
"""


# ---------------------------------------------------------------------------
# USER INPUT
# ---------------------------------------------------------------------------

def prompt_urls() -> List[str]:
    print("=" * 78)

    print(
        f"{APP_NAME} v{VERSION}"
    )

    print("=" * 78)

    print()

    print(
        "Paste up to five SharePoint REST API URLs."
    )

    print(
        "Press Enter to leave a sample slot blank."
    )

    print()

    urls: List[str] = []

    for index in range(1, MAX_SAMPLES + 1):
        urls.append(
            input(
                f"REST API URL {index}: "
            ).strip()
        )

    if not any(urls):
        print()
        print("ERROR: No URLs supplied.")
        return []

    return urls


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main() -> int:
    urls = prompt_urls()

    if not urls:
        return 1

    driver = None

    try:
        print()
        print(
            "Starting Microsoft Edge using your existing profile..."
        )

        driver, browser_name, By = start_browser()

        print()
        print(
            f"Browser started: {browser_name}"
        )

        print()

        samples: List[Sample] = []

        for index, url in enumerate(
            urls,
            start=1,
        ):
            if not url:
                samples.append(
                    Sample(
                        number=index,
                        url="",
                        browser=browser_name,
                        error="No URL supplied.",
                    )
                )

                print(
                    f"Sample {index}: skipped"
                )

                continue

            print(
                f"Sample {index}: opening REST API URL..."
            )

            result = fetch_via_browser(
                driver,
                browser_name,
                By,
                index,
                url,
            )

            samples.append(result)

            if result.ok:
                print(
                    f"  SUCCESS - "
                    f"{result.response_kind}, "
                    f"{len(result.raw_text):,} characters"
                )

            else:
                print(
                    f"  FAILED - {result.error}"
                )

        print()

        print(
            "Building consolidated HTML report..."
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

        print(
            output.resolve()
        )

        print()

        print(
            "Open the HTML report in your browser."
        )

        print(
            "For a valid SharePoint REST retrieval, "
            "the samples should normally show JSON rather than "
            "a small HTML/Text response."
        )

        print()

        return 0

    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


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
