#!/usr/bin/env python3
"""
SharePoint API Reader Baseline - Playwright Edge v3
===================================================

Purpose
-------
Accept up to five EXACT SharePoint REST API URLs that are already known to work
when opened manually in Microsoft Edge.

The program does NOT construct or modify the REST URL.

For each REST API URL:
1. Playwright launches installed Microsoft Edge.
2. Navigates directly to the exact REST URL entered.
3. Captures the main HTTP response body directly.
4. Saves the raw response as XML/JSON.
5. Parses Title, CanvasContent1, and LayoutWebpartsContent.
6. Saves decoded CanvasContent1 as HTML.
7. Prints detailed debugging information to the console.

This is a baseline REST reader. It intentionally does not classify procedures.

Install
-------
python -m pip install playwright beautifulsoup4

Browser
-------
Uses the installed Microsoft Edge via:
    channel="msedge"

Notes
-----
- No username/password is requested or stored.
- Authentication depends on the Edge context available in your environment.
- If SharePoint returns HTTP 401/403, the console will show that clearly.
"""

from __future__ import annotations

import html
import json
import re
import sys
import traceback
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Tuple

try:
    from playwright.sync_api import (
        sync_playwright,
        TimeoutError as PlaywrightTimeoutError,
    )
except ImportError:
    print("[FATAL] Playwright is not installed.")
    print("Run: python -m pip install playwright")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("[FATAL] beautifulsoup4 is not installed.")
    print("Run: python -m pip install beautifulsoup4")
    sys.exit(1)


APP_NAME = "SharePoint API Reader Baseline"
VERSION = "3.0 Exact REST URL + Playwright Edge"
MAX_SAMPLES = 5
NAV_TIMEOUT_MS = 60_000
OUTPUT_PREFIX = "SharePoint_API_Reader_Output"


def log(stage: str, message: str) -> None:
    print(f"[{stage:<14}] {message}", flush=True)


def decode_repeated(value: str) -> str:
    current = value or ""

    for _ in range(6):
        decoded = html.unescape(current)

        if decoded == current:
            break

        current = decoded

    return current


def find_json_value(obj: Any, target: str) -> Optional[Any]:
    if isinstance(obj, dict):
        for key, value in obj.items():
            if str(key).lower() == target.lower():
                return value

        for value in obj.values():
            found = find_json_value(value, target)

            if found is not None:
                return found

    elif isinstance(obj, list):
        for value in obj:
            found = find_json_value(value, target)

            if found is not None:
                return found

    return None


def parse_json_payload(raw: str) -> Tuple[str, str, str]:
    obj = json.loads(raw)

    title = find_json_value(obj, "Title")
    canvas = find_json_value(obj, "CanvasContent1")
    layout = find_json_value(obj, "LayoutWebpartsContent")

    return (
        "" if title is None else str(title),
        "" if canvas is None else str(canvas),
        "" if layout is None else str(layout),
    )


def parse_xml_payload(raw: str) -> Tuple[str, str, str]:
    title = ""
    canvas = ""
    layout = ""

    root = ET.fromstring(raw)

    for element in root.iter():
        local_name = element.tag.split("}")[-1].lower()

        if local_name == "canvascontent1":
            canvas = element.text or ""

        elif local_name == "layoutwebpartscontent":
            layout = element.text or ""

    # Namespace-tolerant fallback.
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

    # Prefer the last non-empty title because Atom payloads may contain more
    # than one title node.
    title_nodes = soup.find_all(
        lambda tag: tag.name
        and tag.name.split(":")[-1].lower() == "title"
    )

    for node in reversed(title_nodes):
        value = re.sub(
            r"\s+",
            " ",
            node.get_text(" ", strip=True),
        ).strip()

        if value:
            title = value
            break

    return title, canvas, layout


def parse_payload(
    raw: str,
    content_type: str,
) -> Tuple[str, str, str, str]:
    stripped = (raw or "").lstrip()

    if not stripped:
        raise ValueError("Response body is empty.")

    content_type_low = (content_type or "").lower()

    if (
        "json" in content_type_low
        or stripped.startswith("{")
        or stripped.startswith("[")
    ):
        title, canvas, layout = parse_json_payload(raw)

        return "JSON", title, canvas, layout

    if (
        "xml" in content_type_low
        or stripped.startswith("<?xml")
        or stripped.startswith("<entry")
        or stripped.startswith("<feed")
    ):
        title, canvas, layout = parse_xml_payload(raw)

        return "XML", title, canvas, layout

    # Some Edge/SharePoint responses may return text/xml without an XML header.
    if "CanvasContent1" in stripped and "<" in stripped:
        try:
            title, canvas, layout = parse_xml_payload(raw)
            return "XML", title, canvas, layout
        except Exception:
            pass

    raise ValueError(
        "Response was received but was not recognized as SharePoint JSON/XML. "
        f"First 250 characters: {repr(stripped[:250])}"
    )


def safe_filename(value: str, fallback: str) -> str:
    value = re.sub(r"\s+", "_", value or "")
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    value = value.strip("._")

    return value or fallback


def main() -> int:
    print("=" * 88)
    print(f"{APP_NAME} - {VERSION}")
    print("=" * 88)
    print()
    print("Paste up to five EXACT SharePoint REST API URLs.")
    print("The program will use the URLs exactly as entered.")
    print("Press Enter for unused slots.")
    print()

    samples = []

    for index in range(1, MAX_SAMPLES + 1):
        url = input(f"REST API URL {index}: ").strip()

        if url:
            samples.append((index, url))

    if not samples:
        log("FATAL", "No REST API URLs supplied.")
        return 1

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path(f"{OUTPUT_PREFIX}_{timestamp}")
    output_dir.mkdir(parents=True, exist_ok=True)

    log("OUTPUT", str(output_dir.resolve()))
    log("START", "Starting Playwright.")
    log("BROWSER", "Browser channel = msedge")
    log("BROWSER", "Headless = False")
    log("BROWSER", f"Navigation timeout = {NAV_TIMEOUT_MS // 1000} seconds")

    results = []

    try:
        with sync_playwright() as playwright:
            log("PLAYWRIGHT", "Initialized successfully.")

            try:
                browser = playwright.chromium.launch(
                    channel="msedge",
                    headless=False,
                )

                log("BROWSER", "Microsoft Edge launched successfully.")

            except Exception as exc:
                log(
                    "BROWSER FAIL",
                    f"{type(exc).__name__}: {exc}",
                )
                raise

            context = browser.new_context()

            log("CONTEXT", "New Edge browser context created.")

            page = context.new_page()
            page.set_default_navigation_timeout(NAV_TIMEOUT_MS)

            log("PAGE", "Page created.")

            # Only report failures relevant to SharePoint API traffic.
            def on_request_failed(request):
                if "/_api/" in request.url.lower():
                    log(
                        "REQUEST FAIL",
                        f"{request.method} {request.url} | "
                        f"failure={request.failure}",
                    )

            page.on("requestfailed", on_request_failed)

            for sample_no, rest_url in samples:
                print()
                print("-" * 88)

                log("SAMPLE", f"Starting sample {sample_no}")
                log("REST URL", rest_url)

                if not rest_url.lower().startswith("https://"):
                    log(
                        "URL FAIL",
                        "REST API URL must begin with https://",
                    )

                    results.append(
                        (sample_no, False, "Invalid URL")
                    )

                    continue

                response = None

                try:
                    log(
                        "NAVIGATE",
                        "Opening exact REST API URL in Microsoft Edge...",
                    )

                    response = page.goto(
                        rest_url,
                        wait_until="domcontentloaded",
                        timeout=NAV_TIMEOUT_MS,
                    )

                    log(
                        "NAVIGATE",
                        "page.goto completed.",
                    )

                except PlaywrightTimeoutError as exc:
                    log(
                        "TIMEOUT",
                        f"Navigation timeout: {exc}",
                    )

                    results.append(
                        (sample_no, False, "Navigation timeout")
                    )

                    continue

                except Exception as exc:
                    log(
                        "NAV FAIL",
                        f"{type(exc).__name__}: {exc}",
                    )

                    results.append(
                        (sample_no, False, str(exc))
                    )

                    continue

                if response is None:
                    log(
                        "HTTP FAIL",
                        "No main-document HTTP response object was returned.",
                    )

                    results.append(
                        (sample_no, False, "No HTTP response")
                    )

                    continue

                status = response.status
                ok = response.ok
                final_url = response.url

                headers = response.headers
                content_type = headers.get("content-type", "")
                content_length = headers.get("content-length", "")

                log("HTTP", f"Status = {status}")
                log("HTTP", f"OK = {ok}")
                log("HTTP", f"Final URL = {final_url}")
                log(
                    "HTTP",
                    f"Content-Type = {content_type or '[not supplied]'}",
                )
                log(
                    "HTTP",
                    f"Content-Length header = "
                    f"{content_length or '[not supplied]'}",
                )

                if status in (401, 403):
                    log(
                        "AUTH FAIL",
                        f"SharePoint returned HTTP {status}. "
                        "Playwright Edge does not have the required authenticated "
                        "session for this request.",
                    )

                    results.append(
                        (sample_no, False, f"HTTP {status}")
                    )

                    continue

                if status >= 400:
                    log(
                        "HTTP FAIL",
                        f"SharePoint returned HTTP {status}.",
                    )

                    results.append(
                        (sample_no, False, f"HTTP {status}")
                    )

                    continue

                try:
                    raw = response.text()

                    log(
                        "BODY",
                        f"Captured response body directly: "
                        f"{len(raw):,} characters.",
                    )

                except Exception as exc:
                    log(
                        "BODY FAIL",
                        f"{type(exc).__name__}: {exc}",
                    )

                    results.append(
                        (sample_no, False, str(exc))
                    )

                    continue

                try:
                    kind, title, canvas, layout = parse_payload(
                        raw,
                        content_type,
                    )

                    log(
                        "PARSE",
                        f"Response recognized as {kind}.",
                    )
                    log(
                        "PARSE",
                        f"Page title = {title or '[not detected]'}",
                    )
                    log(
                        "PARSE",
                        f"CanvasContent1 found = {bool(canvas)}",
                    )
                    log(
                        "PARSE",
                        f"CanvasContent1 raw length = {len(canvas):,}",
                    )
                    log(
                        "PARSE",
                        f"LayoutWebpartsContent found = {bool(layout)}",
                    )
                    log(
                        "PARSE",
                        f"LayoutWebpartsContent raw length = {len(layout):,}",
                    )

                except Exception as exc:
                    log(
                        "PARSE FAIL",
                        f"{type(exc).__name__}: {exc}",
                    )

                    raw_file = (
                        output_dir
                        / f"sample_{sample_no}_unparsed_response.txt"
                    )

                    raw_file.write_text(
                        raw,
                        encoding="utf-8",
                    )

                    log(
                        "SAVE DEBUG",
                        str(raw_file.resolve()),
                    )

                    results.append(
                        (sample_no, False, str(exc))
                    )

                    continue

                extension = (
                    ".json"
                    if kind == "JSON"
                    else ".xml"
                )

                raw_name = safe_filename(
                    title,
                    f"sample_{sample_no}",
                )

                raw_file = (
                    output_dir
                    / f"sample_{sample_no}_{raw_name}_raw_rest{extension}"
                )

                raw_file.write_text(
                    raw,
                    encoding="utf-8",
                )

                log(
                    "SAVE RAW",
                    str(raw_file.resolve()),
                )

                if not canvas:
                    log(
                        "CANVAS FAIL",
                        "REST response was valid, but CanvasContent1 is empty/missing.",
                    )

                    results.append(
                        (
                            sample_no,
                            False,
                            "CanvasContent1 missing",
                        )
                    )

                    continue

                decoded_canvas = decode_repeated(canvas)

                canvas_file = (
                    output_dir
                    / f"sample_{sample_no}_{raw_name}_CanvasContent1.html"
                )

                canvas_file.write_text(
                    decoded_canvas,
                    encoding="utf-8",
                )

                log(
                    "CANVAS",
                    f"Decoded CanvasContent1 length = "
                    f"{len(decoded_canvas):,}",
                )

                log(
                    "SAVE CANVAS",
                    str(canvas_file.resolve()),
                )

                if layout:
                    layout_file = (
                        output_dir
                        / f"sample_{sample_no}_{raw_name}_LayoutWebpartsContent.txt"
                    )

                    layout_file.write_text(
                        decode_repeated(layout),
                        encoding="utf-8",
                    )

                    log(
                        "SAVE LAYOUT",
                        str(layout_file.resolve()),
                    )

                log(
                    "SUCCESS",
                    f"Sample {sample_no} completed successfully.",
                )

                results.append(
                    (sample_no, True, "Success")
                )

            print()
            print("=" * 88)
            print("RUN SUMMARY")
            print("=" * 88)

            for sample_no, success, detail in results:
                print(
                    f"Sample {sample_no}: "
                    f"{'SUCCESS' if success else 'FAILED'} - {detail}"
                )

            success_count = sum(
                1
                for _, success, _ in results
                if success
            )

            print()
            print(
                f"Successful samples: "
                f"{success_count}/{len(results)}"
            )

            print(
                f"Output folder: {output_dir.resolve()}"
            )

            log(
                "BROWSER",
                "Closing Playwright Edge.",
            )

            browser.close()

            return 0 if success_count else 2

    except Exception:
        print()
        log(
            "FATAL",
            "Unhandled error. Full traceback follows.",
        )

        traceback.print_exc()

        return 2


if __name__ == "__main__":
    sys.exit(main())
