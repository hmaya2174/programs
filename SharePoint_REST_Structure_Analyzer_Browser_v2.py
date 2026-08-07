#!/usr/bin/env python3
"""
SharePoint REST Structure Analyzer - Browser Assisted v2

Five SharePoint REST API URLs in -> one consolidated HTML report out.

Why Browser Assisted?
---------------------
Direct Python requests can return HTTP 401 even when the same SharePoint REST
URL works in Chrome. This version retrieves each URL through a real Chromium
browser session, allowing normal corporate SSO/browser authentication to apply.

Preferred browser order:
1. Microsoft Edge
2. Google Chrome

How it works
------------
- Prompts for up to 5 SharePoint REST API URLs.
- Launches Edge/Chrome through Selenium.
- Opens each REST API URL in the browser.
- Reads the returned JSON/XML/text from the browser page.
- Analyzes all samples together.
- Creates ONE HTML diagnostic report.

Dependencies
------------
pip install selenium beautifulsoup4

Notes
-----
- No username/password is requested or stored.
- If your environment automatically signs you in to SharePoint in Edge/Chrome,
  the browser should retrieve the REST response without manual copying.
- If the browser displays a sign-in page, complete the normal browser sign-in
  in the opened window, then return to the console and press Enter.
"""

from __future__ import annotations

import html as html_lib
import json
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
VERSION = "2.0 Browser Assisted"
MAX_SAMPLES = 5
OUTPUT_PREFIX = "SharePoint_REST_Structure_Analysis_Report"
MAX_MATRIX_ROWS = 700
MAX_RAW_CHARS = 2_000_000


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


def esc(v: Any) -> str:
    return html_lib.escape("" if v is None else str(v))


def norm(v: str) -> str:
    return re.sub(r"\s+", " ", v or "").strip()


def short(v: str, n: int = 250) -> str:
    v = norm(v)
    return v if len(v) <= n else v[:n-3] + "..."


def type_name(v: Any) -> str:
    if v is None: return "null"
    if isinstance(v, bool): return "boolean"
    if isinstance(v, int) and not isinstance(v, bool): return "integer"
    if isinstance(v, float): return "number"
    if isinstance(v, dict): return "object"
    if isinstance(v, list): return "array"
    if isinstance(v, str): return "string"
    return type(v).__name__


def flatten_json(v: Any, path: str="$",
                 paths: Optional[Set[str]]=None,
                 types: Optional[Dict[str,str]]=None,
                 depth: int=0) -> Tuple[Set[str], Dict[str,str]]:
    if paths is None: paths=set()
    if types is None: types={}
    if depth > 50:
        return paths, types
    paths.add(path)
    types[path] = type_name(v)

    if isinstance(v, dict):
        for k, child in v.items():
            flatten_json(child, f"{path}.{k}", paths, types, depth+1)
    elif isinstance(v, list):
        p = f"{path}[*]"
        paths.add(p)
        if not v:
            types[p] = "empty-array"
        else:
            for child in v[:250]:
                flatten_json(child, p, paths, types, depth+1)
    return paths, types


def walk_json(v: Any, path: str="$") -> Iterable[Tuple[str,str,Any]]:
    if isinstance(v, dict):
        for k, child in v.items():
            p=f"{path}.{k}"
            yield p, str(k), child
            yield from walk_json(child,p)
    elif isinstance(v,list):
        for child in v[:500]:
            yield from walk_json(child,f"{path}[*]")


def looks_html(v: str) -> bool:
    if not isinstance(v,str) or len(v)<10: return False
    x=v.lower()
    return any(t in x for t in ("<div","<p","<section","<table","<span","<a ","data-sp-","<article"))


def depth(tag: Tag) -> int:
    d=0
    p=tag.parent
    while isinstance(p,Tag):
        d+=1
        p=p.parent
    return d


def analyze_html(source: str, txt: str) -> HtmlAnalysis:
    soup=BeautifulSoup(txt or "","html.parser")
    tags=Counter()
    classes=set()
    ids=set()
    data_attrs=set()
    attrs=set()
    signals=set()
    headings=[]
    max_depth=0

    for tag in soup.find_all(True):
        name=(tag.name or "").lower()
        tags[name]+=1
        max_depth=max(max_depth,depth(tag))

        for a,val in tag.attrs.items():
            al=str(a).lower()
            attrs.add(al)
            if al.startswith("data-"):
                data_attrs.add(al)

            if al=="id":
                s=str(val).strip()
                if s: ids.add(s)

            if al=="class":
                vals=val if isinstance(val,list) else [val]
                for c in vals:
                    c=str(c).strip()
                    if not c: continue
                    classes.add(c)
                    lc=c.lower()
                    if any(k in lc for k in (
                        "tab","accordion","decision","branch","step","expand",
                        "collapse","section","canvas","webpart","web-part","control"
                    )):
                        signals.add("class:"+c)

        role=str(tag.attrs.get("role","")).strip()
        if role: signals.add("role:"+role)

        if name in {"table","ul","ol","section","article","details","summary"}:
            signals.add("tag:"+name)

        if name in {"h1","h2","h3","h4","h5","h6"}:
            t=short(tag.get_text(" ",strip=True),180)
            if t: headings.append(t)

    # Detect nested classes that may be useful for future rule discovery.
    tokens=("tab","accordion","decision","branch","step","section","canvas","webpart")
    nested=Counter()
    for tag in soup.find_all(True):
        cc=tag.attrs.get("class",[])
        if isinstance(cc,str): cc=[cc]
        child=" ".join(str(x).lower() for x in cc)
        if not any(t in child for t in tokens):
            continue
        p=tag.parent
        while isinstance(p,Tag):
            pc=p.attrs.get("class",[])
            if isinstance(pc,str): pc=[pc]
            parent=" ".join(str(x).lower() for x in pc)
            if any(t in parent for t in tokens):
                nested[f"{short(parent,80)} -> {short(child,80)}"]+=1
                break
            p=p.parent

    for k,v in nested.items():
        signals.add(f"nested:{k} ({v})")

    clean=norm(soup.get_text(" ",strip=True))
    return HtmlAnalysis(
        source_field=source,
        length=len(txt or ""),
        text_length=len(clean),
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
        list_count=len(soup.find_all(["ul","ol"])),
        text_preview=clean[:5000],
    )


def extract_json(data: Any) -> Tuple[Set[str],Dict[str,str],List[HtmlAnalysis],Dict[str,str],str]:
    paths,types=flatten_json(data)
    analyses=[]
    interesting={}
    title=""
    special={
        "canvascontent1","layoutwebpartscontent","title","pagelayouttype",
        "promotedstate","description","bannerimageurl","webpartdata",
        "controltype","id","uniqueid","fileleafref","fileref"
    }
    seen=set()

    if isinstance(data,dict):
        for k,v in data.items():
            if str(k).lower()=="title" and isinstance(v,str):
                title=v

    for p,k,v in walk_json(data):
        kl=k.lower()
        if kl in special:
            if isinstance(v,(dict,list)):
                preview=short(json.dumps(v,ensure_ascii=False),600)
            else:
                preview=short(str(v),600)
            interesting[p]=preview
            if kl=="title" and not title and isinstance(v,str):
                title=v

        if isinstance(v,str) and (kl in {"canvascontent1","layoutwebpartscontent"} or looks_html(v)):
            hv=hash(v)
            if hv not in seen:
                seen.add(hv)
                analyses.append(analyze_html(p,v))
    return paths,types,analyses,interesting,title


def import_selenium():
    try:
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options as EdgeOptions
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.common.by import By
        from selenium.common.exceptions import WebDriverException
        return webdriver, EdgeOptions, ChromeOptions, By, WebDriverException
    except ImportError:
        print("\nERROR: Selenium is not installed.")
        print("Install it once using:")
        print("    pip install selenium beautifulsoup4")
        raise


def start_browser():
    webdriver, EdgeOptions, ChromeOptions, By, WebDriverException = import_selenium()

    # First try Edge because Microsoft 365/SharePoint SSO commonly works best there.
    edge_error=""
    try:
        opts=EdgeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--disable-popup-blocking")
        driver=webdriver.Edge(options=opts)
        return driver,"Microsoft Edge",By
    except Exception as exc:
        edge_error=str(exc)

    # Fall back to Chrome.
    try:
        opts=ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--disable-popup-blocking")
        driver=webdriver.Chrome(options=opts)
        return driver,"Google Chrome",By
    except Exception as chrome_exc:
        raise RuntimeError(
            "Could not start Edge or Chrome through Selenium.\n\n"
            f"Edge error: {edge_error}\n\n"
            f"Chrome error: {chrome_exc}\n\n"
            "Make sure Edge or Chrome is installed and Selenium is current."
        )


def browser_body_text(driver, By) -> str:
    try:
        body=driver.find_element(By.TAG_NAME,"body")
        return body.text or ""
    except Exception:
        return ""


def browser_page_source(driver) -> str:
    try:
        return driver.page_source or ""
    except Exception:
        return ""


def looks_like_signin(text: str, source: str) -> bool:
    blob=(text+"\n"+source[:20000]).lower()
    return (
        ("sign in" in blob or "signin" in blob or "login" in blob or "log in" in blob)
        and ("password" in blob or "microsoft" in blob or "account" in blob)
    )


def extract_pre_text(driver, By) -> str:
    # Chrome/Edge JSON viewer usually renders response in <pre>.
    try:
        pres=driver.find_elements(By.TAG_NAME,"pre")
        candidates=[p.text for p in pres if (p.text or "").strip()]
        if candidates:
            return max(candidates,key=len)
    except Exception:
        pass
    return ""


def fetch_via_browser(driver, browser_name: str, By, number: int, url: str) -> Sample:
    s=Sample(number=number,url=url,browser=browser_name)
    try:
        driver.get(url)
        time.sleep(2.0)

        text=extract_pre_text(driver,By)
        if not text:
            text=browser_body_text(driver,By)
        source=browser_page_source(driver)

        if looks_like_signin(text,source):
            print(f"  Sample {number}: browser appears to be on a sign-in page.")
            print("  Complete your normal TD/Microsoft sign-in in the browser window.")
            input("  When the REST response is visible, press Enter here to continue: ")
            time.sleep(1.0)
            text=extract_pre_text(driver,By) or browser_body_text(driver,By)
            source=browser_page_source(driver)

        # Sometimes browser JSON viewer contains body text with pretty JSON.
        raw=(text or "").strip()

        if not raw:
            s.error="Browser returned an empty page."
            return s

        if len(raw)>MAX_RAW_CHARS:
            s.raw_text=raw[:MAX_RAW_CHARS]
        else:
            s.raw_text=raw

        # Detect JSON.
        parsed=None
        try:
            parsed=json.loads(raw)
            s.response_kind="JSON"
        except Exception:
            # Some browser JSON viewers prepend text; attempt to find first JSON object/array.
            candidates=[]
            oi=raw.find("{")
            ai=raw.find("[")
            if oi>=0: candidates.append(oi)
            if ai>=0: candidates.append(ai)
            for start in sorted(candidates):
                try:
                    parsed=json.loads(raw[start:])
                    raw=raw[start:]
                    s.raw_text=raw[:MAX_RAW_CHARS]
                    s.response_kind="JSON"
                    break
                except Exception:
                    pass

        if parsed is not None:
            if isinstance(parsed,dict):
                s.top_fields={str(k) for k in parsed.keys()}
            paths,types,analyses,interesting,title=extract_json(parsed)
            s.json_paths=paths
            s.path_types=types
            s.html_analyses=analyses
            s.interesting_fields=interesting
            s.page_title=title
            s.ok=True
            return s

        # XML/HTML/text fallback.
        low=raw.lstrip().lower()
        if low.startswith("<?xml") or low.startswith("<feed") or low.startswith("<entry"):
            s.response_kind="XML"
        elif "<html" in source.lower() or "<body" in source.lower():
            s.response_kind="HTML/Text"
        else:
            s.response_kind="Text"

        s.html_analyses=[analyze_html("$browser_response",source or raw)]
        soup=BeautifulSoup(source or raw,"html.parser")
        if soup.title:
            s.page_title=norm(soup.title.get_text(" ",strip=True))
        s.ok=True
        return s

    except Exception as exc:
        s.error=str(exc)
        return s


def merged_set(sample: Sample, attr: str) -> Set[str]:
    out=set()
    for h in sample.html_analyses:
        out.update(getattr(h,attr,set()))
    return out


def tag_set(sample: Sample) -> Set[str]:
    out=set()
    for h in sample.html_analyses:
        out.update(h.tag_counts.keys())
    return out


def presence_rows(samples: List[Sample], sets: List[Set[str]], category: str) -> List[dict]:
    universe=set()
    for st in sets: universe.update(st)
    rows=[]
    n=len(samples)
    for item in sorted(universe,key=str.lower):
        p=[item in st for st in sets]
        f=sum(p)
        finding="Consistent" if f==n else ("Unique" if f==1 else "Variant")
        rows.append(dict(category=category,item=item,presence=p,frequency=f,finding=finding))
    return rows


def matrix(rows: List[dict], n: int, heading: str) -> str:
    if not rows:
        return f"<h3>{esc(heading)}</h3><p>No comparable structures detected.</p>"
    hs="".join(f"<th>S{i}</th>" for i in range(1,n+1))
    body=[]
    for r in rows:
        checks="".join(f"<td class='c'>{'✓' if x else '—'}</td>" for x in r["presence"])
        body.append(
            f"<tr class='{r['finding'].lower()}'>"
            f"<td>{esc(r['category'])}</td>"
            f"<td><code>{esc(r['item'])}</code></td>"
            f"{checks}<td class='c'>{r['frequency']}/{n}</td>"
            f"<td><strong>{esc(r['finding'])}</strong></td></tr>"
        )
    return f"""<h3>{esc(heading)}</h3><div class="scroll"><table>
    <thead><tr><th>Category</th><th>Attribute / Pattern</th>{hs}<th>Frequency</th><th>Finding</th></tr></thead>
    <tbody>{''.join(body)}</tbody></table></div>"""


def sample_details(s: Sample) -> str:
    interesting="".join(
        f"<tr><td><code>{esc(p)}</code></td><td>{esc(v)}</td></tr>"
        for p,v in sorted(s.interesting_fields.items())
    ) or "<tr><td colspan='2'>No targeted fields detected.</td></tr>"

    blocks=[]
    for i,h in enumerate(s.html_analyses,1):
        blocks.append(f"""
        <details><summary>HTML / Canvas Fragment {i}: <code>{esc(h.source_field)}</code></summary>
        <div class="pad">
          <table>
            <tr><th>HTML length</th><td>{h.length:,}</td></tr>
            <tr><th>Maximum DOM depth</th><td>{h.max_depth}</td></tr>
            <tr><th>Links / Images / Tables / Lists</th><td>{h.link_count} / {h.image_count} / {h.table_count} / {h.list_count}</td></tr>
            <tr><th>Classes</th><td><code>{esc(", ".join(sorted(h.classes)[:150]))}</code></td></tr>
            <tr><th>data-* attributes</th><td><code>{esc(", ".join(sorted(h.data_attributes)[:150]))}</code></td></tr>
            <tr><th>Structural signals</th><td>{"<br>".join("<code>"+esc(x)+"</code>" for x in sorted(h.structural_signals)[:150]) or "None"}</td></tr>
          </table>
          <h4>Clean text preview</h4><pre>{esc(h.text_preview)}</pre>
        </div></details>
        """)

    return f"""
    <section>
      <h2>Sample {s.number}</h2>
      <table>
        <tr><th>REST URL</th><td class="break"><code>{esc(s.url)}</code></td></tr>
        <tr><th>Browser</th><td>{esc(s.browser)}</td></tr>
        <tr><th>Status</th><td>{'Success' if s.ok else 'Failed'}</td></tr>
        <tr><th>Response type</th><td>{esc(s.response_kind)}</td></tr>
        <tr><th>Page title</th><td>{esc(s.page_title or 'Not detected')}</td></tr>
        <tr><th>Error</th><td>{esc(s.error or 'None')}</td></tr>
      </table>

      <h3>Targeted REST Fields</h3>
      <div class="scroll"><table><thead><tr><th>JSON Path</th><th>Value Preview</th></tr></thead>
      <tbody>{interesting}</tbody></table></div>

      {''.join(blocks) if blocks else '<p>No embedded HTML/Canvas fragments detected.</p>'}

      <details><summary>Raw REST Response — Sample {s.number}</summary>
      <div class="pad"><pre>{esc(s.raw_text)}</pre></div></details>
    </section>
    """


def build_report(samples: List[Sample]) -> str:
    n=len(samples)
    rows=[]
    rows+=presence_rows(samples,[s.json_paths for s in samples],"REST / JSON Path")
    rows+=presence_rows(samples,[merged_set(s,"classes") for s in samples],"CSS Class")
    rows+=presence_rows(samples,[merged_set(s,"data_attributes") for s in samples],"data-* Attribute")
    rows+=presence_rows(samples,[merged_set(s,"attributes") for s in samples],"HTML Attribute")
    rows+=presence_rows(samples,[tag_set(s) for s in samples],"HTML Tag")
    rows+=presence_rows(samples,[merged_set(s,"structural_signals") for s in samples],"Structural Signal")

    rank={"Consistent":0,"Unique":1,"Variant":2}
    ordered=sorted(rows,key=lambda r:(rank[r["finding"]],r["category"],r["item"].lower()))
    ordered=ordered[:MAX_MATRIX_ROWS]
    cons=[r for r in ordered if r["finding"]=="Consistent"]
    uniq=[r for r in ordered if r["finding"]=="Unique"]
    var=[r for r in ordered if r["finding"]=="Variant"]

    success=sum(1 for s in samples if s.ok)

    cards=[]
    for s in samples:
        cards.append(
            f"<div class='card'><strong>Sample {s.number}</strong><br>"
            f"<span class='{'ok' if s.ok else 'bad'}'>{'Success' if s.ok else 'Failed'}</span><br>"
            f"{esc(s.response_kind)}<br>{esc(s.page_title or 'Title not detected')}</div>"
        )

    return f"""<!doctype html>
<html><head><meta charset="utf-8">
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
.c{{text-align:center}} .consistent td{{background:#eef8f0}} .unique td{{background:#fff8c5}}
.variant td{{background:#f6f8fa}} .scroll{{overflow:auto;max-height:650px;border:1px solid #d0d7de}}
pre{{background:#0d1117;color:#e6edf3;padding:14px;border-radius:6px;white-space:pre-wrap;word-break:break-word;max-height:650px;overflow:auto;font-size:12px}}
code{{font-family:Consolas,monospace;word-break:break-word}}
details{{border:1px solid #d0d7de;border-radius:6px;margin:10px 0}}
summary{{cursor:pointer;background:#f6f8fa;padding:10px;font-weight:600}} .pad{{padding:12px}}
.break{{word-break:break-all}}
</style></head>
<body><header>
<h1>{esc(APP_NAME)}</h1>
<div>Version {esc(VERSION)} · Generated {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>
</header><main>
<h2>Consolidated Diagnostic Summary</h2>
<p><strong>{success}/{n}</strong> samples were retrieved successfully through the authenticated browser session.</p>
<div class="cards">{''.join(cards)}</div>

<h2>Consolidated Attribute Comparison Matrix</h2>
<p>The analyzer reports what is structurally consistent, shared as a variant, or unique across the five REST responses. It does not yet assign Level 1/2/3 business meaning.</p>
{matrix(ordered,n,"All Prioritized Attributes and Patterns")}

<h2>Consistent Attributes</h2>
{matrix(cons,n,"Present Across All Five Samples")}

<h2>Unique Variants</h2>
{matrix(uniq,n,"Present in One Sample Only")}

<h2>Shared Variants</h2>
{matrix(var,n,"Present in Two to Four Samples")}

<h2>Individual Evidence</h2>
{''.join(sample_details(s) for s in samples)}
</main></body></html>"""


def prompt_urls() -> List[str]:
    print("="*78)
    print(f"{APP_NAME} - {VERSION}")
    print("="*78)
    print()
    print("Paste five SharePoint REST API URLs.")
    print("The program will open them through Edge/Chrome and create one HTML report.")
    print()
    urls=[]
    for i in range(1,MAX_SAMPLES+1):
        urls.append(input(f"REST API URL {i}: ").strip())
    if not any(urls):
        print("ERROR: No URLs supplied.")
        return []
    return urls


def main() -> int:
    urls=prompt_urls()
    if not urls:
        return 1

    print("\nStarting authenticated browser...")
    driver=None
    try:
        driver,browser_name,By=start_browser()
        print(f"Browser started: {browser_name}")
        print("If corporate SSO is active, SharePoint should authenticate normally.\n")

        samples=[]
        for i,url in enumerate(urls,1):
            if not url:
                samples.append(Sample(number=i,url="",browser=browser_name,error="No URL supplied."))
                print(f"Sample {i}: skipped")
                continue

            print(f"Sample {i}: opening REST URL...")
            s=fetch_via_browser(driver,browser_name,By,i,url)
            samples.append(s)
            if s.ok:
                print(f"  SUCCESS - {s.response_kind}, {len(s.raw_text):,} characters")
            else:
                print(f"  FAILED - {s.error}")

        print("\nBuilding consolidated HTML report...")
        report=build_report(samples)
        stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        out=Path(f"{OUTPUT_PREFIX}_{stamp}.html")
        out.write_text(report,encoding="utf-8")

        print("\n"+"="*78)
        print("REPORT CREATED")
        print("="*78)
        print(out.resolve())
        print("\nYou can open this single HTML file in your browser and upload it for analysis.")
        return 0

    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass


if __name__=="__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        sys.exit(130)
    except Exception:
        print("\nUNEXPECTED ERROR")
        traceback.print_exc()
        sys.exit(2)
