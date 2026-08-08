import sys, json, html, re, urllib.parse, traceback
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

MAX=5
TIMEOUT=60000

def log(s,m): print(f"[{s:<12}] {m}", flush=True)

def rest_url(u):
    x=urllib.parse.urlparse(u.strip())
    if x.scheme!="https" or not x.netloc: raise ValueError("Full HTTPS URL required")
    i=x.path.lower().find("/sitepages/")
    if i<0: raise ValueError("URL must contain /SitePages/")
    site=x.path[:i]; page=x.path.replace("'","''")
    rp=f"{site}/_api/web/GetFileByServerRelativeUrl('{page}')/ListItemAllFields"
    q=urllib.parse.urlencode({"$select":"Title,CanvasContent1,LayoutWebpartsContent"},safe="$,")
    return urllib.parse.urlunparse((x.scheme,x.netloc,rp,"",q,""))

def find_json(o,key):
    if isinstance(o,dict):
        for k,v in o.items():
            if str(k).lower()==key.lower(): return v
        for v in o.values():
            r=find_json(v,key)
            if r is not None:return r
    if isinstance(o,list):
        for v in o:
            r=find_json(v,key)
            if r is not None:return r
    return None

def parse(raw,ctype):
    s=raw.lstrip()
    if "json" in ctype.lower() or s.startswith(("{","[")):
        o=json.loads(raw)
        return "JSON",*(str(find_json(o,k) or "") for k in ("Title","CanvasContent1","LayoutWebpartsContent"))
    root=ET.fromstring(raw); vals={"title":"","canvascontent1":"","layoutwebpartscontent":""}
    for e in root.iter():
        n=e.tag.split("}")[-1].lower()
        if n in vals and e.text: vals[n]=e.text
    return "XML",vals["title"],vals["canvascontent1"],vals["layoutwebpartscontent"]

def decode(v):
    for _ in range(6):
        n=html.unescape(v)
        if n==v:break
        v=n
    return v

def main():
    print("="*80); print("SharePoint API Reader Baseline - Playwright Edge Diagnostic v2"); print("="*80)
    urls=[]
    for i in range(1,MAX+1):
        u=input(f"SharePoint Page URL {i}: ").strip()
        if u: urls.append((i,u))
    if not urls:return 1
    out=Path("SharePoint_API_Reader_Output_"+datetime.now().strftime("%Y%m%d_%H%M%S")); out.mkdir()
    log("OUTPUT",str(out.resolve()))
    try:
      with sync_playwright() as pw:
        log("PLAYWRIGHT","Initialized")
        log("BROWSER","Launching installed Microsoft Edge (channel=msedge, headless=False)")
        browser=pw.chromium.launch(channel="msedge",headless=False)
        log("BROWSER","Edge launched")
        ctx=browser.new_context(); log("CONTEXT","Created")
        page=ctx.new_page(); page.set_default_navigation_timeout(TIMEOUT); log("PAGE","Created")
        results=[]
        for n,u in urls:
            print("\n"+"-"*80); log("SAMPLE",f"Starting {n}"); log("INPUT",u)
            try:rurl=rest_url(u); log("REST URL",rurl)
            except Exception as e: log("URL FAIL",repr(e)); results.append((n,False,str(e))); continue
            try:
                log("NAVIGATE","page.goto REST endpoint")
                resp=page.goto(rurl,wait_until="domcontentloaded",timeout=TIMEOUT)
            except PWTimeout as e: log("TIMEOUT",str(e)); results.append((n,False,"timeout")); continue
            except Exception as e: log("NAV FAIL",f"{type(e).__name__}: {e}"); results.append((n,False,str(e))); continue
            if resp is None: log("HTTP FAIL","No response object"); results.append((n,False,"no response")); continue
            ctype=resp.headers.get("content-type","")
            log("HTTP",f"Status={resp.status} OK={resp.ok}"); log("HTTP",f"Final URL={resp.url}"); log("HTTP",f"Content-Type={ctype}")
            if resp.status in (401,403):
                log("AUTH FAIL",f"SharePoint rejected Playwright context with HTTP {resp.status}")
                results.append((n,False,f"HTTP {resp.status}")); continue
            try: raw=resp.text(); log("BODY",f"Direct response captured: {len(raw):,} chars")
            except Exception as e: log("BODY FAIL",str(e)); results.append((n,False,str(e))); continue
            ext=".json" if "json" in ctype.lower() else ".xml"
            rf=out/f"sample_{n}_raw_rest{ext}"; rf.write_text(raw,encoding="utf-8"); log("SAVE RAW",str(rf.resolve()))
            try:
                kind,title,canvas,layout=parse(raw,ctype)
                log("PARSE",f"Type={kind}"); log("PARSE",f"Title={title or '[not detected]'}")
                log("PARSE",f"CanvasContent1 found={bool(canvas)} chars={len(canvas):,}")
                log("PARSE",f"LayoutWebpartsContent found={bool(layout)} chars={len(layout):,}")
            except Exception as e:
                log("PARSE FAIL",f"{type(e).__name__}: {e}"); results.append((n,False,str(e))); continue
            if not canvas:
                log("CANVAS FAIL","Valid REST response but CanvasContent1 missing/empty"); results.append((n,False,"Canvas missing")); continue
            dc=decode(canvas); cf=out/f"sample_{n}_CanvasContent1.html"; cf.write_text(dc,encoding="utf-8")
            log("CANVAS",f"Decoded={len(dc):,} chars"); log("SAVE CANVAS",str(cf.resolve())); log("SUCCESS",f"Sample {n} complete")
            results.append((n,True,"Success"))
        print("\n"+"="*80+"\nRUN SUMMARY\n"+"="*80)
        for n,ok,d in results: print(f"Sample {n}: {'SUCCESS' if ok else 'FAILED'} - {d}")
        print(f"\nOutput: {out.resolve()}")
        browser.close()
        return 0
    except Exception:
        log("FATAL","Unhandled error; traceback follows"); traceback.print_exc(); return 2

if __name__=="__main__": sys.exit(main())
