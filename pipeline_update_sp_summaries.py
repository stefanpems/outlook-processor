"""Update SP BlogPosts items with summaries via REST API."""
import json, sys, os
sys.stdout = open(sys.stdout.fileno(), mode='w', buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

updates = json.load(open(os.path.join(BASE, "sp_summary_updates.json"), encoding="utf-8"))
print(f"{len(updates)} items to update:")
for u in updates:
    sid = u["sp_id"]
    title = u["sp_title"][:60]
    summary = u["summary"][:80]
    print(f"  ID={sid} {title}")
    print(f"    Summary: {summary}...")

# Connect to Edge and update via SP REST API
from playwright.sync_api import sync_playwright
from urllib.parse import urlparse as _urlparse
from cdp_helper import ensure_edge_cdp

CDP_URL = CONFIG["edge_cdp"]["url"]
SP_API = CONFIG["sharepoint"]["blog_list_api"]
SP_LIST_URL = CONFIG["sharepoint"]["blog_list_url"]
_sp_site_path = _urlparse(CONFIG["sharepoint"]["site_base"]).path.rstrip('/')

ensure_edge_cdp()
p = sync_playwright().start()
browser = p.chromium.connect_over_cdp(CDP_URL)
ctx = browser.contexts[0]
page = ctx.new_page()

page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
page.wait_for_timeout(3000)

# Get request digest for write operations
digest = page.evaluate(f"""async () => {{
    const resp = await fetch(
        "{_sp_site_path}/_api/contextinfo",
        {{ method: "POST", headers: {{ "Accept": "application/json;odata=nometadata" }} }}
    );
    const data = await resp.json();
    return data.FormDigestValue;
}}""")
print(f"\nGot request digest: {digest[:30]}...")

# Update each item
success = 0
for u in updates:
    sp_id = u["sp_id"]
    summary_text = u["summary"].replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    
    result = page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}/items({sp_id})",
                {{
                    method: "POST",
                    headers: {{
                        "Accept": "application/json;odata=nometadata",
                        "Content-Type": "application/json;odata=nometadata",
                        "X-RequestDigest": "{digest}",
                        "IF-MATCH": "*",
                        "X-HTTP-Method": "MERGE"
                    }},
                    body: JSON.stringify({{ Summary: "{summary_text}" }})
                }}
            );
            return {{ ok: resp.ok, status: resp.status }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")
    
    if result.get("ok"):
        success += 1
        print(f"  Updated ID={sp_id}: OK")
    else:
        print(f"  Updated ID={sp_id}: FAILED ({result})")

print(f"\n{success}/{len(updates)} items updated successfully")

page.close()
p.stop()
