"""Fetch VideoPosts list from SharePoint to use for deduplication.

Discovers writable fields and downloads all items to sp_videoposts.json.

Usage: python pipeline_fetch_videoposts.py
"""
import json, sys, os
sys.stdout = open(sys.stdout.fileno(), mode='w', buffering=1)
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SP_API = CONFIG["video_sharepoint"]["list_api"]
SP_LIST_URL = CONFIG["video_sharepoint"]["list_url"]
FIELDS = CONFIG["video_sharepoint"]["fields"]
F_PUBLISHED = FIELDS["published"]
F_DURATION = FIELDS["duration"]
F_YT_ID = FIELDS["yt_id"]
F_ABSTRACT = FIELDS["abstract"]

ensure_edge_cdp()
p = sync_playwright().start()
browser = p.chromium.connect_over_cdp(CDP_URL)
ctx = browser.contexts[0]
page = ctx.new_page()

# Navigate to establish auth context
page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
page.wait_for_timeout(5000)
print(f"Page title: {page.title()}")

# Discover writable fields
fields_result = page.evaluate(f"""async () => {{
    const resp = await fetch(
        "{SP_API}/fields?$filter=Hidden eq false and ReadOnlyField eq false&$select=Title,InternalName,TypeAsString",
        {{ headers: {{ "Accept": "application/json;odata=nometadata" }} }}
    );
    return await resp.json();
}}""")

print("\\nWritable fields:")
for f in fields_result.get("value", []):
    print(f"  {f['InternalName']:30s} ({f['TypeAsString']:15s}) - {f['Title']}")

# Get total item count
count_result = page.evaluate(f"""async () => {{
    const resp = await fetch(
        "{SP_API}/ItemCount",
        {{ headers: {{ "Accept": "application/json;odata=nometadata" }} }}
    );
    return await resp.json();
}}""")
print(f"\\nTotal items: {count_result.get('value', 'unknown')}")

# Fetch sample items to verify field names
sample_result = page.evaluate(f"""async () => {{
    const resp = await fetch(
        "{SP_API}/items?$top=5&$orderby=Id desc&$select=Id,Title,Link,{F_ABSTRACT},{F_DURATION},{F_PUBLISHED},{F_YT_ID}",
        {{ headers: {{ "Accept": "application/json;odata=nometadata" }} }}
    );
    return await resp.json();
}}""")

print("\nSample items (latest 5):")
for item in sample_result.get("value", []):
    print(f"  ID={item.get('Id')} Title={str(item.get('Title',''))[:60]}")
    link = item.get('Link', {})
    if isinstance(link, dict):
        print(f"    Link.Url={link.get('Url', 'N/A')}")
    else:
        print(f"    Link={link}")
    print(f"    Abstract={str(item.get(F_ABSTRACT,''))[:100]}")
    print(f"    Duration={item.get(F_DURATION,'')}")
    print(f"    Published={item.get(F_PUBLISHED,'')}")
    print(f"    yt_ID={item.get(F_YT_ID,'')}")

# Fetch ALL items using ID-based pagination
print("\nFetching all items with ID-based pagination...")
all_items = []
last_id = 0
while True:
    batch = page.evaluate(f"""async () => {{
        const resp = await fetch(
            "{SP_API}/items?$top=500&$filter=Id gt {last_id}&$select=Id,Title,Link,{F_ABSTRACT},{F_YT_ID}&$orderby=Id",
            {{ headers: {{ "Accept": "application/json;odata=nometadata" }} }}
        );
        return await resp.json();
    }}""")
    items = batch.get("value", [])
    if not items:
        break
    all_items.extend(items)
    last_id = items[-1]["Id"]
    print(f"  Fetched {len(all_items)} items (last ID={last_id})...")
    if len(items) < 500:
        break

print(f"Total fetched: {len(all_items)}")

# Save for matching
sp_data = []
for item in all_items:
    link = item.get("Link", {})
    url = link.get("Url", "") if isinstance(link, dict) else str(link or "")
    sp_data.append({
        "id": item.get("Id"),
        "title": item.get("Title", ""),
        "url": url,
        "abstract": item.get(F_ABSTRACT, "") or "",
        "yt_id": item.get(F_YT_ID, "") or "",
    })

out_file = os.path.join(BASE, "sp_videoposts.json")
with open(out_file, "w", encoding="utf-8") as f:
    json.dump(sp_data, f, ensure_ascii=False, indent=2)
print(f"Saved {len(sp_data)} items to sp_videoposts.json")

page.close()
p.stop()
