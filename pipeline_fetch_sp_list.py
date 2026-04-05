"""Fetch the Technology reference list from SharePoint via Edge CDP + REST API."""
import sys, json, os
sys.stdout = open(sys.stdout.fileno(), mode='w', buffering=1)
from playwright.sync_api import sync_playwright

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SP_TECH_LIST_URL = CONFIG["sharepoint"]["tech_list_url"]
SP_TECH_LIST_API = CONFIG["sharepoint"]["tech_list_api"]

p = sync_playwright().start()
browser = p.chromium.connect_over_cdp(CDP_URL)
ctx = browser.contexts[0]

page = ctx.new_page()
# Navigate to the list page first to establish the SharePoint auth context
page.goto(SP_TECH_LIST_URL, wait_until="domcontentloaded", timeout=30000)
page.wait_for_timeout(5000)
print(f"Page title: {page.title()}")

# Use the SharePoint REST API from the page context (inherits auth cookies)
# Derive the relative site path from config
from urllib.parse import urlparse as _urlparse
_sp_site_path = _urlparse(CONFIG["sharepoint"]["site_base"]).path.rstrip('/')

# First, find the correct list name
lists_result = page.evaluate(f"""async () => {{
    const resp = await fetch(
        "{_sp_site_path}/_api/web/lists?$select=Title,ItemCount&$filter=Hidden eq false",
        {{ headers: {{ "Accept": "application/json;odata=nometadata" }} }}
    );
    return await resp.json();
}}""")

print("Available lists:")
for lst in lists_result.get("value", []):
    print(f"  {lst['Title']} ({lst['ItemCount']} items)")

# Try to find the right list name
tech_list_name = None
for lst in lists_result.get("value", []):
    if "tech" in lst["Title"].lower():
        tech_list_name = lst["Title"]
        if "old" not in lst["Title"].lower():
            print(f"\nUsing list: {tech_list_name} ({lst['ItemCount']} items)")
            break

if tech_list_name:
    encoded_name = tech_list_name.replace("'", "''")
    result2 = page.evaluate(f"""async () => {{
        const resp = await fetch(
            "{_sp_site_path}/_api/web/lists/getbytitle('{encoded_name}')/items?$select=Title,Id&$top=500&$orderby=Title",
            {{ headers: {{ "Accept": "application/json;odata=nometadata" }} }}
        );
        return await resp.json();
    }}""")
    techs = [item.get("Title", "") for item in result2.get("value", []) if item.get("Title")]
else:
    techs = []
    print("ERROR: Could not find technology reference list")

print(f"\nFound {len(techs)} technologies via REST API:")
for t in techs:
    print(f"  {t}")

with open(os.path.join(BASE, "tech_list.json"), "w", encoding="utf-8") as f:
    json.dump(techs, f, indent=2, ensure_ascii=False)
print(f"\nSaved to tech_list.json")

page.close()
p.stop()
