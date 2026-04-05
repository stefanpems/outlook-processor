"""
Create a new BlogPosts item in SharePoint via REST API.
Connects to Edge via CDP to reuse auth context.

Usage: python pipeline_sp_create.py <json_file>
  The JSON file must contain: title, published_date, summary, topic, tech, blog_link

Output: JSON to stdout with ok, id, title, link_warning.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SP_API = CONFIG["sharepoint"]["blog_list_api"]
SP_ENTITY_TYPE = CONFIG["sharepoint"]["blog_list_entity_type"]
SP_SITE_BASE = CONFIG["sharepoint"]["site_base"]
SOURCE_MAP = {k: int(v) for k, v in CONFIG["source_map"].items()}
TECH_MAP_IDS = {k: int(v) for k, v in CONFIG["tech_map"].items()}


def get_tech_ids(tech_str):
    """Parse comma-separated tech string and return list of SP lookup IDs."""
    if not tech_str:
        return []
    parts = [t.strip() for t in tech_str.split(",") if t.strip()]
    ids = []
    for t in parts:
        tid = TECH_MAP_IDS.get(t)
        if tid:
            ids.append(tid)
        else:
            print(f"  WARNING: Unknown tech '{t}', skipping", file=sys.stderr)
    return ids


def get_digest(page):
    """Get SharePoint form digest token."""
    sp_base = SP_SITE_BASE.rstrip('/')
    return page.evaluate(f"""async () => {{
        const resp = await fetch(
            "{sp_base}/_api/contextinfo",
            {{ method: "POST", headers: {{ "Accept": "application/json;odata=nometadata" }} }}
        );
        return (await resp.json()).FormDigestValue;
    }}""")


def create_sp_item(page, digest, item_data):
    """Create a new item in the BlogPosts SP list.
    Returns dict with ok, id, title, link_warning."""
    published = item_data.get("published_date", "").replace("-", ".")
    title = item_data.get("title", "")
    topic = item_data.get("topic", "")
    tech_str = item_data.get("tech", "")
    link = item_data.get("blog_link", "")
    summary = item_data.get("summary", "")

    source_id = SOURCE_MAP.get(topic)
    tech_ids = get_tech_ids(tech_str)

    body = {
        "__metadata": {"type": SP_ENTITY_TYPE},
        "Title": title,
        "field_0": published,
        "Summary": summary,
        "Notes": "by Agent",
    }
    if source_id:
        body["SourceNewId"] = source_id
    if tech_ids:
        body["TechId"] = {"results": tech_ids}

    body_json = json.dumps(body)
    body_esc = body_json.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

    result = page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}/items",
                {{
                    method: "POST",
                    headers: {{
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": `{digest}`
                    }},
                    body: `{body_esc}`
                }}
            );
            if (!resp.ok) {{
                const txt = await resp.text();
                return {{ ok: false, status: resp.status, error: txt.substring(0, 500) }};
            }}
            const data = await resp.json();
            return {{ ok: true, id: data.d.Id, title: data.d.Title }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")

    # If created, update Link field via MERGE (SP.FieldUrlValue can't be set in POST)
    if result.get("ok") and link:
        new_id = result["id"]
        link_body = json.dumps({
            "__metadata": {"type": SP_ENTITY_TYPE},
            "Link": {
                "__metadata": {"type": "SP.FieldUrlValue"},
                "Url": link,
                "Description": link
            }
        })
        link_esc = link_body.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
        link_result = page.evaluate(f"""async () => {{
            try {{
                const resp = await fetch(
                    "{SP_API}/items({new_id})",
                    {{
                        method: "POST",
                        headers: {{
                            "Accept": "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": `{digest}`,
                            "IF-MATCH": "*",
                            "X-HTTP-Method": "MERGE"
                        }},
                        body: `{link_esc}`
                    }}
                );
                return {{ ok: resp.ok, status: resp.status }};
            }} catch(e) {{
                return {{ ok: false, error: e.message }};
            }}
        }}""")
        if not link_result.get("ok"):
            result["link_warning"] = f"Link update failed: {link_result}"

    return result


def delete_sp_item(page, digest, sp_id):
    """Delete an SP item by ID."""
    result = page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}/items({sp_id})",
                {{
                    method: "POST",
                    headers: {{
                        "Accept": "application/json;odata=verbose",
                        "X-RequestDigest": `{digest}`,
                        "IF-MATCH": "*",
                        "X-HTTP-Method": "DELETE"
                    }}
                }}
            );
            return {{ ok: resp.ok, status: resp.status }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")
    return result


def update_sp_summary(page, digest, sp_id, summary):
    """Update the Summary field on an existing SP item via MERGE."""
    summary_esc = json.dumps({
        "__metadata": {"type": SP_ENTITY_TYPE},
        "Summary": summary
    }).replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

    result = page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}/items({sp_id})",
                {{
                    method: "POST",
                    headers: {{
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": `{digest}`,
                        "IF-MATCH": "*",
                        "X-HTTP-Method": "MERGE"
                    }},
                    body: `{summary_esc}`
                }}
            );
            return {{ ok: resp.ok, status: resp.status }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")
    return result


def main():
    # Parse arguments
    mode = "create"  # default
    sp_id_to_update = None
    delete_ids = []

    args = sys.argv[1:]
    if len(args) >= 2 and args[0] == "--update-summary":
        mode = "update-summary"
        sp_id_to_update = int(args[1])
        args = args[2:]
    elif len(args) >= 2 and args[0] == "--delete":
        mode = "delete"
        delete_ids = [int(x) for x in args[1].split(",")]
        args = args[2:]

    if args and args[0] != "-":
        item_data = json.load(open(args[0], encoding="utf-8"))
    elif mode != "delete":
        item_data = json.load(sys.stdin)
    else:
        item_data = {}

    p = sync_playwright().start()
    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        sp_page = ctx.new_page()
        sp_page.goto(
            CONFIG["sharepoint"]["blog_list_url"],
            wait_until="domcontentloaded", timeout=30000
        )
        sp_page.wait_for_timeout(3000)
        digest = get_digest(sp_page)

        if mode == "update-summary":
            summary = item_data.get("summary", "")
            print(f"Updating summary on SP item {sp_id_to_update}: {item_data.get('title','')[:70]}", file=sys.stderr)
            result = update_sp_summary(sp_page, digest, sp_id_to_update, summary)
            result["id"] = sp_id_to_update
            result["mode"] = "update-summary"
        elif mode == "delete":
            results = []
            for did in delete_ids:
                print(f"Deleting SP item {did}...", file=sys.stderr)
                r = delete_sp_item(sp_page, digest, did)
                r["id"] = did
                results.append(r)
            result = {"mode": "delete", "results": results}
        else:
            print(f"Creating SP item: {item_data.get('title','')[:70]}", file=sys.stderr)
            result = create_sp_item(sp_page, digest, item_data)

        sp_page.close()
        print(json.dumps(result, indent=2, ensure_ascii=False))

    finally:
        p.stop()


if __name__ == "__main__":
    main()
