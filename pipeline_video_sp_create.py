"""
Create a new VideoPosts item in SharePoint via REST API.
Connects to Edge via CDP to reuse auth context.

Usage:
  echo '{"title":"...","published_date":"...","abstract":"...","topic":"...","tech":"...","video_link":"...","duration":"...","yt_id":"..."}' | python pipeline_video_sp_create.py -
  echo '{"abstract":"...","title":"..."}' | python pipeline_video_sp_create.py --update-abstract <sp_id> -

Output: JSON to stdout with ok, id, title, link_warning.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SP_API = CONFIG["video_sharepoint"]["list_api"]
SP_ENTITY_TYPE = CONFIG["video_sharepoint"]["list_entity_type"]
SP_SITE_BASE = CONFIG["sharepoint"]["site_base"]
SP_LIST_URL = CONFIG["video_sharepoint"]["list_url"]
FIELDS = CONFIG["video_sharepoint"]["fields"]
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
    """Create a new item in the VideoPosts SP list."""
    published = item_data.get("published_date", "").replace("-", ".")
    title = item_data.get("title", "")
    topic = item_data.get("topic", "")
    tech_str = item_data.get("tech", "")
    link = item_data.get("video_link", "")
    abstract = item_data.get("abstract", "")
    duration = item_data.get("duration", "")
    yt_id = item_data.get("yt_id", "")

    source_id = SOURCE_MAP.get(topic)
    tech_ids = get_tech_ids(tech_str)

    body = {
        "__metadata": {"type": SP_ENTITY_TYPE},
        "Title": title,
        FIELDS["published"]: published,
        FIELDS["abstract"]: abstract,
        FIELDS["duration"]: duration,
    }
    if yt_id:
        body[FIELDS["yt_id"]] = yt_id
    if source_id:
        body["SourceId"] = source_id
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

    # Update Link field via MERGE (SP.FieldUrlValue can't be set in POST)
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


def update_sp_abstract(page, digest, sp_id, abstract):
    """Update the Abstract field on an existing SP item via MERGE."""
    body = {
        "__metadata": {"type": SP_ENTITY_TYPE},
        FIELDS["abstract"]: abstract
    }
    body_esc = json.dumps(body).replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

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
                    body: `{body_esc}`
                }}
            );
            return {{ ok: resp.ok, status: resp.status }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")
    return result


def update_sp_all_fields(page, digest, sp_id, item_data):
    """Update ALL fields on an existing SP item via MERGE (reprocess mode).
    Also updates Link via a separate MERGE (SP.FieldUrlValue)."""
    published = item_data.get("published_date", "").replace("-", ".")
    title = item_data.get("title", "")
    topic = item_data.get("topic", "")
    tech_str = item_data.get("tech", "")
    link = item_data.get("video_link", "")
    abstract = item_data.get("abstract", "")
    duration = item_data.get("duration", "")
    yt_id = item_data.get("yt_id", "")

    source_id = SOURCE_MAP.get(topic)
    tech_ids = get_tech_ids(tech_str)

    body = {
        "__metadata": {"type": SP_ENTITY_TYPE},
        "Title": title,
        FIELDS["published"]: published,
        FIELDS["abstract"]: abstract,
        FIELDS["duration"]: duration,
    }
    if yt_id:
        body[FIELDS["yt_id"]] = yt_id
    if source_id:
        body["SourceId"] = source_id
    if tech_ids:
        body["TechId"] = {"results": tech_ids}

    body_esc = json.dumps(body).replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")

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
                    body: `{body_esc}`
                }}
            );
            return {{ ok: resp.ok, status: resp.status }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")

    # Update Link field separately (SP.FieldUrlValue)
    if result.get("ok", True) and link:
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


def main():
    mode = "create"
    sp_id_to_update = None

    args = sys.argv[1:]
    if len(args) >= 2 and args[0] == "--update-abstract":
        mode = "update-abstract"
        sp_id_to_update = int(args[1])
        args = args[2:]
    elif len(args) >= 2 and args[0] == "--update-all":
        mode = "update-all"
        sp_id_to_update = int(args[1])
        args = args[2:]

    if args and args[0] != "-":
        item_data = json.load(open(args[0], encoding="utf-8"))
    else:
        item_data = json.load(sys.stdin)

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        sp_page = ctx.new_page()
        sp_page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
        sp_page.wait_for_timeout(3000)
        digest = get_digest(sp_page)

        if mode == "update-abstract":
            abstract = item_data.get("abstract", "")
            print(f"Updating abstract on SP item {sp_id_to_update}: {item_data.get('title','')[:70]}", file=sys.stderr)
            result = update_sp_abstract(sp_page, digest, sp_id_to_update, abstract)
            result["id"] = sp_id_to_update
            result["mode"] = "update-abstract"
        elif mode == "update-all":
            print(f"Updating all fields on SP item {sp_id_to_update}: {item_data.get('title','')[:70]}", file=sys.stderr)
            result = update_sp_all_fields(sp_page, digest, sp_id_to_update, item_data)
            result["id"] = sp_id_to_update
            result["mode"] = "update-all"
        else:
            print(f"Creating SP item: {item_data.get('title','')[:70]}", file=sys.stderr)
            result = create_sp_item(sp_page, digest, item_data)

        sp_page.close()
        print(json.dumps(result, indent=2, ensure_ascii=False))

    finally:
        p.stop()


if __name__ == "__main__":
    main()
