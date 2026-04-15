"""Create a new item in the SharePoint VideosMSInt list via REST API.
Connects to Edge via CDP to reuse auth context.

Usage:
  python pipeline_teams_sp_create.py <input.json>

Input JSON fields:
  title, published_date, summary, tech, duration (or duration_formatted),
  sha256_id, video_link, meeting_sender (optional, maps to SourceNewId via source_map)

Output: JSON to stdout with ok, id, title, link_warning, link_skipped.
"""
import json, re, os, sys
from urllib.parse import unquote
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
TM_CFG = CONFIG["teams_meeting"]
SP_API = TM_CFG["list_api"]
SP_ENTITY_TYPE = TM_CFG["list_entity_type"]
SP_SITE_BASE = CONFIG["sharepoint"]["site_base"]
SP_LIST_URL = TM_CFG["list_url"]
FIELDS = TM_CFG["fields"]
TECH_MAP_IDS = {k: int(v) for k, v in CONFIG["tech_map"].items()}
SOURCE_MAP = {k: int(v) for k, v in CONFIG["source_map"].items()}


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
    sp_base = SP_SITE_BASE.rstrip("/")
    return page.evaluate(f"""async () => {{
        const resp = await fetch(
            "{sp_base}/_api/contextinfo",
            {{ method: "POST", headers: {{ "Accept": "application/json;odata=nometadata" }} }}
        );
        return (await resp.json()).FormDigestValue;
    }}""")


def discover_entity_type(page):
    """Auto-discover the list entity type from SP REST API."""
    result = page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}?$select=ListItemEntityTypeFullName",
                {{ headers: {{ "Accept": "application/json;odata=verbose" }} }}
            );
            if (!resp.ok) return null;
            const data = await resp.json();
            return data.d.ListItemEntityTypeFullName;
        }} catch(e) {{ return null; }}
    }}""")
    return result


def check_dup_by_sha256(page, sha256_id):
    """Check if an item with the same ID_SHA256 already exists."""
    sha_field = FIELDS["sha256_id"]
    result = page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}/items?$filter={sha_field} eq '{sha256_id}'&$select=Id,Title&$top=1",
                {{ headers: {{ "Accept": "application/json;odata=verbose" }} }}
            );
            if (!resp.ok) return {{ found: false, error: resp.status }};
            const data = await resp.json();
            const items = data.d.results || [];
            if (items.length > 0) {{
                return {{ found: true, id: items[0].Id, title: items[0].Title }};
            }}
            return {{ found: false }};
        }} catch(e) {{
            return {{ found: false, error: e.message }};
        }}
    }}""")
    return result


def create_sp_item(page, digest, entity_type, item_data):
    """Create a new item in the VideosMSInt SP list."""
    published = item_data.get("published_date", "").replace("-", ".")
    title = item_data.get("title", "")
    tech_str = item_data.get("tech", "")
    link = unquote(item_data.get("video_link", ""))  # SP URL field rejects encoded chars
    summary = item_data.get("summary", "")
    duration = item_data.get("duration", "") or item_data.get("duration_formatted", "")
    sha256_id = item_data.get("sha256_id", "")
    meeting_sender = item_data.get("meeting_sender", "")

    tech_ids = get_tech_ids(tech_str)
    source_id = SOURCE_MAP.get(meeting_sender) if meeting_sender else None

    body = {
        "__metadata": {"type": entity_type},
        "Title": title,
        FIELDS["published"]: published,
        FIELDS["summary"]: summary,
        FIELDS["duration"]: duration,
        FIELDS["sha256_id"]: sha256_id,
    }
    if source_id:
        body["SourceNewId"] = source_id
    elif meeting_sender:
        print(f"  WARNING: Unknown meeting_sender '{meeting_sender}', skipping SourceNewId", file=sys.stderr)
    # If URL > 255 chars, store full URL in LongLink (multi-line text field)
    if len(link) > 255 and "long_link" in FIELDS:
        body[FIELDS["long_link"]] = link
    if tech_ids:
        body["TechId"] = {"results": tech_ids}

    body_json = json.dumps(body, ensure_ascii=False)
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
    # Skip if URL > 255 chars — SP Hyperlink field has a 255-char limit;
    # the full URL is already stored in LongLink during the POST above.
    if result.get("ok") and link and len(link) <= 255:
        new_id = result["id"]
        link_body = json.dumps(
            {
                "__metadata": {"type": entity_type},
                "Link": {
                    "__metadata": {"type": "SP.FieldUrlValue"},
                    "Url": link,
                    "Description": link,
                },
            },
            ensure_ascii=False,
        )
        link_esc = (
            link_body.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
        )
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
    elif result.get("ok") and link and len(link) > 255:
        result["link_skipped"] = f"URL too long ({len(link)} chars) for SP Hyperlink field (max 255). Stored in LongLink only."

    return result


def fix_link(page, digest, entity_type, item_id, link):
    """Retry setting the Link field on an existing SP item via MERGE."""
    link_body = json.dumps(
        {
            "__metadata": {"type": entity_type},
            "Link": {
                "__metadata": {"type": "SP.FieldUrlValue"},
                "Url": link,
                "Description": link,
            },
        },
        ensure_ascii=False,
    )
    link_esc = link_body.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
    return page.evaluate(f"""async () => {{
        try {{
            const resp = await fetch(
                "{SP_API}/items({item_id})",
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
            if (!resp.ok) {{
                const txt = await resp.text();
                return {{ ok: false, status: resp.status, error: txt.substring(0, 500) }};
            }}
            return {{ ok: true, status: resp.status }};
        }} catch(e) {{
            return {{ ok: false, error: e.message }};
        }}
    }}""")


def main():
    args = sys.argv[1:]
    if not args:
        print("Usage: python pipeline_teams_sp_create.py <input.json>")
        print("       python pipeline_teams_sp_create.py --fix-link <id> <url>")
        sys.exit(1)

    # --fix-link mode: retry setting Link on an existing item
    if args[0] == "--fix-link":
        if len(args) < 3:
            print("Usage: python pipeline_teams_sp_create.py --fix-link <id> <url>")
            sys.exit(1)
        fix_id = int(args[1])
        fix_url = args[2]
        p = sync_playwright().start()
        try:
            ensure_edge_cdp()
            browser = p.chromium.connect_over_cdp(CDP_URL)
            ctx = browser.contexts[0]
            sp_page = ctx.new_page()
            sp_page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
            sp_page.wait_for_timeout(3000)
            entity_type = discover_entity_type(sp_page) or SP_ENTITY_TYPE
            digest = get_digest(sp_page)
            result = fix_link(sp_page, digest, entity_type, fix_id, fix_url)
            sp_page.close()
            print(json.dumps(result, indent=2, ensure_ascii=False))
        finally:
            p.stop()
        return

    item_data = json.load(open(args[0], encoding="utf-8"))

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        sp_page = ctx.new_page()
        sp_page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
        sp_page.wait_for_timeout(3000)

        # Auto-discover entity type (fallback to config value)
        entity_type = discover_entity_type(sp_page) or SP_ENTITY_TYPE
        if entity_type != SP_ENTITY_TYPE:
            print(
                f"  Entity type discovered: {entity_type} (config had: {SP_ENTITY_TYPE})",
                file=sys.stderr,
            )

        # Check for duplicate by SHA256
        sha256_id = item_data.get("sha256_id", "")
        if sha256_id:
            dup = check_dup_by_sha256(sp_page, sha256_id)
            if dup.get("found"):
                print(
                    json.dumps(
                        {
                            "ok": False,
                            "duplicate": True,
                            "existing_id": dup["id"],
                            "existing_title": dup.get("title", ""),
                            "sha256_id": sha256_id,
                        },
                        indent=2,
                        ensure_ascii=False,
                    )
                )
                sp_page.close()
                return

        digest = get_digest(sp_page)
        print(f"Creating SP item: {item_data.get('title', '')[:70]}", file=sys.stderr)
        result = create_sp_item(sp_page, digest, entity_type, item_data)
        sp_page.close()
        print(json.dumps(result, indent=2, ensure_ascii=False))

    finally:
        p.stop()


if __name__ == "__main__":
    main()
