"""
Categorize and move Viva Engage notification emails in Outlook Web via Playwright CDP.
For a given notification title, searches by sender + subject + keywords in the entire
mailbox, selects ALL results with Ctrl+A, categorizes as 'By agent - Viva Engage',
and moves to 'Social Networks' folder.

Usage:
  python ve-notifications-email-actions.py "<notification_title>" "<author>" "<community_name>"
  python ve-notifications-email-actions.py --batch-file <json_file>

The JSON file must be a list of objects with keys: notification_title, author, community_name.

Output: JSON to stdout with per-item results.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
VE_CFG = CONFIG["viva_engage"]
SENDER = VE_CFG["notification_sender"]
CATEGORY = VE_CFG["processed_category"]
TARGET_FOLDER = VE_CFG["target_folder"]


def do_search(page, notification_title, author, community_name):
    """Search Outlook for VE notification emails. Returns count of results."""
    safe_title = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014\u2011]', ' ', notification_title)
    safe_title = re.sub(r'\s+', ' ', safe_title).strip()
    # Truncate if too long (Outlook search has limits)
    words = safe_title.split()
    short_title = " ".join(words[:10]) if len(words) > 10 else safe_title

    safe_author = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014]', ' ', author).strip()
    safe_community = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014]', ' ', community_name).strip()

    query = f'from:{SENDER} subject:({short_title}) {safe_author} {safe_community}'

    try:
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)
    except Exception:
        page.evaluate("window.location.href = 'https://outlook.cloud.microsoft/mail/'")
        page.wait_for_timeout(4000)
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)

    sb.click()
    page.wait_for_timeout(400)
    page.keyboard.press("Control+a")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(200)
    page.keyboard.type(query, delay=5)
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(5000)

    rows = page.query_selector_all("div[role='option']")
    return len(rows)


def select_all_results(page):
    """Select ALL search results using Ctrl+A."""
    rows = page.query_selector_all("div[role='option']")
    if not rows:
        return 0

    # Click first row to set focus in the list
    try:
        box = rows[0].bounding_box()
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            page.wait_for_timeout(500)
    except Exception:
        return 0

    # Ctrl+A to select all
    page.keyboard.press("Control+a")
    page.wait_for_timeout(800)

    return len(rows)


def do_categorize(page):
    """Categorize selected emails via right-click context menu (Italian labels)."""
    rows = page.query_selector_all("div[role='option']")
    if not rows:
        return False

    # Right-click on the first selected row
    try:
        box = rows[0].bounding_box()
        if not box:
            return False
        page.mouse.click(
            box["x"] + box["width"] / 2,
            box["y"] + box["height"] / 2,
            button="right"
        )
    except Exception:
        return False
    page.wait_for_timeout(1500)

    # Find "Categorizza" menu item
    cat_btn = None
    for item in page.query_selector_all("[role='menuitem']"):
        try:
            if "categorizza" in item.inner_text().lower():
                cat_btn = item
                break
        except Exception:
            pass

    if not cat_btn:
        page.keyboard.press("Escape")
        return False

    cat_btn.evaluate("el => el.click()")
    page.wait_for_timeout(1500)

    # Find and click the category
    for sub in page.query_selector_all("[role='menuitemcheckbox'], [role='menuitem']"):
        try:
            txt = sub.inner_text().strip()
            if CATEGORY in txt:
                checked = sub.get_attribute("aria-checked")
                if checked == "true":
                    page.keyboard.press("Escape")
                    return True
                sub.evaluate("el => el.click()")
                page.wait_for_timeout(800)
                return True
        except Exception:
            pass

    page.keyboard.press("Escape")
    return False


def go_home_tab(page):
    """Click the Home tab in Outlook toolbar."""
    home_tab = page.query_selector("button[role='tab'][name='Home'], [role='tab']:has-text('Home')")
    if home_tab:
        home_tab.click()
        page.wait_for_timeout(1000)
        return True
    for el in page.query_selector_all("[role='tab']"):
        try:
            if "home" in el.inner_text().strip().lower():
                el.click()
                page.wait_for_timeout(1000)
                return True
        except Exception:
            pass
    return False


def do_move(page):
    """Move selected emails to the target folder via top toolbar 'Sposta' button."""
    go_home_tab(page)
    page.wait_for_timeout(500)

    # Find Sposta button in top toolbar (y < 200)
    sposta_btn = None
    for btn in page.query_selector_all("button"):
        try:
            if not btn.is_visible():
                continue
            box = btn.bounding_box()
            if not box or box["y"] > 200:
                continue
            txt = btn.inner_text().strip()
            al = btn.get_attribute("aria-label") or ""
            if txt == "Sposta" or (al == "Sposta" and "alto" not in al):
                sposta_btn = btn
                break
        except Exception:
            pass

    if not sposta_btn:
        for btn in page.query_selector_all("button[aria-haspopup], button[aria-expanded]"):
            try:
                if not btn.is_visible():
                    continue
                box = btn.bounding_box()
                if not box or box["y"] > 200:
                    continue
                al = btn.get_attribute("aria-label") or ""
                if "sposta" in al.lower() and "alto" not in al.lower():
                    sposta_btn = btn
                    break
            except Exception:
                pass

    if not sposta_btn:
        return False

    sposta_btn.click()
    page.wait_for_timeout(2000)

    # Find folder search input
    search_input = page.query_selector(
        "input[placeholder*='erca una cartella'], input[placeholder*='Search folder']"
    )
    if not search_input:
        for inp in page.query_selector_all("input"):
            try:
                if not inp.is_visible():
                    continue
                ph = inp.get_attribute("placeholder") or ""
                if "cartella" in ph.lower() or "folder" in ph.lower() or "cerca" in ph.lower():
                    search_input = inp
                    break
            except Exception:
                pass

    if not search_input:
        page.keyboard.press("Escape")
        return False

    search_input.click()
    page.wait_for_timeout(300)
    search_input.fill(TARGET_FOLDER)
    page.wait_for_timeout(3000)

    # Select first result via keyboard
    page.keyboard.press("ArrowDown")
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(2000)
    return True


def process_item(page, notification_title, author, community_name):
    """Search, select all, categorize, and move emails for one notification."""
    print(f"  Searching: {notification_title[:70]}", file=sys.stderr)
    count = do_search(page, notification_title, author, community_name)

    result = {
        "notification_title": notification_title,
        "author": author,
        "community_name": community_name,
        "emails_found": count,
        "categorized": False,
        "moved": False,
    }

    if count == 0:
        return result

    selected = select_all_results(page)
    print(f"  Found {count} emails, selected {selected}", file=sys.stderr)

    cat_ok = do_categorize(page)
    result["categorized"] = cat_ok
    page.keyboard.press("Escape")
    page.wait_for_timeout(500)

    # Re-select all for move (categorize may have changed focus)
    select_all_results(page)
    page.wait_for_timeout(300)

    mov_ok = do_move(page)
    result["moved"] = mov_ok
    page.wait_for_timeout(2000)

    return result


def main():
    items = []

    if len(sys.argv) >= 3 and sys.argv[1] == "--batch-file":
        items = json.load(open(sys.argv[2], encoding="utf-8"))
    elif len(sys.argv) >= 4:
        items = [{
            "notification_title": sys.argv[1],
            "author": sys.argv[2],
            "community_name": sys.argv[3],
        }]
    else:
        print("Usage: python ve-notifications-email-actions.py \"<title>\" \"<author>\" \"<community>\"")
        print("       python ve-notifications-email-actions.py --batch-file <json_file>")
        sys.exit(1)

    print(f"Processing {len(items)} item(s)", file=sys.stderr)

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]

        page = None
        for pg in ctx.pages:
            if "outlook.office.com" in pg.url or "outlook.cloud.microsoft" in pg.url:
                page = pg
                break
        if not page:
            page = ctx.new_page()
            page.goto("https://outlook.cloud.microsoft/mail/",
                       wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(5000)

        results = []
        total_found = 0
        total_cat = 0
        total_mov = 0

        for item in items:
            r = process_item(
                page,
                item["notification_title"],
                item.get("author", ""),
                item.get("community_name", ""),
            )
            results.append(r)
            total_found += r["emails_found"]
            if r["categorized"]:
                total_cat += r["emails_found"]
            if r["moved"]:
                total_mov += r["emails_found"]

        output = {
            "total_items": len(items),
            "total_emails_found": total_found,
            "total_categorized": total_cat,
            "total_moved": total_mov,
            "details": results,
        }
        print(json.dumps(output, indent=2, ensure_ascii=False))

    except Exception as ex:
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)
            page.keyboard.press("Escape")
        except Exception:
            pass
        error_result = {
            "error": str(ex),
            "total_items": len(items),
            "total_emails_found": 0,
            "total_categorized": 0,
            "total_moved": 0,
        }
        print(json.dumps(error_result, indent=2, ensure_ascii=False))
        sys.exit(1)

    finally:
        p.stop()


if __name__ == "__main__":
    main()
