"""
Categorize and move Viva Engage notification emails in Outlook Web via Playwright CDP.
For each conversation title, searches the entire mailbox by subject, categorizes as
"By agent - Viva Engage", and moves to "Social Networks" folder.

Usage:
  python pipeline_ve_email_actions.py <title>
  python pipeline_ve_email_actions.py --titles-file <json_file>

The JSON file must be a list of strings (conversation titles).

Output: JSON to stdout with per-title results.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
CATEGORY = CONFIG["viva_engage"]["processed_category"]
TARGET_FOLDER = CONFIG["viva_engage"]["target_folder"]


def do_search(page, title):
    """Search Outlook for emails with the given title in the subject.
    Searches entire mailbox (all folders). Returns count of results."""
    safe = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014\u2011]', ' ', title)
    safe = re.sub(r'\s+', ' ', safe).strip()
    query = f'subject:("{safe}") -"By agent - Viva Engage" -"PescoPedia"'

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
    sb.fill(query)
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(5000)

    rows = page.query_selector_all("div[role='option']")
    return len(rows)


def has_category(page):
    """Check if the currently displayed email in the reading pane has the processed category."""
    try:
        rp_text = page.locator('[role="main"]').first.inner_text()
        return CATEGORY.lower() in rp_text.lower()
    except Exception:
        return False


def get_current_row_folder(page):
    """Get folder name from the currently selected row's preview text."""
    row = page.query_selector("div[role='option'][aria-selected='true']")
    if not row:
        rows = page.query_selector_all("div[role='option']")
        row = rows[0] if rows else None
    if not row:
        return ""
    try:
        txt = row.inner_text().strip()
        lines = [l.strip() for l in txt.split("\n") if l.strip()]
        if lines:
            last_line = lines[-1]
            parts = [p.strip() for p in last_line.replace("\t", "|").split("|") if p.strip()]
            if parts:
                return parts[-1]
    except Exception:
        pass
    return ""


def do_categorize_one(page, row):
    """Categorize a single email row via right-click context menu (Italian labels)."""
    try:
        box = row.bounding_box()
        if not box:
            return False
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        page.wait_for_timeout(400)
        page.mouse.click(
            box["x"] + box["width"] / 2,
            box["y"] + box["height"] / 2,
            button="right"
        )
    except Exception:
        return False
    page.wait_for_timeout(1500)

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


def do_move_one(page, row):
    """Move a single email row to the target folder via top toolbar Sposta button."""
    try:
        box = row.bounding_box()
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            page.wait_for_timeout(400)
    except Exception:
        return False

    go_home_tab(page)
    page.wait_for_timeout(500)

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

    page.keyboard.press("ArrowDown")
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(2000)
    return True


def process_title(page, title):
    """Search, categorize, and move all emails matching a single title."""
    print(f"  Searching: {title[:70]}", file=sys.stderr)
    count = do_search(page, title)

    result = {
        "title": title,
        "emails_found": count,
        "categorized_count": 0,
        "moved_count": 0,
    }

    if count == 0:
        return result

    cat_count = 0
    mov_count = 0
    processed = 0

    while True:
        rows = page.query_selector_all("div[role='option']")
        if not rows:
            break

        row = rows[0]

        # Click to select and load reading pane
        try:
            box = row.bounding_box()
            if box:
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                page.wait_for_timeout(1500)
        except Exception:
            break

        already_cat = has_category(page)
        row_folder = get_current_row_folder(page)
        in_target = TARGET_FOLDER.lower() in row_folder.lower()

        if already_cat and in_target:
            # Shouldn't appear due to -"By agent - Viva Engage" filter, but safety check
            break

        if not already_cat:
            cat_ok = do_categorize_one(page, row)
            if cat_ok:
                cat_count += 1
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)

        if not in_target:
            # Re-fetch rows since categorize may have changed DOM
            rows = page.query_selector_all("div[role='option']")
            if rows:
                mov_ok = do_move_one(page, rows[0])
                if mov_ok:
                    mov_count += 1
                page.wait_for_timeout(2000)

        processed += 1
        if processed >= count:
            break

    result["categorized_count"] = cat_count
    result["moved_count"] = mov_count
    return result


def main():
    titles = []

    if len(sys.argv) >= 3 and sys.argv[1] == "--titles-file":
        titles = json.load(open(sys.argv[2], encoding="utf-8"))
    elif len(sys.argv) >= 2:
        titles = [sys.argv[1]]
    else:
        print("Usage: python pipeline_ve_email_actions.py <title>")
        print("       python pipeline_ve_email_actions.py --titles-file <json_file>")
        sys.exit(1)

    print(f"Processing {len(titles)} conversation title(s)", file=sys.stderr)

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

        for title in titles:
            r = process_title(page, title)
            results.append(r)
            total_found += r["emails_found"]
            total_cat += r["categorized_count"]
            total_mov += r["moved_count"]

        output = {
            "total_titles": len(titles),
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
            "total_titles": len(titles),
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
