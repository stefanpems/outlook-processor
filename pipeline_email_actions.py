"""
Categorize and/or move a single email in Outlook Web via Playwright CDP.
Reuses the tested patterns from legacy_batch_all.py for safe email selection,
categorization (Italian labels), and move via keyboard navigation.

Usage:
  python pipeline_email_actions.py categorize <email_title>
  python pipeline_email_actions.py move <email_title>
  python pipeline_email_actions.py both <email_title>

Output: JSON to stdout with categorized, moved, emails_found.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp, ensure_all_folders_scope

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SENDER = CONFIG["outlook"]["sender"]
SUBJECT_PREFIX = CONFIG["outlook"]["subject_prefix"]
CATEGORY = CONFIG["outlook"]["processed_category"]
TARGET_FOLDER = CONFIG["outlook"]["target_folder"]
EXCLUDE_TERMS = CONFIG["outlook"].get("exclude_terms", [])


INCLUDE_PROCESSED = "--include-processed" in sys.argv


def do_search(page, title):
    """Search Outlook for a specific blog email. Returns count of results."""
    safe = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014]', ' ', title)
    safe = re.sub(r'\s+', ' ', safe).strip()
    words = safe.split()
    short = " ".join(words[:6]) if len(words) > 6 else safe
    query = f'from:{SENDER} subject:({SUBJECT_PREFIX} {short})'
    if not INCLUDE_PROCESSED:
        query += ' -"By agent - Blog"'
    for term in EXCLUDE_TERMS:
        query += f' -{term}'

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

    ensure_all_folders_scope(page)

    rows = page.query_selector_all("div[role='option']")
    return len(rows)


def select_blog_emails(page):
    """Select only email rows whose subject contains [Blog-]. NEVER Ctrl+A."""
    rows = page.query_selector_all("div[role='option']")
    if not rows:
        return []

    matched = []
    for row in rows:
        try:
            text = row.inner_text()
        except Exception:
            continue
        if "[Blog-" not in text:
            continue
        matched.append(row)

    if not matched:
        return []

    # Click first match
    try:
        matched[0].click(timeout=5000)
    except Exception:
        box = matched[0].bounding_box()
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        else:
            return []
    page.wait_for_timeout(300)

    # Ctrl+click remaining matches
    for row in matched[1:]:
        try:
            box = row.bounding_box()
            if box:
                page.mouse.click(
                    box["x"] + box["width"] / 2,
                    box["y"] + box["height"] / 2,
                    modifiers=["Control"]
                )
                page.wait_for_timeout(200)
        except Exception:
            continue
    page.wait_for_timeout(500)
    return matched


def has_category(page):
    """Check if the currently displayed email in the reading pane has the processed category."""
    try:
        rp_text = page.locator('[role="main"]').first.inner_text()
        return CATEGORY.lower() in rp_text.lower()
    except Exception:
        return False


def get_current_row_folder(page):
    """Get folder name from the currently selected row's preview text.
    In search results, the folder appears as the last segment of the row text."""
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
        # Single-click first to ensure only this row is selected
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        page.wait_for_timeout(400)
        # Right-click for context menu
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


def do_move_one(page, row=None):
    """Move a single email row to the target folder via top toolbar Sposta button.
    If row is None, moves the currently selected email without re-clicking."""
    if row is not None:
        try:
            box = row.bounding_box()
            if box:
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                page.wait_for_timeout(400)
        except Exception:
            return False

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
        # Fallback: look for aria-haspopup buttons
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


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) < 2:
        print("Usage: python pipeline_email_actions.py <categorize|move|both> <email_title> [--include-processed]")
        sys.exit(1)

    action = args[0].lower()
    title = args[1]

    print(f"Processing email: {title[:70]}", file=sys.stderr)

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]

        # Find or create Outlook tab
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

        count = do_search(page, title)
        result = {
            "title": title,
            "emails_found": count,
            "categorized": False,
            "moved": False,
        }

        if count == 0:
            result["error"] = "No emails found"
            print(json.dumps(result, indent=2))
            return

        matched = select_blog_emails(page)
        if not matched:
            result["error"] = "No [Blog-] emails in results"
            print(json.dumps(result, indent=2))
            return

        result["blog_emails"] = len(matched)

        # Process emails one at a time, re-searching after each to avoid
        # stale DOM / stale reading-pane issues in virtualized list.
        cat_count = 0
        mov_count = 0
        max_iter = result["blog_emails"] + 5  # safety limit

        for iteration in range(max_iter):
            if iteration > 0:
                # Re-search from scratch for fresh results
                page.wait_for_timeout(2000)
                count = do_search(page, title)
                if count == 0:
                    break

            rows = page.query_selector_all("div[role='option']")
            blog_rows = [r for r in rows if "[Blog-" in (r.inner_text() or "")]
            if not blog_rows:
                break

            row = blog_rows[0]

            # Click to select and load reading pane
            try:
                box = row.bounding_box()
                if box:
                    page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                    page.wait_for_timeout(2500)
            except Exception:
                continue

            # Dual check: skip if already categorized AND in target folder
            already_cat = has_category(page)
            row_folder = get_current_row_folder(page)
            in_target = TARGET_FOLDER.lower() in row_folder.lower()

            needs_cat = action in ("categorize", "both") and not already_cat
            needs_move = action in ("move", "both") and not in_target

            if not needs_cat and not needs_move:
                print(f"  Skip (already done): cat={already_cat} folder={row_folder}", file=sys.stderr)
                break  # search excludes processed; if first result is done, stop

            if needs_cat:
                cat_ok = do_categorize_one(page, row)
                if cat_ok:
                    cat_count += 1
                page.keyboard.press("Escape")
                page.wait_for_timeout(500)

            if needs_move:
                # Move the currently selected email directly — do NOT re-fetch
                # rows, because after categorize the email may have vanished
                # from filtered results but is still selected in the UI.
                mov_ok = do_move_one(page)
                if mov_ok:
                    mov_count += 1
                page.wait_for_timeout(2000)

        result["categorized"] = cat_count > 0
        result["moved"] = mov_count > 0
        result["categorized_count"] = cat_count
        result["moved_count"] = mov_count

        print(json.dumps(result, indent=2))

    except Exception as ex:
        # Try to recover
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(300)
            page.keyboard.press("Escape")
        except Exception:
            pass
        error_result = {
            "title": title,
            "error": str(ex),
            "categorized": False,
            "moved": False,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    finally:
        p.stop()


if __name__ == "__main__":
    main()
