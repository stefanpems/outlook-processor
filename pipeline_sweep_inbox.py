"""
Sweep all remaining uncategorized blog emails from Posta in arrivo (Inbox).

Strategy:
1. Navigate to Inbox
2. Search for blog emails scoped to "Cartella corrente" (Current folder)
3. This returns ONLY emails still in Inbox (= unprocessed ones)
4. For each email: categorize + move
5. After each move, the email disappears from Inbox results
6. Re-search after each batch. Loop until 0 results.

This avoids the broken keyboard-navigation-through-277-items approach.
"""
import json, re, os, sys, time
from playwright.sync_api import sync_playwright

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SENDER = CONFIG["outlook"]["sender"]
SUBJECT_PREFIX = CONFIG["outlook"]["subject_prefix"]
CATEGORY = CONFIG["outlook"]["processed_category"]
TARGET_FOLDER = CONFIG["outlook"]["target_folder"]
EXCLUDE_TERMS = CONFIG["outlook"].get("exclude_terms", [])

DATE_FROM = sys.argv[1] if len(sys.argv) > 1 else "2026-03-01"
DATE_TO   = sys.argv[2] if len(sys.argv) > 2 else "2026-03-31"


def connect():
    p = sync_playwright().start()
    browser = p.chromium.connect_over_cdp(CDP_URL)
    ctx = browser.contexts[0]
    page = None
    for pg in ctx.pages:
        url = pg.url.lower()
        if "outlook.office.com" in url or "outlook.cloud.microsoft" in url:
            page = pg
            break
    if not page:
        page = ctx.new_page()
        page.goto("https://outlook.cloud.microsoft/mail/",
                   wait_until="commit", timeout=30000)
        page.wait_for_timeout(5000)
    return p, browser, page


def navigate_to_inbox(page):
    """Navigate to Inbox root to ensure search scopes to it."""
    page.evaluate("window.location.href = 'https://outlook.cloud.microsoft/mail/'")
    page.wait_for_timeout(4000)


def do_search(page):
    """Search for blog emails and scroll to find ALL [Blog-] rows in Posta in arrivo.
    Returns only rows that are NOT in the target folder (i.e. still in Inbox)."""
    # Clear and fill search
    try:
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)
    except Exception:
        navigate_to_inbox(page)
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)

    sb.click()
    page.wait_for_timeout(400)
    page.keyboard.press("Control+a")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(300)

    query = f"from:{SENDER} subject:{SUBJECT_PREFIX} received:{DATE_FROM}..{DATE_TO} -\"By agent - Blog\""
    for term in EXCLUDE_TERMS:
        query += f' -{term}'
    sb.fill(query)
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(5000)

    # Try to scope search to "Cartella corrente" (Current folder)
    scope_to_current_folder(page)
    page.wait_for_timeout(2000)

    # Scroll through the virtualized list to find ALL inbox rows
    # We look for rows where the folder is NOT the target folder
    inbox_subjects = set()  # Track unique subjects found in Inbox
    all_inbox_subjects = []  # Ordered list of subjects
    all_seen_previews = set()  # Track ALL unique rows (for end-of-list detection)

    def scan_visible_rows():
        """Scan currently visible rows. Return (new_inbox, new_total) count."""
        new_inbox = 0
        new_total = 0
        rows = page.query_selector_all("div[role='option']")
        for r in rows:
            try:
                txt = r.inner_text()
                preview = txt.strip()[:200]
                if preview not in all_seen_previews:
                    all_seen_previews.add(preview)
                    new_total += 1

                if "[Blog-" not in txt:
                    continue
                # Check if the target folder appears anywhere in the row text
                # Processed emails show TARGET_FOLDER; unprocessed show "Posta in arrivo"
                if TARGET_FOLDER.lower() in txt.lower():
                    continue
                # Extract subject
                subject = ""
                for line in txt.split("\n"):
                    line = line.strip()
                    if "[Blog-" in line and len(line) > 10:
                        subject = line
                        break
                if subject and subject not in inbox_subjects:
                    inbox_subjects.add(subject)
                    all_inbox_subjects.append(subject)
                    new_inbox += 1
            except Exception:
                pass
        return new_inbox, new_total

    # Initial scan
    scan_visible_rows()

    # Scroll through the listbox to find more
    listbox = page.locator('[role="listbox"]').first
    bbox = listbox.bounding_box()
    if bbox:
        cx = bbox['x'] + bbox['width'] / 2
        cy = bbox['y'] + bbox['height'] / 2
        stable = 0
        max_scrolls = 100  # 277 items / ~5 per scroll ≈ 55 scrolls needed
        for scroll_i in range(max_scrolls):
            page.mouse.move(cx, cy)
            page.mouse.wheel(0, 400)
            page.wait_for_timeout(1500)
            new_inbox, new_total = scan_visible_rows()
            # Only stop if we're not seeing ANY new rows at all (end of list)
            if new_total == 0:
                stable += 1
            else:
                stable = 0
            if stable >= 3:
                break
            # Progress every 10 scrolls
            if (scroll_i + 1) % 10 == 0:
                print(f"    ... scrolled {scroll_i+1}x, seen {len(all_seen_previews)} total rows, {len(all_inbox_subjects)} inbox")

    print(f"  Found {len(all_inbox_subjects)} unique [Blog-] emails still in Inbox")
    for s in all_inbox_subjects:
        print(f"    - {s[:70]}")

    return all_inbox_subjects


def scope_to_current_folder(page):
    """Try to switch search scope to 'Cartella corrente' (Current folder).
    In Outlook Web, after search, there's often a scope pill/button."""
    # Look for scope buttons/links near search area
    # Common patterns: "Cartella corrente", "Current folder", "Cassetta postale corrente"
    for el in page.query_selector_all("button, [role='option'], [role='menuitemradio'], span, a"):
        try:
            if not el.is_visible():
                continue
            box = el.bounding_box()
            if not box or box["y"] > 250:  # scope controls are near the top
                continue
            txt = el.inner_text().strip().lower()
            if "cartella corrente" in txt and "cassetta" not in txt:
                el.click()
                page.wait_for_timeout(2000)
                print("  Scoped to: Cartella corrente")
                return True
        except Exception:
            pass

    # Try the search refinement area - look for "Cassetta postale corrente" and switch
    for el in page.query_selector_all("[role='tab'], [role='option'], button"):
        try:
            if not el.is_visible():
                continue
            txt = el.inner_text().strip().lower()
            if "cartella corrente" in txt:
                el.click()
                page.wait_for_timeout(2000)
                print("  Scoped to: Cartella corrente")
                return True
        except Exception:
            pass

    print("  WARNING: Could not find 'Cartella corrente' scope selector")
    return False


def find_and_click_row(page, subject):
    """Find a specific email row by subject text and click it. Returns the row or None."""
    rows = page.query_selector_all("div[role='option']")
    for r in rows:
        try:
            txt = r.inner_text()
            if subject in txt:
                box = r.bounding_box()
                if box:
                    page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                    page.wait_for_timeout(1500)
                    return r
        except Exception:
            pass
    return None


def search_for_one(page, subject):
    """Do a targeted search for a specific email subject. Returns found row or None."""
    # Extract a short title from [Blog-Topic] Title
    title = subject.split("]", 1)[-1].strip() if "]" in subject else subject
    safe = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014]', ' ', title)
    safe = re.sub(r'\s+', ' ', safe).strip()
    words = safe.split()
    short = " ".join(words[:6]) if len(words) > 6 else safe

    try:
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)
    except Exception:
        navigate_to_inbox(page)
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)

    sb.click()
    page.wait_for_timeout(400)
    page.keyboard.press("Control+a")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(200)

    query = f"from:{SENDER} subject:({SUBJECT_PREFIX} {short}) -\"By agent - Blog\""
    for term in EXCLUDE_TERMS:
        query += f' -{term}'
    sb.fill(query)
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(5000)

    # Find the row that has [Blog- and is in Inbox (not in target folder)
    rows = page.query_selector_all("div[role='option']")
    for r in rows:
        try:
            txt = r.inner_text()
            if "[Blog-" not in txt:
                continue
            lines = [l.strip() for l in txt.split("\n") if l.strip()]
            folder_info = lines[-1] if lines else ""
            # Skip if already in target folder
            if TARGET_FOLDER.lower() in folder_info.lower():
                continue
            box = r.bounding_box()
            if box:
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                page.wait_for_timeout(1500)
                return r
        except Exception:
            pass
    return None


def has_category_in_row(page, row):
    """Check if the email in the reading pane (after clicking row) has the category."""
    try:
        box = row.bounding_box()
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            page.wait_for_timeout(1500)
    except Exception:
        return False
    try:
        rp_text = page.locator('[role="main"]').first.inner_text()
        return CATEGORY.lower() in rp_text.lower()
    except Exception:
        return False


def get_row_subject(row):
    """Get the subject text from a row."""
    try:
        return row.inner_text().strip()[:200]
    except Exception:
        return ""


def get_row_folder_info(row):
    """Get folder info from row text (last line usually has folder name)."""
    try:
        txt = row.inner_text().strip()
        lines = [l.strip() for l in txt.split("\n") if l.strip()]
        if lines:
            last_line = lines[-1]
            return last_line
    except Exception:
        pass
    return ""


def do_categorize(page, row):
    """Categorize a single email row via right-click context menu."""
    try:
        box = row.bounding_box()
        if not box:
            return False
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
        page.wait_for_timeout(400)
        page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2, button="right")
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
                    return True  # Already has category
                sub.evaluate("el => el.click()")
                page.wait_for_timeout(800)
                return True
        except Exception:
            pass

    page.keyboard.press("Escape")
    return False


def go_home_tab(page):
    """Click the Home tab."""
    for el in page.query_selector_all("[role='tab']"):
        try:
            if "home" in el.inner_text().strip().lower():
                el.click()
                page.wait_for_timeout(1000)
                return True
        except Exception:
            pass
    return False


def do_move(page, row):
    """Move a single email row to TARGET_FOLDER via Sposta button."""
    try:
        box = row.bounding_box()
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            page.wait_for_timeout(400)
    except Exception:
        return False

    go_home_tab(page)
    page.wait_for_timeout(500)

    # Find Sposta button in toolbar (y < 200)
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

    search_input = None
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


def extract_subject(row):
    """Extract the [Blog-...] subject line from row text."""
    try:
        txt = row.inner_text()
        for line in txt.split("\n"):
            line = line.strip()
            if "[Blog-" in line and len(line) > 10:
                return line
    except Exception:
        pass
    return ""


def main():
    print(f"=== Inbox Sweep: {DATE_FROM} to {DATE_TO} ===")
    p, browser, page = connect()

    total_categorized = 0
    total_moved = 0
    total_skipped = 0
    processed_subjects = set()
    sweep_round = 0
    max_rounds = 10  # Safety limit

    try:
        while sweep_round < max_rounds:
            sweep_round += 1
            print(f"\n--- Sweep round {sweep_round} ---")

            # Navigate to Inbox and do a full scroll-search to find all inbox emails
            navigate_to_inbox(page)
            inbox_subjects = do_search(page)

            if not inbox_subjects:
                print("  No unprocessed [Blog-] emails found in Inbox. Done!")
                break

            # Process each inbox email by doing a targeted search
            round_processed = 0
            for subject in inbox_subjects:
                if subject in processed_subjects:
                    print(f"  SKIP (already attempted): {subject[:60]}")
                    total_skipped += 1
                    continue

                print(f"\n  PROCESSING: {subject[:70]}")

                # Do a targeted search for this specific email
                row = search_for_one(page, subject)
                if not row:
                    print(f"    Not found in targeted search, skipping")
                    processed_subjects.add(subject)
                    continue

                # Check if already has category
                already_cat = False
                try:
                    rp_text = page.locator('[role="main"]').first.inner_text()
                    already_cat = CATEGORY.lower() in rp_text.lower()
                except Exception:
                    pass

                # Re-find the row (DOM may have shifted)
                rows_now = page.query_selector_all("div[role='option']")
                current_row = None
                for r in rows_now:
                    try:
                        txt = r.inner_text()
                        if "[Blog-" in txt and TARGET_FOLDER.lower() not in txt.lower().split("\n")[-1]:
                            current_row = r
                            break
                    except Exception:
                        pass

                if not current_row:
                    print(f"    Row disappeared, skipping")
                    processed_subjects.add(subject)
                    continue

                # Categorize if needed
                if not already_cat:
                    cat_ok = do_categorize(page, current_row)
                    if cat_ok:
                        total_categorized += 1
                        print(f"    Categorized: OK")
                    else:
                        print(f"    Categorized: FAILED")
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                else:
                    print(f"    Already categorized")

                # Re-find row for move
                rows_now = page.query_selector_all("div[role='option']")
                current_row = None
                for r in rows_now:
                    try:
                        txt = r.inner_text()
                        if "[Blog-" in txt and TARGET_FOLDER.lower() not in txt.lower().split("\n")[-1]:
                            current_row = r
                            break
                    except Exception:
                        pass

                if not current_row:
                    print(f"    Row disappeared after categorize, counting as moved")
                    total_moved += 1
                    processed_subjects.add(subject)
                    round_processed += 1
                    continue

                # Move
                mov_ok = do_move(page, current_row)
                if mov_ok:
                    total_moved += 1
                    print(f"    Moved: OK")
                else:
                    print(f"    Moved: FAILED")

                processed_subjects.add(subject)
                round_processed += 1
                page.wait_for_timeout(1000)

            if round_processed == 0:
                print("  No emails processed this round. Breaking to avoid infinite loop.")
                break

            print(f"\n  Round {sweep_round} done: processed {round_processed}")

    finally:
        p.stop()

    print(f"\n{'='*60}")
    print(f"=== SWEEP COMPLETE ===")
    print(f"  Rounds: {sweep_round}")
    print(f"  Categorized: {total_categorized}")
    print(f"  Moved: {total_moved}")
    print(f"  Skipped (already attempted): {total_skipped}")
    print(f"  Unique subjects processed: {len(processed_subjects)}")


if __name__ == "__main__":
    main()
