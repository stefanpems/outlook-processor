"""
Retrieve Viva Engage notification emails from Outlook Web for a date range.
Connects to Edge via CDP, searches Outlook, scrolls virtualized list,
extracts email data from reading pane.

Usage: python ve-notifications-retrieve.py YYYY-MM-DD YYYY-MM-DD [--include-processed]
Output: Saves results to ve_notifications_cache.json and prints summary JSON.
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
EXCLUDE_TERMS = VE_CFG.get("exclude_terms", [])
CACHE_FILE = os.path.join(BASE, "ve_notifications_cache.json")


def connect_to_outlook():
    """Connect to Edge via CDP and find/create an Outlook tab."""
    ensure_edge_cdp()
    p = sync_playwright().start()
    browser = p.chromium.connect_over_cdp(CDP_URL)
    ctx = browser.contexts[0]

    page = None
    fallback = None
    for pg in ctx.pages:
        url = pg.url
        if "outlook.office.com/mail" in url or "outlook.cloud.microsoft/mail" in url:
            if "/mail/id/" not in url:
                page = pg
                break
            elif fallback is None:
                fallback = pg

    if not page:
        if fallback:
            page = fallback
            page.evaluate("window.location.href = 'https://outlook.cloud.microsoft/mail/'")
            page.wait_for_timeout(4000)
        else:
            page = ctx.new_page()
            page.goto("https://outlook.cloud.microsoft/mail/", wait_until="commit", timeout=30000)
            page.wait_for_timeout(4000)

    print(f"Using page: {page.url[:80]}")
    return p, browser, page


def search_emails(page, date_from, date_to, include_processed=False):
    """Execute Outlook search for VE notification emails in date range. Returns visible count."""
    page.evaluate("window.location.href = 'https://outlook.cloud.microsoft/mail/'")
    page.wait_for_timeout(4000)

    search_input = page.locator("#topSearchInput")
    search_input.click()
    page.wait_for_timeout(500)
    search_input.fill("")
    page.wait_for_timeout(300)

    query = f"from:{SENDER} received:{date_from}..{date_to}"
    for term in EXCLUDE_TERMS:
        query += f' -{term}'
    if not include_processed:
        query += f' -"{CATEGORY}"'
    print(f"  Query: {query}")
    search_input.fill(query)
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(5000)

    for attempt in range(4):
        items = page.locator('[role="listbox"] [role="option"]').all()
        if items:
            return len(items)
        print(f"  No results yet (attempt {attempt+1}/4), waiting...")
        page.wait_for_timeout(3000)

    print("  No results found after retries")
    return 0


def count_all_results(page):
    """Scroll the virtualized listbox and count all unique items."""
    seen = set()

    def read_visible():
        items = page.locator('[role="listbox"] [role="option"]').all()
        new_count = 0
        for item in items:
            try:
                preview = item.inner_text().strip()[:200]
                if preview and preview not in seen:
                    seen.add(preview)
                    new_count += 1
            except Exception:
                pass
        return new_count

    read_visible()

    bbox = page.locator('[role="listbox"]').first.bounding_box()
    if not bbox:
        return len(seen)

    cx = bbox['x'] + bbox['width'] / 2
    cy = bbox['y'] + bbox['height'] / 2

    stable = 0
    while stable < 3:
        page.mouse.move(cx, cy)
        page.mouse.wheel(0, 400)
        page.wait_for_timeout(1500)
        n = read_visible()
        stable = 0 if n > 0 else stable + 1

    return len(seen)


def get_reading_pane_fingerprint(page):
    """Get a short fingerprint of the current reading pane content."""
    try:
        rp_text = page.locator('[role="main"]').first.inner_text()
        return rp_text.strip()[:300]
    except Exception:
        return ""


def wait_for_pane_change(page, old_fingerprint, max_wait_ms=5000):
    """Wait until the reading pane content changes from old_fingerprint."""
    waited = 0
    step = 500
    while waited < max_wait_ms:
        page.wait_for_timeout(step)
        waited += step
        new_fp = get_reading_pane_fingerprint(page)
        if new_fp and new_fp != old_fingerprint:
            return True
    return False


def read_reading_pane(page):
    """Extract VE notification email data from the current reading pane."""
    try:
        rp = page.locator('[role="main"]').first
        rp_text = rp.inner_text()
    except Exception:
        return None

    if not rp_text or len(rp_text) < 20:
        return None

    # Subject: first substantial line (usually the email subject)
    subject = ""
    for line in rp_text.split("\n"):
        line = line.strip()
        if len(line) > 10 and ":" in line:
            subject = line
            break
    if not subject:
        # Fallback: just get the first line with enough characters
        for line in rp_text.split("\n"):
            line = line.strip()
            if len(line) > 15:
                subject = line
                break

    if not subject:
        return None

    # Date
    received_date = ""
    m = re.search(r'(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})', rp_text)
    if m:
        parts = m.group(1).split("/")
        received_date = f"{parts[2]}-{parts[1]}-{parts[0]}T{m.group(2)}:00"

    # Parse post type and title from subject.
    # VE notification subjects look like "Question: How to configure X" or "Announcement: New feature Y"
    post_type = ""
    post_title = subject
    type_match = re.match(r'^(Question|Announcement|Praise|Discussion|Poll|Article)\s*:\s*(.+)', subject, re.IGNORECASE)
    if type_match:
        post_type = type_match.group(1).strip()
        post_title = type_match.group(2).strip()

    # Extract community link (link following "Pubblicato in " or "Posted in ")
    community_name = ""
    community_url = ""
    try:
        links = page.locator('[role="main"] a[href]').all()
        for link in links:
            try:
                parent_text = link.evaluate("el => el.parentElement ? el.parentElement.innerText : ''")
                if "pubblicato in" in parent_text.lower() or "posted in" in parent_text.lower():
                    community_name = link.inner_text().strip()
                    community_url = link.get_attribute("href") or ""
                    break
            except Exception:
                continue
    except Exception:
        pass

    # Extract author
    author = ""
    author_match = re.search(r'(?:Pubblicato in .+?\n)(.+?)(?:\n|ha pubblicato|posted)', rp_text, re.IGNORECASE)
    if author_match:
        author = author_match.group(1).strip()

    # Extract Viva Engage thread URL from links
    thread_url = ""
    try:
        for link in page.locator('[role="main"] a[href]').all():
            href = link.get_attribute("href") or ""
            if "engage.cloud.microsoft" in href and "thread" in href:
                thread_url = href
                break
    except Exception:
        pass

    # Full body text of the email (for later analysis)
    body_text = rp_text

    return {
        "subject": subject,
        "received_date": received_date,
        "post_type": post_type,
        "post_title": post_title,
        "community_name": community_name,
        "community_url": community_url,
        "author": author,
        "thread_url": thread_url,
        "body_text": body_text[:3000],
        "processed": False,
    }


def has_processed_category(page):
    """Check if the currently displayed email has the processed category label."""
    try:
        rp_text = page.locator('[role="main"]').first.inner_text()
        return CATEGORY.lower() in rp_text.lower()
    except Exception:
        return False


def extract_all_via_keyboard(page, total_count, include_processed=False):
    """Click first result, then navigate with Down arrow. Extract from reading pane."""
    emails = []
    null_count = 0
    stuck_count = 0
    last_subject = ""

    # Scroll listbox back to top
    bbox = page.locator('[role="listbox"]').first.bounding_box()
    if bbox:
        cx = bbox['x'] + bbox['width'] / 2
        cy = bbox['y'] + bbox['height'] / 2
        page.mouse.move(cx, cy)
        for _ in range(20):
            page.mouse.wheel(0, -500)
        page.wait_for_timeout(1000)

    # Click the first item
    first = page.locator('[role="listbox"] [role="option"]').first
    first.click()
    page.wait_for_timeout(2000)

    for i in range(total_count):
        em = read_reading_pane(page)

        if em is None:
            page.wait_for_timeout(2000)
            em = read_reading_pane(page)

        if em:
            current_subject = em.get('subject', '')
            if current_subject == last_subject:
                stuck_count += 1
            else:
                stuck_count = 0
            last_subject = current_subject

            if stuck_count >= 5:
                print(f"    [{i+1}/{total_count}] STUCK: navigation appears frozen at: {current_subject[:60]}")
                print(f"    Attempting to unstick by scrolling listbox...")
                lb_box = page.locator('[role="listbox"]').first.bounding_box()
                if lb_box:
                    lx = lb_box['x'] + lb_box['width'] / 2
                    ly = lb_box['y'] + lb_box['height'] / 2
                    page.mouse.move(lx, ly)
                    page.mouse.wheel(0, 300)
                    page.wait_for_timeout(1500)
                    options = page.locator('[role="listbox"] [role="option"]').all()
                    if len(options) > 1:
                        options[-1].click()
                        page.wait_for_timeout(2000)
                        stuck_count = 0
                        last_subject = ""
                        em = read_reading_pane(page)
                        if em is None:
                            print(f"    [{i+1}/{total_count}] STILL NULL after unstick attempt")
                            null_count += 1
                            if i < total_count - 1:
                                page.keyboard.press("ArrowDown")
                                page.wait_for_timeout(1500)
                            continue

            if include_processed or not has_processed_category(page):
                emails.append(em)
                print(f"    [{i+1}/{total_count}] {em.get('post_type', '?')}: {em.get('post_title', '?')[:60]}")
            else:
                print(f"    [{i+1}/{total_count}] SKIP (already categorized): {em.get('subject','')[:60]}")
        else:
            null_count += 1
            print(f"    [{i+1}/{total_count}] NULL: reading pane returned no data")

        if i < total_count - 1:
            fp_before = get_reading_pane_fingerprint(page)
            page.keyboard.press("ArrowDown")
            changed = wait_for_pane_change(page, fp_before, max_wait_ms=5000)
            if not changed:
                page.wait_for_timeout(1000)
            page.wait_for_timeout(500)

    if null_count > 0:
        print(f"  WARNING: {null_count} items returned NULL from reading pane (silently missed)")

    return emails


def main():
    args = [a for a in sys.argv[1:] if a != "--include-processed"]
    include_processed = "--include-processed" in sys.argv

    if len(args) < 2:
        print("Usage: python ve-notifications-retrieve.py YYYY-MM-DD YYYY-MM-DD [--include-processed]")
        sys.exit(1)

    date_from = args[0]
    date_to = args[1]
    mode_label = " (including already-processed)" if include_processed else ""
    print(f"=== Retrieving VE notification emails from {date_from} to {date_to}{mode_label} ===")

    p, browser, page = connect_to_outlook()

    try:
        initial = search_emails(page, date_from, date_to, include_processed=include_processed)
        if initial == 0:
            print("No emails found.")
            # Initialize empty cache
            cache = {"emails": [], "date_from": date_from, "date_to": date_to}
            json.dump(cache, open(CACHE_FILE, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
            result = {"total": 0, "emails": []}
            print(json.dumps(result))
            return

        total_count = count_all_results(page)
        print(f"  Total unique items: {total_count}")

        emails = extract_all_via_keyboard(page, total_count, include_processed=include_processed)

        # Deduplicate by email identity (subject + date)
        seen_keys = set()
        unique_emails = []
        for em in emails:
            key = (em.get("subject", ""), em.get("received_date", ""))
            if key not in seen_keys:
                seen_keys.add(key)
                unique_emails.append(em)
        if len(unique_emails) < len(emails):
            print(f"  Removed {len(emails) - len(unique_emails)} duplicates")
        emails = unique_emails

        emails.sort(key=lambda e: e.get("received_date", ""), reverse=True)
        print(f"  Extracted {len(emails)} emails")

        # Save to cache
        cache = {"emails": emails, "date_from": date_from, "date_to": date_to}
        json.dump(cache, open(CACHE_FILE, "w", encoding="utf-8"), indent=2, ensure_ascii=False)

        result = {"total": len(emails), "date_from": date_from, "date_to": date_to}
        print(json.dumps(result, indent=2))

    finally:
        p.stop()


if __name__ == "__main__":
    main()
