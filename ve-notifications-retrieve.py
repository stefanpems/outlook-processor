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
    while stable < 5:
        page.mouse.move(cx, cy)
        page.mouse.wheel(0, 600)
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


def wait_for_pane_change(page, old_fingerprint, max_wait_ms=15000):
    """Wait until the reading pane content changes from old_fingerprint.
    Uses Python-side polling to avoid cross-boundary string comparison issues."""
    elapsed = 0
    interval = 500  # ms
    while elapsed < max_wait_ms:
        page.wait_for_timeout(interval)
        elapsed += interval
        current = get_reading_pane_fingerprint(page)
        if current and current != old_fingerprint:
            return True
    return False


def expand_and_wait_for_content(page, max_wait_ms=8000):
    """Wait for VE notification email content to fully load in the reading pane."""

    # Wait until "Pubblicato in" / "Posted in" appears — single JS check in browser, no round-trips
    try:
        page.wait_for_function(
            """() => {
                const el = document.querySelector('[role="main"]');
                if (!el) return false;
                const t = el.innerText.toLowerCase();
                return t.includes('pubblicato in') || t.includes('posted in');
            }""",
            timeout=max_wait_ms
        )
    except Exception:
        pass


def read_reading_pane(page):
    """Extract VE notification email data from the current reading pane."""
    # Wait for full content to load
    expand_and_wait_for_content(page)

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

    # Extract all link data in a single JS call (community + thread URL)
    community_name = ""
    community_url = ""
    thread_url = ""
    try:
        link_data = page.evaluate("""() => {
            const main = document.querySelector('[role="main"]');
            if (!main) return {community: null, thread: null};
            const links = Array.from(main.querySelectorAll('a[href]'));
            let community = null, thread = null, groupFallback = null;
            for (const a of links) {
                const href = a.href || '';
                const text = a.innerText || '';
                if (!community && a.parentElement) {
                    const pt = (a.parentElement.innerText || '').toLowerCase();
                    if (pt.includes('pubblicato in') || pt.includes('posted in')) {
                        community = {name: text.trim(), url: href};
                    }
                }
                if (!thread && href.includes('engage.cloud.microsoft') && href.includes('/threads/')) {
                    thread = href;
                }
                if (!groupFallback && href.includes('engage.cloud.microsoft') && href.includes('/groups/')) {
                    groupFallback = href;
                }
            }
            return {community, thread: thread || groupFallback};
        }""")
        if link_data.get("community"):
            community_name = link_data["community"]["name"]
            community_url = link_data["community"]["url"]
        thread_url = link_data.get("thread") or ""
    except Exception:
        pass

    # Extract author — name appears on a line before "Pubblicato in" or after initials block
    author = ""
    lines = rp_text.split("\n")
    for i, line in enumerate(lines):
        low = line.strip().lower()
        if "pubblicato in" in low or "annuncio pubblicato in" in low or "posted in" in low:
            for j in range(max(0, i - 5), i):
                candidate = lines[j].strip()
                if (candidate and len(candidate) > 3 and
                    not candidate.startswith("http") and
                    "riepilog" not in candidate.lower() and
                    "mostra" not in candidate.lower() and
                    "cambia" not in candidate.lower() and
                    ":" not in candidate[:3]):
                    author = candidate.split(",")[0].strip()
                    break
            break

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


def has_processed_category(page, rp_text=None):
    """Check if the currently displayed email has the processed category label."""
    try:
        if rp_text is None:
            rp_text = page.locator('[role="main"]').first.inner_text()
        return CATEGORY.lower() in rp_text.lower()
    except Exception:
        return False


def extract_all_via_keyboard(page, total_count, include_processed=False):
    """Click first result, then navigate with Down arrow. Extract from reading pane.
    Uses total_count as estimate for logging; continues until end of list is detected."""
    emails = []
    null_count = 0
    MAX_ADVANCE_RETRIES = 4  # stop after N consecutive ArrowDown with no pane change

    # Scroll listbox back to top
    bbox = page.locator('[role="listbox"]').first.bounding_box()
    if bbox:
        cx = bbox['x'] + bbox['width'] / 2
        cy = bbox['y'] + bbox['height'] / 2
        page.mouse.move(cx, cy)
        for _ in range(20):
            page.mouse.wheel(0, -500)
        page.wait_for_timeout(1000)

    # Skip "Risultati principali" section — click the first option AFTER the second header
    # Outlook search results have: [Header: "Risultati principali"] [dup options...] [Header: "Tutti i risultati"/date] [real options...]
    # The second header's name varies (e.g., "Tutti i risultati", date-based). We find the
    # first option that follows ANY header after the first one.
    first_real_option_index = page.evaluate("""() => {
        const lb = document.querySelector('[role="listbox"]');
        if (!lb) return 0;
        const options = lb.querySelectorAll('[role="option"]');
        const headers = lb.querySelectorAll('[role="button"]');
        if (headers.length < 2) return 0; // no "Risultati principali" section
        // Get the bounding rect of the second header
        const secondHeader = headers[1];
        const headerBottom = secondHeader.getBoundingClientRect().bottom;
        // Find the first option below that header
        for (let i = 0; i < options.length; i++) {
            if (options[i].getBoundingClientRect().top >= headerBottom) return i;
        }
        return 0;
    }""")
    
    if first_real_option_index > 0:
        print(f"  Skipping 'Risultati principali' section ({first_real_option_index} items), starting at option index {first_real_option_index}")
    
    first = page.locator('[role="listbox"] [role="option"]').nth(first_real_option_index)
    first.click()
    page.wait_for_timeout(2000)

    i = 0
    while True:
        try:
            em = read_reading_pane(page)

            if em is None:
                # Retry: go back to previous email, then forward again to force reload
                page.keyboard.press("ArrowUp")
                page.wait_for_timeout(1000)
                page.keyboard.press("ArrowDown")
                page.wait_for_timeout(10000)
                em = read_reading_pane(page)

            if em:
                if include_processed or not has_processed_category(page, em.get('body_text', '')):
                    emails.append(em)
                    print(f"    [{i+1}/{total_count}+] {em.get('post_type', '?')}: {em.get('post_title', '?')[:60]}")
                else:
                    print(f"    [{i+1}/{total_count}+] SKIP (already categorized): {em.get('subject','')[:60]}")
            else:
                null_count += 1
                print(f"    [{i+1}/{total_count}+] NULL: reading pane returned no data")

            # Re-focus the message list (Outlook may have moved focus to reading pane)
            try:
                page.evaluate("""() => {
                    const opts = document.querySelectorAll('[role="listbox"] [role="option"]');
                    for (const opt of opts) {
                        if (opt.getAttribute('tabindex') === '0' || opt.getAttribute('aria-selected') === 'true') {
                            opt.focus();
                            return;
                        }
                    }
                    const lb = document.querySelector('[role="listbox"]');
                    if (lb) lb.focus();
                }""")
                page.wait_for_timeout(300)
            except Exception:
                pass

            # Try to advance to next item
            advanced = False
            for attempt in range(MAX_ADVANCE_RETRIES):
                fp_before = get_reading_pane_fingerprint(page)
                page.keyboard.press("ArrowDown")
                changed = wait_for_pane_change(page, fp_before, max_wait_ms=15000)
                if changed:
                    advanced = True
                    page.wait_for_timeout(500)
                    break
                else:
                    page.wait_for_timeout(1000)
                    # On second retry, scroll listbox to force virtualization to load more
                    if attempt == 1:
                        lb_box = page.locator('[role="listbox"]').first.bounding_box()
                        if lb_box:
                            lx = lb_box['x'] + lb_box['width'] / 2
                            ly = lb_box['y'] + lb_box['height'] / 2
                            page.mouse.move(lx, ly)
                            page.mouse.wheel(0, 400)
                            page.wait_for_timeout(1500)

            if not advanced:
                print(f"  End of list reached after {i+1} items (pane unchanged for {MAX_ADVANCE_RETRIES} retries)")
                break

        except Exception as e:
            print(f"  ERROR at item {i+1}: {type(e).__name__}: {e}")
            print(f"  Returning {len(emails)} emails collected so far")
            break

        i += 1

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
