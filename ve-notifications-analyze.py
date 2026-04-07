"""
Analyze a Viva Engage notification email content in Outlook reading pane.
Opens a specific email by index in the search results and extracts detailed
content from the email body: community, author, post date, post text, etc.

Usage: python ve-notifications-analyze.py <index>
  - <index>: 0-based index of current search result to analyze

The script expects the Outlook search results from ve-notifications-retrieve.py
to still be visible. It clicks on the specified row and reads the reading pane.

Output: JSON to stdout with extracted fields.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
CACHE_FILE = os.path.join(BASE, "ve_notifications_cache.json")


def connect_to_outlook():
    """Connect to Edge via CDP and find an existing Outlook tab."""
    ensure_edge_cdp()
    p = sync_playwright().start()
    browser = p.chromium.connect_over_cdp(CDP_URL)
    ctx = browser.contexts[0]

    page = None
    for pg in ctx.pages:
        url = pg.url
        if "outlook.office.com/mail" in url or "outlook.cloud.microsoft/mail" in url:
            page = pg
            break

    if not page:
        page = ctx.new_page()
        page.goto("https://outlook.cloud.microsoft/mail/", wait_until="commit", timeout=30000)
        page.wait_for_timeout(4000)

    return p, browser, page


def extract_email_details(page):
    """Extract detailed VE notification content from the reading pane."""
    try:
        rp = page.locator('[role="main"]').first
        rp_text = rp.inner_text()
    except Exception:
        return None

    result = {
        "subject": "",
        "received_date": "",
        "post_type": "",
        "post_title": "",
        "community_name": "",
        "community_url": "",
        "author": "",
        "post_date": "",
        "post_body_text": "",
        "thread_url": "",
    }

    # Subject line
    for line in rp_text.split("\n"):
        line = line.strip()
        if len(line) > 10:
            result["subject"] = line
            break

    # Date
    m = re.search(r'(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2})', rp_text)
    if m:
        parts = m.group(1).split("/")
        result["received_date"] = f"{parts[2]}-{parts[1]}-{parts[0]}T{m.group(2)}:00"

    # Post type and title from subject
    subj = result["subject"]
    type_match = re.match(r'^(Question|Announcement|Praise|Discussion|Poll|Article)\s*:\s*(.+)', subj, re.IGNORECASE)
    if type_match:
        result["post_type"] = type_match.group(1).strip()
        result["post_title"] = type_match.group(2).strip()
    else:
        result["post_title"] = subj

    # Community name and URL
    try:
        links = page.locator('[role="main"] a[href]').all()
        for link in links:
            try:
                parent_text = link.evaluate("el => el.parentElement ? el.parentElement.innerText : ''")
                if "pubblicato in" in parent_text.lower() or "posted in" in parent_text.lower():
                    result["community_name"] = link.inner_text().strip()
                    result["community_url"] = link.get_attribute("href") or ""
                    break
            except Exception:
                continue
    except Exception:
        pass

    # Author - look for the name after "Pubblicato in <community>" or similar pattern
    lines = rp_text.split("\n")
    for i, line in enumerate(lines):
        if "pubblicato in" in line.lower() or "posted in" in line.lower():
            # The author name is typically on the next non-empty line
            for j in range(i + 1, min(i + 5, len(lines))):
                candidate = lines[j].strip()
                if candidate and len(candidate) > 2 and not candidate.startswith("http"):
                    result["author"] = candidate
                    break
            break

    # Post date - look for date patterns in the body area
    date_patterns = [
        r'(\d{1,2}\s+(?:gen|feb|mar|apr|mag|giu|lug|ago|set|ott|nov|dic)\w*\.?\s+\d{4})',
        r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\.?\s+\d{4})',
        r'(\d{4}-\d{2}-\d{2})',
    ]
    for pat in date_patterns:
        m = re.search(pat, rp_text, re.IGNORECASE)
        if m:
            result["post_date"] = m.group(1)
            break

    # Thread URL
    try:
        for link in page.locator('[role="main"] a[href]').all():
            href = link.get_attribute("href") or ""
            if "engage.cloud.microsoft" in href and ("thread" in href or "groups" in href):
                result["thread_url"] = href
                break
    except Exception:
        pass

    # Post body text - the main content area of the email
    # Capture from the email body text, trimming header/footer noise
    result["post_body_text"] = rp_text[:3000]

    return result


def main():
    if len(sys.argv) < 2:
        print("Usage: python ve-notifications-analyze.py <index>")
        sys.exit(1)

    idx = int(sys.argv[1])

    p, browser, page = connect_to_outlook()

    try:
        # Click the item at the given index in the listbox
        items = page.locator('[role="listbox"] [role="option"]').all()
        if idx >= len(items):
            print(json.dumps({"error": f"Index {idx} out of range (have {len(items)} items)"}))
            return

        items[idx].click()
        page.wait_for_timeout(2000)

        details = extract_email_details(page)
        if details:
            print(json.dumps(details, indent=2, ensure_ascii=False))
        else:
            print(json.dumps({"error": "Could not extract email details"}))

    finally:
        p.stop()


if __name__ == "__main__":
    main()
