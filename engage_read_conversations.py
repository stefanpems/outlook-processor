#!/usr/bin/env python3
"""
Read recent conversations from a Viva Engage community via Edge CDP.

Navigates to the named community, ensures "All conversations" + "Recent activity"
sort, scrolls the virtualized feed, expands replies/truncated text, extracts each
conversation's full text with dates, and stops when all dates in a thread fall
before the cutoff (reference_date - days).

Usage:
    python engage_read_conversations.py "<community_name>" <days>

Arguments:
    community_name  Exact community name as shown in the Viva Engage sidebar
    days            Number of days of recent activity to include

Output:
    JSON to stdout with all conversation data for LLM summarization.
    Progress messages go to stderr.
"""

import sys, json, time, re, os
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config():
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Navigation
# ---------------------------------------------------------------------------

def navigate_to_community(page, community_name):
    """Navigate to Viva Engage and open the named community."""
    page.goto("https://engage.cloud.microsoft/", timeout=60000)
    time.sleep(5)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    # Dismiss any overlay / popup
    try:
        page.keyboard.press("Escape")
        time.sleep(1)
    except Exception:
        pass

    # Find and click the community link in the sidebar
    link = page.get_by_text(community_name, exact=True)
    if link.count() == 0:
        raise RuntimeError(f"Community '{community_name}' not found in sidebar")
    link.first.click()
    time.sleep(5)
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    title = page.title()
    print(f"  Page title: {title}", file=sys.stderr)


def verify_filters(page):
    """Ensure 'All conversations' and 'Recent activity' sort are active."""
    dropdowns = page.query_selector_all("button[aria-haspopup]")
    active = set()
    for d in dropdowns:
        try:
            txt = d.inner_text().strip()
            if txt in ("All conversations", "Recent activity"):
                active.add(txt)
        except Exception:
            pass

    # If "All conversations" is not active, try to select it
    if "All conversations" not in active:
        for d in dropdowns:
            try:
                txt = d.inner_text().strip()
                if "conversation" in txt.lower():
                    d.click()
                    time.sleep(1)
                    opt = page.get_by_text("All conversations", exact=True)
                    if opt.count() > 0:
                        opt.first.click()
                        time.sleep(2)
                    break
            except Exception:
                pass

    # If "Recent activity" is not active, try to select it
    if "Recent activity" not in active:
        for d in dropdowns:
            try:
                txt = d.inner_text().strip()
                if "recent" in txt.lower():
                    d.click()
                    time.sleep(1)
                    opt = page.get_by_text("Recent activity", exact=True)
                    if opt.count() > 0:
                        opt.first.click()
                        time.sleep(2)
                    break
            except Exception:
                pass

    print("  Filters: All conversations / Recent activity", file=sys.stderr)


# ---------------------------------------------------------------------------
# Date parsing  (Viva Engage relative dates → absolute datetime)
# ---------------------------------------------------------------------------

_WEEKDAYS = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}
_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _parse_one(text, ref):
    """Convert a single Viva Engage date token to datetime (or None)."""
    text = text.strip()

    # "Xh" / "Xm" / "Xd"
    m = re.match(r"(\d+)([hmd])", text, re.IGNORECASE)
    if m:
        n, u = int(m.group(1)), m.group(2).lower()
        delta = {"h": timedelta(hours=n), "m": timedelta(minutes=n), "d": timedelta(days=n)}
        return ref - delta.get(u, timedelta())

    # "Yesterday at H:MM AM/PM"
    m = re.match(r"Yesterday\s+at\s+(\d{1,2}):(\d{2})\s*(AM|PM)", text, re.IGNORECASE)
    if m:
        h = int(m.group(1)) % 12 + (12 if m.group(3).upper() == "PM" else 0)
        return (ref - timedelta(days=1)).replace(hour=h, minute=int(m.group(2)), second=0, microsecond=0)

    # "Day at H:MM AM/PM"  (e.g. "Fri at 3:19 PM")
    m = re.match(r"(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+at\s+(\d{1,2}):(\d{2})\s*(AM|PM)", text, re.IGNORECASE)
    if m:
        target = _WEEKDAYS[m.group(1).lower()]
        days_back = (ref.weekday() - target) % 7
        h = int(m.group(2)) % 12 + (12 if m.group(4).upper() == "PM" else 0)
        return (ref - timedelta(days=days_back)).replace(hour=h, minute=int(m.group(3)), second=0, microsecond=0)

    # "Mon DD, YYYY"
    m = re.match(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}),\s*(\d{4})", text, re.IGNORECASE)
    if m:
        return datetime(int(m.group(3)), _MONTHS[m.group(1).lower()], int(m.group(2)))

    # "Mon DD" (no year)
    m = re.match(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})", text, re.IGNORECASE)
    if m:
        cand = ref.replace(month=_MONTHS[m.group(1).lower()], day=int(m.group(2)),
                           hour=0, minute=0, second=0, microsecond=0)
        if cand > ref:
            cand = cand.replace(year=ref.year - 1)
        return cand

    return None


def extract_dates(text, ref):
    """Return a list[datetime] of all posting-date tokens found in *text*."""
    dates = []
    patterns = [
        r"(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)",
        r"Yesterday\s+at\s+\d{1,2}:\d{2}\s*(?:AM|PM)",
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s*\d{4}",
        r"(?<!\w)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}(?!\d|,)",
        r"(?<!\w)\d+[hmd](?!\w)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            d = _parse_one(m.group(), ref)
            if d:
                dates.append(d)
    return dates


# ---------------------------------------------------------------------------
# Feed interaction helpers
# ---------------------------------------------------------------------------

def get_thread_headings(page):
    """Return list[dict] of currently visible thread headings with thread URLs."""
    return page.evaluate(r"""
        (() => {
            const els = document.querySelectorAll('[id*="heading-thread-"]');
            return Array.from(els).map(e => {
                // Extract the base64 thread token from the id
                const token = e.id.replace('heading-thread-', '');
                // Build the thread URL using the current page path
                const groupPath = window.location.pathname.match(/\/main\/groups\/[^/]+/);
                const threadUrl = groupPath
                    ? window.location.origin + groupPath[0] + '/thread/' + token
                    : '';
                // Check if this thread block contains images
                let container = e.closest('[class*="thread"]') || e.parentElement;
                // Walk up a few levels to find the thread container
                for (let i = 0; i < 5 && container; i++) {
                    if (container.querySelectorAll('img[src*="blob"], img[src*="attachment"], img[src*="graph"], img[src*="sharepoint"], [data-testid*="image"]').length > 0) break;
                    container = container.parentElement;
                }
                const hasImages = container
                    ? container.querySelectorAll('img[src*="blob"], img[src*="attachment"], img[src*="graph"], img[src*="sharepoint"], [data-testid*="image"]').length > 0
                    : false;
                return {
                    id: e.id,
                    text: e.innerText.substring(0, 200),
                    thread_url: threadUrl,
                    has_images: hasImages
                };
            });
        })()
    """)


def setup_clipboard_interceptor(page):
    """Grant clipboard permissions for the current page context."""
    try:
        page.context.grant_permissions(
            ["clipboard-read", "clipboard-write"],
            origin="https://engage.cloud.microsoft",
        )
    except Exception as e:
        print(f"  Warning: Could not grant clipboard permissions: {e}", file=sys.stderr)


def _clear_clipboard():
    """No-op: clipboard cleared at read time."""
    pass


def _read_clipboard_from_page(page):
    """Read clipboard text via the browser's clipboard API."""
    try:
        return page.evaluate("navigator.clipboard.readText()") or ""
    except Exception:
        return ""


def get_thread_url_via_copy_link(page, heading_id):
    """Get the canonical thread URL via the '…' menu → 'Copy link' button."""
    try:
        # Scroll the heading into view first
        page.evaluate("""
            (hid) => {
                const h = document.getElementById(hid);
                if (h) h.scrollIntoView({block: 'center', behavior: 'instant'});
            }
        """, heading_id)
        time.sleep(1)

        # Find the '…' (more options) button near the heading using aria-label
        more_btn = page.evaluate(r"""
            (hid) => {
                const h = document.getElementById(hid);
                if (!h) return null;
                const hr = h.getBoundingClientRect();
                let el = h;
                for (let i = 0; i < 15 && el; i++) {
                    el = el.parentElement;
                    if (!el) break;
                    const btns = el.querySelectorAll('button[aria-label*="more"]');
                    for (const btn of btns) {
                        const r = btn.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0
                            && r.y > 0 && r.y < window.innerHeight
                            && Math.abs(r.y - hr.y) < 250) {
                            return {x: r.x + r.width/2, y: r.y + r.height/2};
                        }
                    }
                }
                return null;
            }
        """, heading_id)

        if not more_btn:
            print(f"  Warning: '...' button not found for {heading_id}", file=sys.stderr)
            return ""

        page.mouse.click(more_btn['x'], more_btn['y'])
        time.sleep(1.5)

        # Click "Copy link" in the opened menu
        copy_link = page.get_by_text("Copy link", exact=True)
        if copy_link.count() > 0:
            copy_link.first.click(timeout=5000)
            time.sleep(1.5)
        else:
            print(f"  Warning: 'Copy link' menu item not found for {heading_id}", file=sys.stderr)
            page.keyboard.press("Escape")
            time.sleep(0.3)
            return ""

        url = _read_clipboard_from_page(page)

        # Dismiss toast / popup
        page.keyboard.press("Escape")
        time.sleep(0.3)

        if url and url.startswith("http"):
            print(f"  Copied URL: {url[:80]}…", file=sys.stderr)
            return url
        else:
            print(f"  Warning: clipboard did not contain a URL (got: '{url[:50]}')", file=sys.stderr)
            return ""

    except Exception as e:
        print(f"  Warning: Copy link failed for {heading_id}: {e}", file=sys.stderr)
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        return ""


def scroll_feed_to_top(page):
    """Scroll the large overflow container (the feed) to the top."""
    page.evaluate("""
        (() => {
            for (const el of document.querySelectorAll('*')) {
                const s = window.getComputedStyle(el);
                if ((s.overflowY === 'auto' || s.overflowY === 'scroll')
                    && el.scrollHeight > el.clientHeight + 50
                    && el.scrollHeight > 5000) {
                    el.scrollTop = 0;
                    return;
                }
            }
        })()
    """)


def scroll_feed_down(page, amount=1200):
    """Scroll the feed container down by *amount* pixels."""
    page.evaluate("""
        (amount) => {
            for (const el of document.querySelectorAll('*')) {
                const s = window.getComputedStyle(el);
                if ((s.overflowY === 'auto' || s.overflowY === 'scroll')
                    && el.scrollHeight > el.clientHeight + 50
                    && el.scrollHeight > 5000) {
                    el.scrollTop += amount;
                    return;
                }
            }
        }
    """, amount)


def expand_visible_content(page):
    """Click 'see more' (text expansion) and reply-count buttons that are visible."""
    buttons = page.query_selector_all("button")
    for btn in buttons:
        try:
            if not btn.is_visible():
                continue
            txt = btn.inner_text().strip()
            lo = txt.lower()

            # Expand truncated post body
            if lo == "see more":
                btn.evaluate("e => e.click()")
                time.sleep(0.5)
                continue

            # Expand collapsed reply chains  (e.g. "3 replies", "Show 2 more answers")
            if len(txt) < 50 \
               and ("repl" in lo or "answer" in lo) \
               and any(c.isdigit() for c in txt) \
               and "hide" not in lo and "collapse" not in lo and "less" not in lo:
                btn.evaluate("e => e.click()")
                time.sleep(1)
        except Exception:
            pass


def get_main_text(page):
    """Return the inner text of <main> (or [role=main])."""
    main = page.query_selector("main, [role=main]")
    return main.inner_text() if main else ""


# ---------------------------------------------------------------------------
# Thread text extraction
# ---------------------------------------------------------------------------

def extract_thread_text(full_text, heading_text, all_headings, current_id):
    """
    Isolate the portion of *full_text* that belongs to the thread identified by
    *current_id* / *heading_text*, using the other headings as boundaries.
    """
    # Find start of this thread
    for length in (60, 30, 15):
        start = full_text.find(heading_text[:length])
        if start != -1:
            break
    else:
        return full_text  # fallback: return everything

    # Find end: the earliest position of any OTHER heading after start
    end = len(full_text)
    for h in all_headings:
        if h["id"] == current_id:
            continue
        for length in (60, 30):
            pos = full_text.find(h["text"][:length], start + 10)
            if pos != -1 and pos < end:
                end = pos
                break

    # Optionally trim at the "Write a comment" block that closes a top-level thread
    marker = "Write a comment\nDrag files to attach"
    mpos = full_text.find(marker, start + 10)
    if mpos != -1 and mpos < end:
        end = mpos + len(marker)

    return full_text[start:end].strip()


def detect_type(text):
    """Classify thread as question / announcement / discussion."""
    head = text[:500].upper()
    if "QUESTION" in head:
        return "question"
    if "ANNOUNCEMENT" in head:
        return "announcement"
    return "discussion"


# ---------------------------------------------------------------------------
# Main reading loop
# ---------------------------------------------------------------------------

def read_conversations(page, days, ref):
    """
    Scroll the community feed and collect conversations whose activity
    falls within [ref - days .. ref].  Stop at the first thread where
    ALL parsed dates are older than the cutoff.
    """
    cutoff = (ref - timedelta(days=days)).replace(hour=0, minute=0, second=0, microsecond=0)
    conversations = []
    seen_ids = set()
    no_new_rounds = 0
    MAX_IDLE = 10          # stop after this many scroll rounds with no new threads

    scroll_feed_to_top(page)
    time.sleep(3)
    setup_clipboard_interceptor(page)

    while no_new_rounds < MAX_IDLE:
        headings = get_thread_headings(page)
        found_new = False

        for h in headings:
            if h["id"] in seen_ids:
                continue

            found_new = True
            seen_ids.add(h["id"])
            no_new_rounds = 0

            # Expand truncated text & collapsed replies
            expand_visible_content(page)
            time.sleep(1)

            # Re-read after expansion
            full_text = get_main_text(page)
            cur_headings = get_thread_headings(page)

            thread_text = extract_thread_text(full_text, h["text"], cur_headings, h["id"])
            dates = extract_dates(thread_text, ref)

            # --- Stopping condition ---
            if dates and all(d < cutoff for d in dates):
                print(f"  Stop: '{h['text'][:50]}…' — all dates before {cutoff.date()}", file=sys.stderr)
                return conversations

            # Detect image attachments from heading or text
            has_images = h.get("has_images", False)
            if not has_images:
                has_images = bool(re.search(r'\d+\s+attachment', thread_text, re.IGNORECASE))

            # Get canonical thread URL via '...' menu → 'Copy link'
            thread_url = get_thread_url_via_copy_link(page, h["id"])
            if not thread_url:
                thread_url = h.get("thread_url", "")  # fallback to constructed URL

            conversations.append({
                "type": detect_type(thread_text),
                "heading": h["text"],
                "thread_url": thread_url,
                "has_images": has_images,
                "raw_text": thread_text,
                "dates": sorted(set(d.strftime("%Y-%m-%d %H:%M") for d in dates)),
                "most_recent": max(dates).strftime("%Y-%m-%d") if dates else None,
            })
            print(f"  [{len(conversations)}] {h['text'][:60]}…", file=sys.stderr)

        if not found_new:
            no_new_rounds += 1

        scroll_feed_down(page, 1200)
        time.sleep(3)

    print(f"  Finished: no more threads after {MAX_IDLE} idle scroll rounds", file=sys.stderr)
    return conversations


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) < 3:
        print("Usage: python engage_read_conversations.py <community_name> <days>", file=sys.stderr)
        sys.exit(1)

    community_name = sys.argv[1]
    days = int(sys.argv[2])

    config = load_config()
    cdp_url = config["edge_cdp"]["url"]
    ref = datetime.now()

    print(f"=== Reading '{community_name}' — last {days} day(s) ===", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(cdp_url)
        context = browser.contexts[0]
        page = context.new_page()

        try:
            navigate_to_community(page, community_name)
            verify_filters(page)
            conversations = read_conversations(page, days, ref)
        finally:
            page.close()

    result = {
        "community": community_name,
        "days": days,
        "cutoff_date": (ref - timedelta(days=days)).strftime("%Y-%m-%d"),
        "reference_date": ref.strftime("%Y-%m-%d %H:%M"),
        "total_conversations": len(conversations),
        "conversations": conversations,
    }

    sys.stdout.reconfigure(encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
