"""Download YouTube video transcript via Playwright CDP + Edge.

Connects to the running Edge debug instance, navigates to the YouTube video page,
opens the transcript panel, and scrapes all transcript segments from the DOM.
Outputs the transcript as a text file in the workspace root.

Usage:
    python yt_transcript.py <video_id_or_url> [--clean]

Arguments:
    video_id_or_url   YouTube video ID (e.g. 2C6G9M1aOko) or full URL
    --clean           Strip timestamps from the output (text only)

Output:
    yt_transcripts/yt_<VIDEO_ID>.txt

Requires:
    - Edge running with --remote-debugging-port=9222
    - Playwright for Python (pip install playwright)
"""
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp
import time, sys, re, json, os

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))
CDP_URL = CONFIG["edge_cdp"]["url"]


def parse_video_id(value):
    """Extract video ID from a URL or return as-is if already an ID."""
    m = re.search(r"(?:v=|youtu\.be/)([\w-]{11})", value)
    return m.group(1) if m else value


def open_transcript_panel(page):
    """Try multiple strategies to open the transcript panel."""
    # Strategy 1: click "Show transcript" button in expanded description
    try:
        more_desc = page.query_selector(
            "tp-yt-paper-button#expand, #description-inline-expander #expand"
        )
        if more_desc and more_desc.is_visible():
            more_desc.click()
            time.sleep(1)
    except Exception:
        pass

    transcript_btn = page.query_selector(
        "button:has-text('Show transcript'), button:has-text('Mostra trascrizione')"
    )
    if transcript_btn and transcript_btn.is_visible():
        transcript_btn.click()
        time.sleep(2)
        return True

    # Strategy 2: three-dot menu → Show transcript
    page.evaluate("""() => {
        const el = document.querySelector('#actions-inner, #menu-container');
        if (el) el.scrollIntoView({behavior: 'instant', block: 'center'});
    }""")
    time.sleep(1)

    more_btn = page.query_selector(
        'button[aria-label="More actions"], ytd-menu-renderer yt-button-shape button'
    )
    if more_btn:
        try:
            more_btn.scroll_into_view_if_needed(timeout=5000)
            time.sleep(0.5)
            more_btn.click(timeout=5000)
            time.sleep(1)
            items = page.query_selector_all("ytd-menu-service-item-renderer")
            for item in items:
                txt = item.inner_text().strip().lower()
                if "transcript" in txt or "trascrizione" in txt:
                    item.click()
                    time.sleep(2)
                    return True
        except Exception:
            pass

    # Strategy 3: JS fallback
    result = page.evaluate("""() => {
        const btn = document.querySelector('button[aria-label="Show transcript"]');
        if (btn) { btn.click(); return true; }
        return false;
    }""")
    if result:
        time.sleep(2)
        return True

    return False


def extract_transcript(page):
    """Extract transcript text from the transcript panel DOM."""
    selectors = [
        "ytd-transcript-segment-renderer",
        "#segments-container ytd-transcript-segment-renderer",
        "ytd-engagement-panel-section-list-renderer[target-id*='transcript'] .segment-text",
    ]

    for sel in selectors:
        segments = page.query_selector_all(sel)
        if segments:
            lines = []
            for seg in segments:
                text = seg.inner_text().strip()
                if text:
                    lines.append(text)
            if lines:
                return "\n".join(lines)

    # Fallback: entire transcript panel text
    panel = page.query_selector(
        "ytd-engagement-panel-section-list-renderer[target-id*='transcript']"
    )
    if panel:
        return panel.inner_text()

    return None


def clean_transcript(text):
    """Remove timestamp lines (e.g. '0:00', '12:34') leaving only spoken text."""
    lines = text.splitlines()
    cleaned = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        # Skip lines that are only a timestamp like "0:00" or "1:23:45"
        if re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", stripped):
            continue
        cleaned.append(stripped)
    return "\n".join(cleaned)


def main():
    if len(sys.argv) < 2:
        print("Usage: python yt_transcript.py <video_id_or_url> [--clean]")
        sys.exit(1)

    video_id = parse_video_id(sys.argv[1])
    do_clean = "--clean" in sys.argv
    video_url = f"https://www.youtube.com/watch?v={video_id}"

    with sync_playwright() as p:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.new_page()
        # Mute video elements before they start playing
        page.add_init_script("""
            new MutationObserver(() => {
                document.querySelectorAll('video').forEach(v => { v.muted = true; });
            }).observe(document.documentElement, {childList: true, subtree: true});
        """)

        try:
            page.goto(video_url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=20000)
            title = page.title().replace(" - YouTube", "").strip()
            print(f"Video: {title}")

            # Dismiss cookie consent if present
            try:
                consent = page.query_selector(
                    "button[aria-label*='Accept'], button:has-text('Accept all')"
                )
                if consent and consent.is_visible():
                    consent.click()
                    time.sleep(1)
            except Exception:
                pass

            # Scroll down to make video action area visible
            page.evaluate("window.scrollBy(0, 400)")
            time.sleep(1)

            # Check if transcript panel is already open (from a previous session)
            already_open = page.query_selector("ytd-transcript-segment-renderer")
            if not already_open:
                if not open_transcript_panel(page):
                    print("ERROR: Could not open transcript panel")
                    page.screenshot(path="debug_yt_transcript.png", full_page=False)
                    print("Screenshot saved to debug_yt_transcript.png")
                    sys.exit(1)

            transcript_text = extract_transcript(page)

            if not transcript_text:
                print("ERROR: Could not extract transcript text")
                page.screenshot(path="debug_yt_transcript.png", full_page=False)
                print("Screenshot saved to debug_yt_transcript.png")
                sys.exit(1)

            if do_clean:
                transcript_text = clean_transcript(transcript_text)

            out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "yt_transcripts")
            os.makedirs(out_dir, exist_ok=True)
            out_file = os.path.join(out_dir, f"yt_{video_id}.txt")
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(transcript_text)

            print(f"Saved {len(transcript_text)} chars to {out_file}")
            # Output as JSON for machine consumption
            print(json.dumps({
                "video_id": video_id,
                "title": title,
                "file": out_file,
                "chars": len(transcript_text),
            }))

        finally:
            page.close()


if __name__ == "__main__":
    main()
