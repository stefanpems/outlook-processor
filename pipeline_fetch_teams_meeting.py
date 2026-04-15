"""Fetch Teams meeting recording metadata and transcript from a SharePoint Stream page.
Connects to Edge via CDP, navigates to the video page (muted), extracts metadata
(title, date, duration) and transcript text.

Usage: python pipeline_fetch_teams_meeting.py <url>
Output: JSON to stdout with title, published_date, duration_seconds, duration_formatted,
        sha256_id, transcript_path, transcript_length.
        Optional: detected_meeting_sender (e.g. "LevelUp", "Ninja", "CCP", etc.)
        auto-detected from page content and URL patterns.
"""
import json, re, os, sys, hashlib, time
from urllib.parse import unquote
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))
CDP_URL = CONFIG["edge_cdp"]["url"]

TRANSCRIPTS_DIR = os.path.join(BASE, "teams_transcripts")
os.makedirs(TRANSCRIPTS_DIR, exist_ok=True)


def compute_sha256(url):
    """Compute SHA256 hex digest of the URL as-is (not decoded)."""
    return hashlib.sha256(url.encode("utf-8")).hexdigest()


def format_duration(seconds):
    """Format seconds as 'Xh YYm' or 'Ym'."""
    seconds = int(seconds)
    minutes = round(seconds / 60)
    if minutes >= 60:
        h = minutes // 60
        m = minutes % 60
        return f"{h}h {m:02d}m"
    return f"{minutes}m"


def parse_filename_from_url(url):
    """Extract the decoded filename from the id= query param."""
    m = re.search(r"[?&]id=([^&]+)", url)
    if m:
        path = unquote(m.group(1))
        # Last segment is the filename
        return path.rsplit("/", 1)[-1] if "/" in path else path
    return None


def extract_date_from_filename(filename):
    """Extract date from Teams recording filename.
    Pattern: ...-YYYYMMDD_HHMMSS-Meeting Recording.mp4
    """
    if not filename:
        return None
    m = re.search(r"(\d{4})(\d{2})(\d{2})_\d{6}[-\u2013]Meeting Recording", filename)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return None


def extract_title_from_filename(filename):
    """Extract meeting title from Teams recording filename.
    Strip the -YYYYMMDD_HHMMSS-Meeting Recording.mp4 suffix.
    """
    if not filename:
        return None
    cleaned = re.sub(
        r"[-\u2013]\d{8}_\d{6}[-\u2013]Meeting Recording\.mp4$",
        "",
        filename,
        flags=re.IGNORECASE,
    )
    return cleaned.strip() if cleaned != filename else None


def extract_transcript_with_scroll(page):
    """Extract full transcript by scrolling the virtualized ms-List inside Stream.

    The Stream transcript panel uses a virtualized list (ms-List) inside a
    scrollable FocusZone. Only ~30 entries are rendered at a time. We scroll
    to materialize all pages and collect entry text.
    """
    # Find the scrollable container inside #OneTranscript
    scroller_sel = """(() => {
        const candidates = document.querySelectorAll('#OneTranscript *');
        for (const el of candidates) {
            const style = window.getComputedStyle(el);
            if ((style.overflowY === 'auto' || style.overflowY === 'scroll') &&
                el.scrollHeight > el.clientHeight + 50) return el;
        }
        return null;
    })()"""

    info = page.evaluate(f"""() => {{
        const el = {scroller_sel};
        if (!el) return null;
        return {{ scrollHeight: el.scrollHeight, clientHeight: el.clientHeight }};
    }}""")

    if not info:
        # Fallback: try reading #OneTranscript text directly (non-virtualized)
        text = page.evaluate("""() => {
            const el = document.getElementById('OneTranscript');
            return el ? el.innerText.trim() : '';
        }""")
        return text

    scroll_height = info["scrollHeight"]
    client_height = info["clientHeight"]

    # If not virtualized (everything fits), just read it
    if scroll_height <= client_height + 50:
        text = page.evaluate("""() => {
            const el = document.getElementById('OneTranscript');
            return el ? el.innerText.trim() : '';
        }""")
        return text

    # Scroll to collect all entries from the virtualized list
    scroll_step = client_height - 50
    all_entries = {}
    max_iterations = 300

    for iteration in range(max_iterations):
        # Collect currently rendered entries
        batch = page.evaluate("""() => {
            const items = document.querySelectorAll('[id^="sub-entry-"]');
            const result = [];
            for (const item of items) {
                const m = item.id.match(/sub-entry-(\\d+)/);
                if (m) result.push({ i: parseInt(m[1]), t: item.innerText.trim() });
            }
            // Also get scroll position
            const candidates = document.querySelectorAll('#OneTranscript *');
            let scrollTop = 0, scrollHeight = 0, clientHeight = 0;
            for (const el of candidates) {
                const style = window.getComputedStyle(el);
                if ((style.overflowY === 'auto' || style.overflowY === 'scroll') &&
                    el.scrollHeight > el.clientHeight + 50) {
                    scrollTop = el.scrollTop;
                    scrollHeight = el.scrollHeight;
                    clientHeight = el.clientHeight;
                    break;
                }
            }
            return { entries: result, scrollTop, scrollHeight, clientHeight };
        }""")

        for e in batch["entries"]:
            if e["i"] not in all_entries:
                all_entries[e["i"]] = e["t"]

        # Check if we reached the bottom
        if batch["scrollTop"] + batch["clientHeight"] >= batch["scrollHeight"] - 20:
            break

        # Scroll down
        page.evaluate(f"""() => {{
            const candidates = document.querySelectorAll('#OneTranscript *');
            for (const el of candidates) {{
                const style = window.getComputedStyle(el);
                if ((style.overflowY === 'auto' || style.overflowY === 'scroll') &&
                    el.scrollHeight > el.clientHeight + 50) {{
                    el.scrollTop += {scroll_step};
                    return;
                }}
            }}
        }}""")
        page.wait_for_timeout(300)

    if not all_entries:
        return ""

    # Sort by index and join
    sorted_entries = sorted(all_entries.items())
    print(
        f"Transcript: {len(sorted_entries)} entries collected via scroll",
        file=sys.stderr,
    )
    return "\n".join(text for _, text in sorted_entries)


def fetch_teams_meeting(url):
    """Connect to Edge CDP, navigate to Stream page, extract metadata + transcript."""
    sha256_id = compute_sha256(url)
    transcript_path = os.path.join(TRANSCRIPTS_DIR, f"{sha256_id}.txt")

    # Parse filename from URL for fallback metadata
    filename = parse_filename_from_url(url)
    filename_date = extract_date_from_filename(filename)
    filename_title = extract_title_from_filename(filename)

    ensure_edge_cdp()
    p = sync_playwright().start()
    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        page = ctx.new_page()

        # Mute all media elements — same pattern as pipeline_fetch_video.py
        page.add_init_script("""
            // Mute video/audio elements as they appear
            new MutationObserver(() => {
                document.querySelectorAll('video, audio').forEach(el => {
                    el.muted = true;
                    el.volume = 0;
                });
            }).observe(document.documentElement, {childList: true, subtree: true});
            // Override play() to ensure muting
            const origPlay = HTMLMediaElement.prototype.play;
            HTMLMediaElement.prototype.play = function() {
                this.muted = true;
                this.volume = 0;
                return origPlay.call(this);
            };
        """)

        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass  # Stream pages keep loading; domcontentloaded is enough
            page.wait_for_timeout(5000)  # Extra wait for SPA / Stream player init

            # --- Extract title ---
            title = page.evaluate("""() => {
                const selectors = [
                    '[data-automation-id="TextBlock"][role="heading"]',
                    '.onePlayer-titleText',
                    '.StreamTitleComponent-title',
                    'h1.FileViewer-title',
                    'h1',
                    '[data-automation-id="titleField"]',
                ];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el && el.innerText.trim()) return el.innerText.trim();
                }
                return document.title
                    .replace(/ - Microsoft Stream.*$/i, '')
                    .replace(/ - .*\\.sharepoint\\.com$/i, '')
                    .replace(/\\.mp4$/i, '')
                    .trim();
            }""")

            if not title or title == "Stream" or len(title) < 3:
                title = filename_title or title
            # Strip trailing .mp4 if present
            title = re.sub(r"\.mp4$", "", title, flags=re.IGNORECASE).strip()

            print(f"Title: {title}", file=sys.stderr)

            # --- Extract published date ---
            # The metadata area below the video (title, date, views) loads with
            # a slight delay on Stream pages.  Wait a few seconds for it.
            page.wait_for_timeout(5000)

            page_date = page.evaluate("""() => {
                // Try various date selectors in Stream / SP
                const selectors = [
                    '[data-automation-id="detailsDate"]',
                    'time[datetime]',
                    '[data-automation-id="date"]',
                ];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el) {
                        const dt = el.getAttribute('datetime') || el.innerText;
                        if (dt && dt.trim()) return dt.trim();
                    }
                }
                const meta = document.querySelector(
                    'meta[name="created"], meta[property="article:published_time"]'
                );
                if (meta && meta.content) return meta.content;
                // Fallback: scan text near the title for a date string.
                // Stream renders "Month DD, YYYY" (e.g. "April 13, 2026") in
                // the metadata strip below the video title / above view count.
                const body = document.body.innerText || '';
                const m = body.match(
                    /(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s*\d{4}/
                );
                if (m) return m[0];
                return '';
            }""")

            published_date = ""
            if page_date:
                # Try ISO-like YYYY-MM-DD or YYYY/MM/DD
                m = re.search(r"(\d{4})[-/](\d{2})[-/](\d{2})", page_date)
                if m:
                    published_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                else:
                    # Try "Month DD, YYYY" (English)
                    import datetime as _dt
                    for fmt in ("%B %d, %Y", "%b %d, %Y"):
                        try:
                            d = _dt.datetime.strptime(page_date.strip(), fmt)
                            published_date = d.strftime("%Y-%m-%d")
                            break
                        except ValueError:
                            continue
            if not published_date:
                published_date = filename_date or ""

            print(f"Published: {published_date}", file=sys.stderr)

            # --- Extract duration from <video> element ---
            duration_seconds = page.evaluate("""() => {
                const video = document.querySelector('video');
                if (video && video.duration && isFinite(video.duration)) {
                    return Math.round(video.duration);
                }
                return 0;
            }""")

            # If duration not available yet, briefly play to load metadata
            if not duration_seconds:
                print("Duration not in DOM, playing briefly to load metadata...",
                      file=sys.stderr)
                page.evaluate("""() => {
                    const video = document.querySelector('video');
                    if (video) {
                        video.muted = true;
                        video.volume = 0;
                        video.play().catch(() => {});
                    }
                }""")
                page.wait_for_timeout(4000)
                duration_seconds = page.evaluate("""() => {
                    const video = document.querySelector('video');
                    if (video) {
                        video.pause();
                        if (video.duration && isFinite(video.duration)) {
                            return Math.round(video.duration);
                        }
                    }
                    return 0;
                }""")

            # Also try reading duration text from the player UI
            if not duration_seconds:
                dur_text = page.evaluate("""() => {
                    const selectors = [
                        '.onePlayer-durationDisplay',
                        '[data-automation-id="duration"]',
                        '.vjs-duration-display',
                        '.vjs-remaining-time-display',
                        '.duration',
                    ];
                    for (const sel of selectors) {
                        const el = document.querySelector(sel);
                        if (el && el.innerText.trim()) return el.innerText.trim();
                    }
                    return '';
                }""")
                if dur_text:
                    parts = dur_text.replace("-", "").strip().split(":")
                    try:
                        if len(parts) == 3:
                            duration_seconds = (
                                int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
                            )
                        elif len(parts) == 2:
                            duration_seconds = int(parts[0]) * 60 + int(parts[1])
                    except ValueError:
                        pass

            duration_formatted = format_duration(duration_seconds) if duration_seconds else ""
            print(f"Duration: {duration_formatted} ({duration_seconds}s)", file=sys.stderr)

            # --- Transcript extraction ---
            transcript_text = ""

            # Try clicking a "Transcript" button/tab in the player
            transcript_opened = page.evaluate("""() => {
                const btns = document.querySelectorAll('button, [role="tab"], [role="button"]');
                for (const btn of btns) {
                    const text = (btn.innerText || '').toLowerCase();
                    const label = (btn.getAttribute('aria-label') || '').toLowerCase();
                    if (text.includes('transcript') || label.includes('transcript') ||
                        text.includes('trascrizione') || label.includes('trascrizione')) {
                        btn.click();
                        return 'clicked: ' + (btn.innerText || label).trim();
                    }
                }
                return '';
            }""")

            if transcript_opened:
                print(f"Transcript button: {transcript_opened}", file=sys.stderr)
                page.wait_for_timeout(3000)

            # --- Stream transcript uses a virtualized ms-List inside a scrollable
            # FocusZone. We must scroll to materialize all pages, then collect
            # the text from each entry.
            transcript_text = extract_transcript_with_scroll(page)

            # Save transcript
            permission_denied = False
            if transcript_text:
                # Check for permission-denied messages — only on short texts
                # (actual permission pages are brief; long transcripts may
                # legitimately contain phrases like "request access")
                lower = transcript_text.lower()
                is_short = len(transcript_text) < 500
                if is_short and ("don't have permission" in lower
                        or "non hai l'autorizzazione" in lower
                        or "request access" in lower):
                    print(
                        "WARNING: Transcript permission denied — "
                        + transcript_text.strip()[:120],
                        file=sys.stderr,
                    )
                    transcript_text = ""
                    transcript_path = ""
                    permission_denied = True
                else:
                    with open(transcript_path, "w", encoding="utf-8") as f:
                        f.write(transcript_text)
                    print(
                        f"Transcript saved: {transcript_path} ({len(transcript_text)} chars)",
                        file=sys.stderr,
                    )
            else:
                print("WARNING: Could not extract transcript from DOM", file=sys.stderr)
                page.screenshot(path="debug_teams_meeting.png", full_page=False)
                print("Debug screenshot: debug_teams_meeting.png", file=sys.stderr)
                transcript_path = ""

            # --- Auto-detect meeting sender from page content & URL ---
            detected_meeting_sender = ""
            try:
                page_text = page.evaluate("() => document.body.innerText || ''")
                page_url_decoded = unquote(url).lower()

                if re.search(r"levelup", page_text, re.IGNORECASE):
                    detected_meeting_sender = "LevelUp"
                elif re.search(r"bootcamp", page_text, re.IGNORECASE):
                    detected_meeting_sender = "Ninja"
                elif (page_url_decoded.startswith(
                        "https://microsoft-my.sharepoint.com/personal/customerconnection_microsoft_com/")
                      or "Accelerated Collaboration Forum" in page_text):
                    detected_meeting_sender = "CCP"
                elif (page_url_decoded.startswith(
                        "https://microsoft.sharepoint.com/sites/securitysolutions")
                      or "Security Global Connection Call" in page_text):
                    detected_meeting_sender = "SSA"
                elif (page_url_decoded.startswith(
                        "https://microsoft.sharepoint.com/teams/azureadidentitychamps")
                      or "Entra Expert Connect" in page_text):
                    detected_meeting_sender = "EEC"
                elif (page_url_decoded.startswith(
                        "https://microsoft.sharepoint.com/teams/identity-cc/")
                      or "Identity Connected Community" in page_text):
                    detected_meeting_sender = "IIC"
                elif (re.search(r"sentinel", page_text, re.IGNORECASE)
                      and "Office Hours" in page_text):
                    detected_meeting_sender = "Sentinel"
                elif ("Office Hours" in page_text
                      and (re.search(r"Defender for Cloud Apps", page_text, re.IGNORECASE)
                           or re.search(r"\bMDA\b", page_text))):
                    detected_meeting_sender = "MDA"
                elif ("Office Hours" in page_text
                      and (re.search(r"Defender for Cloud", page_text, re.IGNORECASE)
                           or re.search(r"\bMDC\b", page_text))):
                    detected_meeting_sender = "MDC"
                elif ("Office Hours" in page_text
                      and (re.search(r"Defender for Endpoint", page_text, re.IGNORECASE)
                           or re.search(r"\bMDE\b", page_text))):
                    detected_meeting_sender = "MDE"
                elif ("Office Hours" in page_text
                      and (re.search(r"Defender for Identity", page_text, re.IGNORECASE)
                           or re.search(r"\bMDI\b", page_text))):
                    detected_meeting_sender = "MDI"
                elif ("Field Connection Forum" in page_text
                      or re.search(r"Defender for Office", page_text, re.IGNORECASE)
                      or re.search(r"\bMDO\b", page_text)):
                    detected_meeting_sender = "MDO"
                elif (page_url_decoded.startswith(
                        "https://microsoft-my.sharepoint.com/personal/deverett_microsoft_com")
                      or "Deep Dive" in page_text):
                    detected_meeting_sender = "ECS"
                elif "?id=%2fsites%2froadmaphub%2fshared%20documents%2fvideos%2fidentity%20%26%20network%20access%2f" in page_url_decoded:
                    detected_meeting_sender = "IdAdv"

                if detected_meeting_sender:
                    print(f"Auto-detected meeting sender: {detected_meeting_sender}",
                          file=sys.stderr)
            except Exception:
                pass

            result = {
                "title": title,
                "published_date": published_date,
                "duration_seconds": int(duration_seconds) if duration_seconds else 0,
                "duration_formatted": duration_formatted,
                "sha256_id": sha256_id,
                "transcript_path": transcript_path,
                "transcript_length": len(transcript_text),
                "transcript_permission_denied": permission_denied,
                "url": url,
            }
            if detected_meeting_sender:
                result["detected_meeting_sender"] = detected_meeting_sender
            return result

        finally:
            page.close()
    finally:
        p.stop()


def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline_fetch_teams_meeting.py <url>")
        sys.exit(1)

    url = sys.argv[1]
    result = fetch_teams_meeting(url)
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
