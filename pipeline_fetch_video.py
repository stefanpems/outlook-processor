"""
Fetch video metadata from a YouTube page via Playwright CDP.
Extracts: title, published date, duration, description, chapters.

Usage: python pipeline_fetch_video.py <url>
Output: JSON to stdout with video_id, title, published_date, duration_seconds,
        duration_formatted, description, chapters.
"""
import json, re, os, sys
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))
CDP_URL = CONFIG["edge_cdp"]["url"]


def parse_video_id(url):
    """Extract YouTube video ID from URL."""
    m = re.search(r"(?:v=|youtu\.be/)([\w-]{11})", url)
    return m.group(1) if m else None


def format_duration(seconds):
    """Format seconds as 'Xh YYm' or 'Ym'. Rounds seconds to nearest minute."""
    seconds = int(seconds)
    minutes = round(seconds / 60)
    if minutes >= 60:
        h = minutes // 60
        m = minutes % 60
        return f"{h}h {m:02d}m"
    return f"{minutes}m"


def parse_chapters_from_text(text, video_id):
    """Parse chapter timestamps from description text.
    Chapters look like: '0:00 Introduction' or '01:23:45 Topic Name'
    """
    chapters = []
    for line in text.split('\n'):
        line = line.strip()
        m = re.match(r'^(\d{1,2}:\d{2}(?::\d{2})?)\s+(.+)', line)
        if m:
            time_str = m.group(1)
            title = m.group(2).strip()
            parts = time_str.split(':')
            if len(parts) == 3:
                secs = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
            else:
                secs = int(parts[0]) * 60 + int(parts[1])
            chapters.append({
                "title": title,
                "time": time_str,
                "seconds": secs,
                "url": f"https://www.youtube.com/watch?v={video_id}&t={secs}s"
            })
    return chapters


def fetch_video_metadata(url):
    """Connect to Edge CDP, navigate to YouTube page, extract metadata."""
    video_id = parse_video_id(url)
    if not video_id:
        return {"error": "Not a YouTube URL", "url": url, "video_id": None}

    canonical_url = f"https://www.youtube.com/watch?v={video_id}"

    ensure_edge_cdp()
    p = sync_playwright().start()
    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        page = ctx.new_page()
        # Force-mute via CDP: override volume/muted setters so YouTube cannot unmute
        _MUTE_SCRIPT = """
            // Lock volume to 0 and muted to true at prototype level
            const vDesc = Object.getOwnPropertyDescriptor(HTMLMediaElement.prototype, 'volume');
            Object.defineProperty(HTMLMediaElement.prototype, 'volume', {
                get: function() { return 0; },
                set: function(v) { if (vDesc && vDesc.set) vDesc.set.call(this, 0); },
                configurable: true
            });
            const mDesc = Object.getOwnPropertyDescriptor(HTMLMediaElement.prototype, 'muted');
            Object.defineProperty(HTMLMediaElement.prototype, 'muted', {
                get: function() { return true; },
                set: function(v) { if (mDesc && mDesc.set) mDesc.set.call(this, true); },
                configurable: true
            });
            // Also intercept AudioContext to silence Web Audio API
            const OrigAC = window.AudioContext || window.webkitAudioContext;
            if (OrigAC) {
                const origResume = OrigAC.prototype.resume;
                OrigAC.prototype.resume = function() {
                    this.suspend(); return Promise.resolve();
                };
            }
        """
        cdp = page.context.new_cdp_session(page)
        cdp.send("Page.addScriptToEvaluateOnNewDocument", {"source": _MUTE_SCRIPT})

        try:
            page.goto(canonical_url, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=20000)

            # Dismiss cookie consent if present
            try:
                consent = page.query_selector(
                    "button[aria-label*='Accept'], button:has-text('Accept all')"
                )
                if consent and consent.is_visible():
                    consent.click()
                    page.wait_for_timeout(1000)
            except Exception:
                pass

            # Scroll down to make description area visible
            page.evaluate("window.scrollBy(0, 400)")
            page.wait_for_timeout(1000)

            # Try to expand description
            try:
                expand_btn = page.query_selector(
                    "tp-yt-paper-button#expand, "
                    "#description-inline-expander #expand, "
                    "#expand[role='button']"
                )
                if expand_btn and expand_btn.is_visible():
                    expand_btn.click()
                    page.wait_for_timeout(1500)
            except Exception:
                pass

            # Extract metadata from ytInitialPlayerResponse + DOM
            metadata = page.evaluate("""() => {
                const result = {};

                // From ytInitialPlayerResponse (global JS variable)
                if (typeof ytInitialPlayerResponse !== 'undefined') {
                    const vd = ytInitialPlayerResponse.videoDetails || {};
                    result.title = vd.title || '';
                    result.lengthSeconds = parseInt(vd.lengthSeconds || '0', 10);
                    result.shortDescription = vd.shortDescription || '';

                    const mf = ytInitialPlayerResponse.microformat;
                    if (mf && mf.playerMicroformatRenderer) {
                        result.publishDate = mf.playerMicroformatRenderer.publishDate || '';
                        result.category = mf.playerMicroformatRenderer.category || '';
                    }
                }

                // Fallback: from meta tags
                if (!result.title) {
                    const og = document.querySelector('meta[property="og:title"]');
                    result.title = og ? og.content : document.title.replace(' - YouTube', '');
                }
                if (!result.publishDate) {
                    const dp = document.querySelector('meta[itemprop="datePublished"]');
                    if (dp) result.publishDate = dp.content;
                }

                // Get description text from DOM (after expansion)
                const descEl = document.querySelector(
                    '#description-inline-expander, ytd-text-inline-expander'
                );
                if (descEl) {
                    result.descriptionText = descEl.innerText;
                }

                // Get chapters from DOM if available
                const chapterEls = document.querySelectorAll(
                    'ytd-macro-markers-list-item-renderer'
                );
                if (chapterEls.length > 0) {
                    result.chapters = [];
                    chapterEls.forEach(el => {
                        const titleEl = el.querySelector(
                            '#details h4, #details .macro-markers'
                        );
                        const timeEl = el.querySelector('#time');
                        if (titleEl && timeEl) {
                            result.chapters.push({
                                title: titleEl.innerText.trim(),
                                time: timeEl.innerText.trim()
                            });
                        }
                    });
                }

                return result;
            }""")

            # Parse chapters from description text (most complete source)
            desc_text = metadata.get('descriptionText') or metadata.get('shortDescription') or ''
            desc_chapters = parse_chapters_from_text(desc_text, video_id)

            if desc_chapters:
                # Prefer description chapters — always the complete list
                metadata['chapters'] = desc_chapters
            elif metadata.get('chapters'):
                # Fallback: enrich DOM chapters with seconds and URL
                for ch in metadata.get('chapters', []):
                    ts = ch.get('time', '')
                    parts = ts.split(':')
                    try:
                        if len(parts) == 3:
                            secs = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
                        elif len(parts) == 2:
                            secs = int(parts[0]) * 60 + int(parts[1])
                        else:
                            secs = 0
                    except ValueError:
                        secs = 0
                    ch['seconds'] = secs
                    ch['url'] = f"https://www.youtube.com/watch?v={video_id}&t={secs}s"

            # Format published date
            pub_date = metadata.get('publishDate', '')
            if pub_date:
                m = re.match(r'(\d{4})-(\d{2})-(\d{2})', pub_date)
                if m:
                    pub_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

            duration_secs = metadata.get('lengthSeconds', 0)

            result = {
                "video_id": video_id,
                "url": canonical_url,
                "title": metadata.get('title', ''),
                "published_date": pub_date,
                "duration_seconds": duration_secs,
                "duration_formatted": format_duration(duration_secs),
                "description": desc_text,
                "chapters": metadata.get('chapters', []),
            }

            return result

        finally:
            page.close()
    finally:
        p.stop()


def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline_fetch_video.py <url>")
        sys.exit(1)

    url = sys.argv[1]
    result = fetch_video_metadata(url)
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
