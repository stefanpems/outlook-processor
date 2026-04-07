"""
Read replies/comments from a Viva Engage conversation thread page via Edge CDP.
Opens the thread URL, expands all replies/comments, and extracts the full text.

Usage: python ve-notifications-process.py <thread_url>

Output: JSON to stdout with thread_url, full_text (raw text of the thread).
"""
import json, re, os, sys, time
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]


def expand_all_content(page):
    """Click 'see more', 'N replies', 'Show X more answers' buttons to expand all content."""
    max_rounds = 10
    for _ in range(max_rounds):
        clicked = False
        buttons = page.query_selector_all("button")
        for btn in buttons:
            try:
                if not btn.is_visible():
                    continue
                txt = btn.inner_text().strip()
                lo = txt.lower()

                # Expand truncated post/reply body
                if lo == "see more":
                    btn.evaluate("e => e.click()")
                    time.sleep(0.5)
                    clicked = True
                    continue

                # Expand collapsed reply chains
                if len(txt) < 60 \
                   and ("repl" in lo or "answer" in lo or "comment" in lo) \
                   and any(c.isdigit() for c in txt) \
                   and "hide" not in lo and "collapse" not in lo and "less" not in lo:
                    btn.evaluate("e => e.click()")
                    time.sleep(1)
                    clicked = True
                    continue

                # "Show more" / "View more" buttons
                if lo in ("show more", "view more", "load more"):
                    btn.evaluate("e => e.click()")
                    time.sleep(1)
                    clicked = True
                    continue
            except Exception:
                pass

        if not clicked:
            break
        time.sleep(1)


def main():
    if len(sys.argv) < 2:
        print("Usage: python ve-notifications-process.py <thread_url>")
        sys.exit(1)

    thread_url = sys.argv[1]

    ensure_edge_cdp()
    p = sync_playwright().start()

    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        page = ctx.new_page()

        page.goto(thread_url, timeout=60000)
        time.sleep(5)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

        # Dismiss any overlay/popup
        try:
            page.keyboard.press("Escape")
            time.sleep(1)
        except Exception:
            pass

        # Expand all content
        expand_all_content(page)
        time.sleep(1)

        # Read full page text
        main_el = page.query_selector("main, [role=main]")
        full_text = main_el.inner_text() if main_el else page.inner_text("body")

        result = {
            "thread_url": thread_url,
            "full_text": full_text[:10000],
        }

        page.close()
        print(json.dumps(result, indent=2, ensure_ascii=False))

    except Exception as ex:
        print(json.dumps({"error": str(ex), "thread_url": thread_url}))
        sys.exit(1)

    finally:
        p.stop()


if __name__ == "__main__":
    main()
