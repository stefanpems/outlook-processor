"""
Build an HTML email digest of blog posts from the SharePoint BlogPosts list,
grouped by topic, and save it to the output directory.

Usage:
  python pipeline_email_report.py [--from-date YYYY-MM-DD] [--to-date YYYY-MM-DD] [--tech "tech1,tech2"]

Defaults:
  --from-date : yesterday
  --to-date   : yesterday
  --tech      : (empty = all items)

Output: JSON to stdout with html_path, subject, total_items, topics_count.
"""
import json, re, os, sys, argparse, html as html_mod
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SP_API = CONFIG["sharepoint"]["blog_list_api"]
SP_LIST_URL = CONFIG["sharepoint"]["blog_list_url"]
COLOR_PALETTE = CONFIG.get("topic_color_palette", [
    "#F0E6D3", "#D3E8F0", "#D3F0D6", "#F0D3E6", "#E6F0D3", "#D3D8F0",
    "#F0DAD3", "#D3F0EA", "#E8D3F0", "#F0F0D3", "#D3EAF0", "#E6D3F0",
    "#F0D3D3", "#D3F0D3", "#D3D3F0", "#F0ECD3", "#F0D3EC", "#D3F0F0",
])

TOPIC_COLORS = {}


def get_topic_color(topic):
    if topic not in TOPIC_COLORS:
        idx = len(TOPIC_COLORS) % len(COLOR_PALETTE)
        TOPIC_COLORS[topic] = COLOR_PALETTE[idx]
    return TOPIC_COLORS[topic]


def parse_args():
    parser = argparse.ArgumentParser(description="Build blog digest HTML from SP.")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    parser.add_argument("--from-date", default=yesterday,
                        help="Start date YYYY-MM-DD (default: yesterday)")
    parser.add_argument("--to-date", default=yesterday,
                        help="End date YYYY-MM-DD (default: yesterday)")
    parser.add_argument("--tech", default="",
                        help="Comma-separated technologies to filter by")
    parser.add_argument("--recipients", default="",
                        help="Semicolon-separated email addresses (default: from config.json)")
    return parser.parse_args()


def fetch_items(page, date_from_dot, date_to_dot, tech_filter):
    """Fetch SP items with expanded lookups, filtered by date range."""
    all_items = []
    last_id = 0
    filter_clause = f"field_0 ge '{date_from_dot}' and field_0 le '{date_to_dot}'"

    while True:
        full_filter = f"Id gt {last_id} and {filter_clause}"
        url = (
            f"{SP_API}/items?$top=500"
            f"&$filter={full_filter}"
            f"&$select=Id,Title,Link,Summary,field_0,SourceNew/Title,Tech/Title"
            f"&$expand=SourceNew,Tech"
            f"&$orderby=Id"
        )

        batch = page.evaluate("""async (url) => {
            const resp = await fetch(url, {
                headers: { "Accept": "application/json;odata=nometadata" }
            });
            return await resp.json();
        }""", url)

        items = batch.get("value", [])
        if not items:
            break
        all_items.extend(items)
        last_id = items[-1]["Id"]
        print(f"  Fetched {len(all_items)} items (last ID={last_id})...", file=sys.stderr)
        if len(items) < 500:
            break

    # Process items
    results = []
    for item in all_items:
        link = item.get("Link", {})
        url = link.get("Url", "") if isinstance(link, dict) else str(link or "")

        source_new = item.get("SourceNew")
        topic = source_new.get("Title", "") if isinstance(source_new, dict) else ""

        tech_list = item.get("Tech", []) or []
        techs = []
        if isinstance(tech_list, list):
            for t in tech_list:
                if isinstance(t, dict):
                    techs.append(t.get("Title", ""))

        # Apply tech filter (case-insensitive substring match on tech labels)
        if tech_filter:
            tech_lower = [t.lower() for t in techs]
            if not any(tf.lower() in tech_lower for tf in tech_filter):
                continue

        results.append({
            "id": item.get("Id"),
            "title": item.get("Title", ""),
            "url": url,
            "summary": item.get("Summary", "") or "",
            "published": item.get("field_0", ""),
            "topic": topic or "Unknown",
            "techs": techs,
        })

    return results


# ---------------------------------------------------------------------------
# CSS — compatible with Outlook / OWA email rendering
# ---------------------------------------------------------------------------
CSS = """\
body {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  background-color: #f4f5f7; color: #1a1a2e;
  line-height: 1.6; margin: 0; padding: 0;
}
.wrapper {
  max-width: 960px; margin: 0 auto; padding: 24px 16px;
}
h1 {
  color: #1a1a2e; font-size: 24px; font-weight: 700;
  border-bottom: 3px solid #4361ee;
  padding-bottom: 8px; margin: 0 0 20px 0;
}
.stats-table td {
  padding: 8px 20px 8px 0; font-size: 14px; color: #333;
}
.stats-table b { color: #4361ee; }
.stats-bar {
  background-color: #eef0f8; padding: 10px 16px;
  border-radius: 6px; margin-bottom: 24px;
}
.toc-box {
  background-color: #ffffff; padding: 16px 20px;
  border: 1px solid #e0e2e8; border-radius: 6px;
  margin-bottom: 24px;
}
.toc-box h2 {
  font-size: 16px; color: #4361ee;
  margin: 0 0 10px 0;
}
.toc-box ul {
  list-style-type: disc; margin: 0; padding: 0 0 0 24px;
}
.toc-box li {
  padding: 4px 0; font-size: 14px;
}
.toc-box a { color: #4361ee; text-decoration: none; font-weight: 600; }
.toc-box .count { color: #888; font-size: 13px; margin-left: 4px; }
.topic-section {
  background-color: #ffffff; border: 1px solid #e0e2e8;
  border-radius: 6px; padding: 16px 20px; margin-bottom: 20px;
}
.topic-header {
  font-size: 18px; color: #1a1a2e; font-weight: 700;
  margin: 0 0 12px 0;
}
.badge {
  background-color: #4361ee; color: #ffffff; font-size: 12px;
  padding: 2px 10px; border-radius: 10px;
  margin-left: 8px; font-weight: 400;
}
.article {
  border-bottom: 1px solid #eee; padding: 12px 0;
}
.article:last-child { border-bottom: none; }
.article-title {
  font-size: 15px; font-weight: 600; margin: 0;
}
.article-title a { color: #4361ee; text-decoration: none; }
.article-meta {
  font-size: 13px; color: #666; margin: 4px 0 8px 0;
}
.date { font-weight: 600; }
.tag {
  background-color: #e8eaf6; padding: 2px 10px; border-radius: 10px;
  font-size: 12px; color: #4361ee; margin-left: 6px;
  display: inline-block;
}
.article-summary {
  font-size: 14px; line-height: 1.55; color: #333;
}
.article-summary b { color: #1a1a2e; }
.article-summary ul { margin: 6px 0 6px 20px; padding: 0; }
.article-summary li { margin-bottom: 3px; }
.back-link {
  display: inline-block; margin-top: 10px; color: #4361ee;
  font-size: 13px; text-decoration: none; font-weight: 600;
}
.footer-bar {
  text-align: center; color: #999; font-size: 12px;
  margin-top: 28px; padding-top: 12px;
  border-top: 1px solid #e0e0e0;
}
"""


def build_html(items, date_from, date_to, tech_filter):
    """Build HTML report grouped by topic."""

    # Group by topic
    topics = {}
    for item in items:
        topics.setdefault(item["topic"], []).append(item)

    sorted_topics = sorted(topics.keys())

    # Sort articles: published date descending, then title ascending (stable sort)
    for t in sorted_topics:
        topics[t].sort(key=lambda e: e.get("title", "").lower())
        topics[t].sort(key=lambda e: e.get("published", "") or "0000.00.00", reverse=True)

    date_from_dot = date_from.replace("-", ".")
    date_to_dot = date_to.replace("-", ".")
    date_label_html = f"From: {date_from_dot} To: {date_to_dot}"

    h = []
    h.append('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">')
    h.append(f'<title>PescoPedia Blog Digest - {date_label_html}</title>')
    h.append(f'<style>\n{CSS}</style>')
    h.append('</head><body>')
    h.append('<div class="wrapper">')

    # Header
    h.append(f'<h1>PescoPedia Blog Digest</h1>')
    h.append(f'<p style="font-size:15px;color:#555;margin:-12px 0 20px 0;">{date_label_html}</p>')

    # Stats — use a table for email-safe spacing
    h.append('<div class="stats-bar"><table class="stats-table"><tr>')
    h.append(f'<td>Articles: <b>{len(items)}</b></td>')
    h.append(f'<td>Topics: <b>{len(sorted_topics)}</b></td>')
    if tech_filter:
        h.append(f'<td>Tech filter: '
                 f'<b>{html_mod.escape(", ".join(tech_filter))}</b></td>')
    h.append('</tr></table></div>')

    # Table of contents
    h.append('<div class="toc-box" id="toc"><h2>Topics</h2><ul>')
    for t in sorted_topics:
        tid = re.sub(r'[^a-zA-Z0-9]', '-', t).lower()
        h.append(f'<li><a href="#{tid}">{html_mod.escape(t)}</a>'
                 f'<span class="count"> ({len(topics[t])})</span></li>')
    h.append('</ul></div>')

    # Separator between TOC and topic sections
    h.append('<hr style="border:none;border-top:2px solid #e0e2e8;margin:28px 0;">')

    # Topic sections
    for t in sorted_topics:
        tid = re.sub(r'[^a-zA-Z0-9]', '-', t).lower()
        color = get_topic_color(t)

        h.append(f'<div class="topic-section" id="{tid}" '
                 f'style="border-left: 4px solid {color};">')
        h.append(f'<h2 class="topic-header">{html_mod.escape(t)}'
                 f'<span class="badge">{len(topics[t])}</span></h2>')

        for item in topics[t]:
            title_esc = html_mod.escape(item["title"])
            url = item["url"]
            pub = item["published"]
            techs = item["techs"]
            summary = item["summary"]

            h.append('<div class="article">')

            # Title (clickable)
            if url:
                url_esc = html_mod.escape(url)
                h.append(f'<p class="article-title">'
                         f'<a href="{url_esc}" target="_blank">{title_esc}</a></p>')
            else:
                h.append(f'<p class="article-title">{title_esc}</p>')

            # Meta: date + tech tags
            meta_parts = []
            if pub:
                meta_parts.append(f'<span class="date">{html_mod.escape(pub)}</span>')
            for tech in techs:
                meta_parts.append(f'<span class="tag">{html_mod.escape(tech)}</span>')
            if meta_parts:
                h.append(f'<p class="article-meta">{" ".join(meta_parts)}</p>')

            # Summary (may contain HTML formatting from pipeline)
            if summary:
                h.append(f'<div class="article-summary">{summary}</div>')

            h.append('</div>')

        h.append('<a href="#toc" class="back-link">&uarr; Back to index</a>')
        h.append('</div>')

    h.append('<div class="footer-bar">Generated by Blog Digest Pipeline</div>')
    h.append('</div>')  # wrapper
    h.append('</body></html>')

    return '\n'.join(h)


def main():
    args = parse_args()
    date_from = args.from_date
    date_to = args.to_date
    tech_filter = ([t.strip() for t in args.tech.split(",") if t.strip()]
                   if args.tech else [])
    recipients = (args.recipients if args.recipients
                  else CONFIG.get("email_report", {}).get("default_recipients", ""))

    date_from_dot = date_from.replace("-", ".")
    date_to_dot = date_to.replace("-", ".")

    print(f"Fetching SP items from {date_from_dot} to {date_to_dot}...", file=sys.stderr)
    if tech_filter:
        print(f"Tech filter: {', '.join(tech_filter)}", file=sys.stderr)

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        sp_page = ctx.new_page()
        sp_page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
        sp_page.wait_for_timeout(5000)

        items = fetch_items(sp_page, date_from_dot, date_to_dot, tech_filter)
        sp_page.close()
    finally:
        p.stop()

    print(f"Found {len(items)} items.", file=sys.stderr)

    if not items:
        print(json.dumps({
            "html_path": "",
            "subject": "",
            "total_items": 0,
            "topics_count": 0,
            "date_from": date_from,
            "date_to": date_to,
            "message": "No items found for the specified date range and filters."
        }, indent=2))
        return

    html_content = build_html(items, date_from, date_to, tech_filter)

    # Save to output directory
    os.makedirs(os.path.join(BASE, "output"), exist_ok=True)
    html_filename = f"Blog_Notifications-Digest-From-{date_from_dot}-To-{date_to_dot}.html"
    html_path = os.path.join(BASE, "output", html_filename)
    if os.path.exists(html_path):
        nn = 2
        while True:
            html_filename = f"Blog_Notifications-Digest-From-{date_from_dot}-To-{date_to_dot}-{nn:02d}.html"
            html_path = os.path.join(BASE, "output", html_filename)
            if not os.path.exists(html_path):
                break
            nn += 1
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Build subject line
    subject = f"PescoPedia Blog Digest - From: {date_from_dot} To: {date_to_dot}"

    topics_count = len(set(item["topic"] for item in items))

    result = {
        "html_path": html_path,
        "subject": subject,
        "recipients": recipients,
        "total_items": len(items),
        "topics_count": topics_count,
        "date_from": date_from,
        "date_to": date_to,
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
