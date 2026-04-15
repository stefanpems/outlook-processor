"""
Build an HTML email digest of Teams meeting recordings from the SharePoint
VideosMSInt list, grouped by topic, and save it to the output directory.

Usage:
  python pipeline_teams_email_report.py [--from-date YYYY-MM-DD] [--to-date YYYY-MM-DD]
                                        [--recipients "a@x.com;b@x.com"]

Defaults:
  --from-date : yesterday
  --to-date   : yesterday

Output: JSON to stdout with html_path, subject, recipients, total_items, topics_count.
"""
import json, re, os, sys, argparse, html as html_mod
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
TM_CFG = CONFIG["teams_meeting"]
SP_API = TM_CFG["list_api"]
SP_LIST_URL = TM_CFG["list_url_ghcpview"]
FIELDS = TM_CFG["fields"]
F_PUBLISHED = FIELDS["published"]
F_SUMMARY = FIELDS["summary"]
F_DURATION = FIELDS["duration"]
F_LONG_LINK = FIELDS["long_link"]
SOURCE_DISPLAY = TM_CFG.get("source_display_names", {})
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
    parser = argparse.ArgumentParser(
        description="Build Teams recordings digest HTML from SP VideosMSInt.")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    parser.add_argument("--from-date", default=yesterday,
                        help="Start date YYYY-MM-DD (default: yesterday)")
    parser.add_argument("--to-date", default=yesterday,
                        help="End date YYYY-MM-DD (default: yesterday)")
    parser.add_argument("--recipients", default="",
                        help="Semicolon-separated email addresses (default: from config.json)")
    return parser.parse_args()


def fetch_items(page, date_from_dot, date_to_dot):
    """Fetch SP VideosMSInt items with expanded lookups, filtered by date range."""
    all_items = []
    last_id = 0
    filter_clause = (
        f"{F_PUBLISHED} ge '{date_from_dot}' and {F_PUBLISHED} le '{date_to_dot}'"
    )

    while True:
        full_filter = f"Id gt {last_id} and {filter_clause}"
        url = (
            f"{SP_API}/items?$top=500"
            f"&$filter={full_filter}"
            f"&$select=Id,Title,Link,{F_SUMMARY},{F_PUBLISHED},{F_DURATION},"
            f"{F_LONG_LINK},SourceNew/Title,SourceNew/Description,Tech/Title"
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
        print(f"  Fetched {len(all_items)} items (last ID={last_id})...",
              file=sys.stderr)
        if len(items) < 500:
            break

    # Process items
    results = []
    for item in all_items:
        # RecordingLink: prefer LongLink, fall back to Link.Url
        long_link = item.get(F_LONG_LINK, "") or ""
        link_obj = item.get("Link", {})
        short_link = (link_obj.get("Url", "") if isinstance(link_obj, dict)
                      else str(link_obj or ""))
        recording_link = long_link if long_link else short_link

        source = item.get("SourceNew")
        topic_raw = source.get("Title", "") if isinstance(source, dict) else ""
        topic_desc = source.get("Description", "") if isinstance(source, dict) else ""
        topic = SOURCE_DISPLAY.get(topic_raw, topic_desc or topic_raw)

        tech_list = item.get("Tech", []) or []
        techs = []
        if isinstance(tech_list, list):
            for t in tech_list:
                if isinstance(t, dict):
                    techs.append(t.get("Title", ""))

        results.append({
            "id": item.get("Id"),
            "title": item.get("Title", ""),
            "url": recording_link,
            "summary": item.get(F_SUMMARY, "") or "",
            "published": item.get(F_PUBLISHED, ""),
            "topic": topic or "Unknown",
            "techs": techs,
            "duration": item.get(F_DURATION, "") or "",
        })

    return results


# ---------------------------------------------------------------------------
# CSS — compatible with Outlook / OWA email rendering
# (same as pipeline_video_email_report.py)
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
.duration-tag {
  background-color: #fff3e0; padding: 2px 10px; border-radius: 10px;
  font-size: 12px; color: #e65100; margin-left: 6px;
  display: inline-block;
}
.article-summary {
  font-size: 14px; line-height: 1.55; color: #333;
}
.article-summary b { color: #1a1a2e; }
.article-summary ul { margin: 6px 0 6px 20px; padding: 0; }
.article-summary li { margin-bottom: 3px; }
.article-summary a { color: #4361ee; text-decoration: none; }
.back-link {
  display: inline-block; margin-top: 10px; color: #4361ee;
  font-size: 13px; text-decoration: none; font-weight: 600;
}
.footer-bar {
  text-align: center; color: #999; font-size: 12px;
  margin-top: 28px; padding-top: 12px;
  border-top: 1px solid #e0e0e0;
}
.info-box {
  background-color: #e8f4fd; border: 1px solid #b3d7f2;
  border-left: 4px solid #2196f3; border-radius: 6px;
  padding: 14px 18px; margin-bottom: 14px;
  font-size: 14px; color: #1a1a2e; line-height: 1.55;
}
.warning-box {
  background-color: #fff8e1; border: 1px solid #ffe0a0;
  border-left: 4px solid #ff9800; border-radius: 6px;
  padding: 14px 18px; margin-bottom: 14px;
  font-size: 14px; color: #1a1a2e; line-height: 1.55;
}
.info-box .box-icon, .warning-box .box-icon {
  font-size: 18px; margin-right: 8px; vertical-align: middle;
}
.warning-box ul {
  margin: 8px 0 4px 20px; padding: 0;
}
.warning-box li {
  margin-bottom: 6px;
}
.warning-box a { color: #4361ee; text-decoration: none; font-weight: 600; }
"""


def build_html(items, date_from, date_to):
    """Build HTML report grouped by topic."""

    # Group by topic
    topics = {}
    for item in items:
        topics.setdefault(item["topic"], []).append(item)

    sorted_topics = sorted(topics.keys())

    # Sort articles: published date descending, then title ascending
    for t in sorted_topics:
        topics[t].sort(key=lambda e: e.get("title", "").lower())
        topics[t].sort(key=lambda e: e.get("published", "") or "0000.00.00",
                       reverse=True)

    date_from_dot = date_from.replace("-", ".")
    date_to_dot = date_to.replace("-", ".")
    date_label_html = f"From: {date_from_dot} To: {date_to_dot}"

    h = []
    h.append('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">')
    h.append(f'<title>PescoPedia Recordings Digest - {date_label_html}</title>')
    h.append(f'<style>\n{CSS}</style>')
    h.append('</head><body>')
    h.append('<div class="wrapper">')

    # Header
    h.append('<h1>PescoPedia Recordings Digest</h1>')
    h.append(f'<p style="font-size:15px;color:#555;margin:-12px 0 20px 0;">'
             f'{date_label_html}</p>')

    # Stats
    h.append('<div class="stats-bar"><table class="stats-table"><tr>')
    h.append(f'<td>Recordings: <b>{len(items)}</b></td>')
    h.append(f'<td>Topics: <b>{len(sorted_topics)}</b></td>')
    h.append('</tr></table></div>')

    # Table of contents
    h.append('<div class="toc-box" id="toc"><h2>Topics</h2><ul>')
    for t in sorted_topics:
        tid = re.sub(r'[^a-zA-Z0-9]', '-', t).lower()
        h.append(f'<li><a href="#{tid}">{html_mod.escape(t)}</a>'
                 f'<span class="count"> ({len(topics[t])})</span></li>')
    h.append('</ul></div>')

    # Info box
    h.append('<div class="info-box">')
    h.append('<span class="box-icon">\u2139\uFE0F</span> ')
    h.append('The meetings listed in this report are ordered by the video '
             'publication date on the streaming platform. This date does not '
             'necessarily correspond to the actual date on which the meeting '
             'took place.')
    h.append('</div>')

    # Warning box
    h.append('<div class="warning-box">')
    h.append('<span class="box-icon">\u26A0\uFE0F</span> ')
    h.append('Please note that clicking the links in the meeting titles may '
             'result in an access denied message if you did not attend the '
             'meeting. In such cases, two options are recommended to gain access:')
    h.append('<ul>')
    h.append('<li>If the meeting belongs to a community that copies recordings '
             'to a shared and broadly accessible location, identify the relevant '
             'repository and search for the meeting there. You can also search '
             'the repository in <a href="https://microsofteur-my.sharepoint.com'
             '/:l:/g/personal/stefanpe_microsoft_com/JACjtdOsxzGsQpQbi9QQgRmqA'
             'T7I5RPNlgsRwwB9iTqZGRk">PescoPedia / Resources</a>.</li>')
    h.append('<li>Otherwise, try requesting access directly from the access '
             'denied page.</li>')
    h.append('</ul></div>')

    # Separator
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
            duration = item["duration"]

            h.append('<div class="article">')

            # Title (clickable)
            if url:
                url_esc = html_mod.escape(url)
                h.append(f'<p class="article-title">'
                         f'<a href="{url_esc}" target="_blank">'
                         f'{title_esc}</a></p>')
            else:
                h.append(f'<p class="article-title">{title_esc}</p>')

            # Meta: date + tech tags + duration
            meta_parts = []
            if pub:
                meta_parts.append(
                    f'<span class="date">{html_mod.escape(pub)}</span>')
            for tech in techs:
                meta_parts.append(
                    f'<span class="tag">{html_mod.escape(tech)}</span>')
            if duration:
                meta_parts.append(
                    f'<span class="duration-tag">'
                    f'{html_mod.escape(duration)}</span>')
            if meta_parts:
                h.append(f'<p class="article-meta">{" ".join(meta_parts)}</p>')

            # Summary (may contain HTML formatting from pipeline)
            if summary:
                h.append(f'<div class="article-summary">{summary}</div>')

            h.append('</div>')

        h.append('<a href="#toc" class="back-link">&uarr; Back to index</a>')
        h.append('</div>')

    h.append('<div class="footer-bar">'
             'Generated by Recordings Digest Pipeline</div>')
    h.append('</div>')  # wrapper
    h.append('</body></html>')

    return '\n'.join(h)


def main():
    args = parse_args()
    date_from = args.from_date
    date_to = args.to_date
    recipients = (args.recipients if args.recipients
                  else CONFIG.get("email_report", {}).get("default_recipients", ""))

    date_from_dot = date_from.replace("-", ".")
    date_to_dot = date_to.replace("-", ".")

    print(f"Fetching SP VideosMSInt items from {date_from_dot} to {date_to_dot}...",
          file=sys.stderr)

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        sp_page = ctx.new_page()
        sp_page.goto(SP_LIST_URL, wait_until="domcontentloaded", timeout=30000)
        sp_page.wait_for_timeout(5000)

        items = fetch_items(sp_page, date_from_dot, date_to_dot)
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
            "message": "No recording items found for the specified date range."
        }, indent=2))
        return

    html_content = build_html(items, date_from, date_to)

    # Save to output directory
    os.makedirs(os.path.join(BASE, "output"), exist_ok=True)
    html_filename = (
        f"Recordings_Digest-From-{date_from_dot}-To-{date_to_dot}.html"
    )
    html_path = os.path.join(BASE, "output", html_filename)
    if os.path.exists(html_path):
        for seq in range(2, 100):
            html_filename = (
                f"Recordings_Digest-From-{date_from_dot}-To-{date_to_dot}"
                f"-{seq:02d}.html"
            )
            html_path = os.path.join(BASE, "output", html_filename)
            if not os.path.exists(html_path):
                break
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Build subject line
    subject = (
        f"PescoPedia Recordings Digest - From: {date_from_dot} To: {date_to_dot}"
    )

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
