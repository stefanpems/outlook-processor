"""
Build an HTML report from Viva Engage conversation summaries.

Usage:
  python engage_build_html.py < summaries.json
  python engage_build_html.py --input summaries.json

Input JSON format:
{
  "date_label": "2026-04-06",
  "days": 1,
  "communities": [
    {
      "community": "Community Name",
      "conversations": [
        {
          "type": "question",
          "title": "Short title",
          "thread_url": "https://engage.cloud.microsoft/...",
          "has_images": false,
          "author": "Name",
          "date": "2026-04-05",
          "summary_lines": [
            "<b>Question:</b> ...",
            "<b>Author:</b> ...",
            "<b>Answer:</b> ..."
          ]
        }
      ]
    }
  ]
}

Output: JSON to stdout with html_path.
"""
import json, re, os, sys, argparse, html as html_mod
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

OUTPUT_DIR = os.path.join(BASE, CONFIG.get("output", {}).get("dir", "output"))
COLOR_PALETTE = CONFIG.get("topic_color_palette", [
    "#F0E6D3", "#D3E8F0", "#D3F0D6", "#F0D3E6", "#E6F0D3", "#D3D8F0",
    "#F0DAD3", "#D3F0EA", "#E8D3F0", "#F0F0D3", "#D3EAF0", "#E6D3F0",
    "#F0D3D3", "#D3F0D3", "#D3D3F0", "#F0ECD3", "#F0D3EC", "#D3F0F0",
])

COMMUNITY_COLORS = {}


def get_community_color(name):
    if name not in COMMUNITY_COLORS:
        idx = len(COMMUNITY_COLORS) % len(COLOR_PALETTE)
        COMMUNITY_COLORS[name] = COLOR_PALETTE[idx]
    return COMMUNITY_COLORS[name]


TYPE_LABELS = {
    "question": ("Question", "#e74c3c"),
    "announcement": ("Announcement", "#2ecc71"),
    "discussion": ("Discussion", "#9b59b6"),
}

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
.community-section {
  background-color: #ffffff; border: 1px solid #e0e2e8;
  border-radius: 6px; padding: 16px 20px; margin-bottom: 20px;
}
.community-header {
  font-size: 18px; color: #1a1a2e; font-weight: 700;
  margin: 0 0 12px 0;
}
.badge {
  background-color: #4361ee; color: #ffffff; font-size: 12px;
  padding: 2px 10px; border-radius: 10px;
  margin-left: 8px; font-weight: 400;
}
.conversation {
  border-bottom: 1px solid #eee; padding: 12px 0;
}
.conversation:last-child { border-bottom: none; }
.conv-title {
  font-size: 15px; font-weight: 600; margin: 0;
}
.conv-title a { color: #4361ee; text-decoration: none; }
.type-badge {
  font-size: 11px; padding: 2px 8px; border-radius: 10px;
  color: #ffffff; margin-right: 8px; font-weight: 600;
  display: inline-block; vertical-align: middle;
}
.conv-body {
  font-size: 14px; line-height: 1.55; color: #333;
  margin-top: 6px;
}
.conv-body b { color: #1a1a2e; }
.images-tag {
  font-size: 12px; color: #888; margin-left: 8px;
}
.back-link {
  display: inline-block; margin-top: 10px; color: #4361ee;
  font-size: 13px; text-decoration: none; font-weight: 600;
}
.footer-bar {
  text-align: center; color: #999; font-size: 12px;
  margin-top: 28px; padding-top: 12px;
  border-top: 1px solid #e0e0e0;
}
.no-conversations {
  font-size: 14px; color: #888; font-style: italic;
  padding: 8px 0;
}
"""


def build_html(data):
    communities = data.get("communities", [])
    date_label = data.get("date_label", datetime.now().strftime("%Y-%m-%d"))
    days = data.get("days", 1)

    total_convs = sum(len(c.get("conversations", [])) for c in communities)

    h = []
    h.append('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">')
    date_from = data.get("date_from", "")
    date_to = data.get("date_to", "")
    if date_from and date_to:
        date_from_fmt = date_from.replace("-", ".")
        date_to_fmt = date_to.replace("-", ".")
        title_text = f"PescoPedia Viva Engage Conversational Digest - From: {date_from_fmt} To: {date_to_fmt}"
        subtitle = f"From: {date_from_fmt} To: {date_to_fmt}"
    else:
        title_text = "PescoPedia Viva Engage Conversational Digest"
        subtitle = f"Last {days} day{'s' if days != 1 else ''} (as of {date_label})"

    h.append(f'<title>{html_mod.escape(title_text)}</title>')
    h.append(f'<style>\n{CSS}</style>')
    h.append('</head><body>')
    h.append('<div class="wrapper">')

    # Header
    h.append('<h1>PescoPedia Viva Engage Conversational Digest</h1>')
    h.append(f'<p style="font-size:15px;color:#555;margin:-12px 0 20px 0;">'
             f'{html_mod.escape(subtitle)}</p>')

    # Stats bar
    h.append('<div class="stats-bar"><table class="stats-table"><tr>')
    h.append(f'<td>Communities: <b>{len(communities)}</b></td>')
    h.append(f'<td>Conversations: <b>{total_convs}</b></td>')
    h.append(f'<td>Period: <b>{days} day{"s" if days != 1 else ""}</b></td>')
    h.append('</tr></table></div>')

    # Table of contents
    h.append('<div class="toc-box" id="toc"><h2>Communities</h2><ul>')
    for comm in communities:
        name = comm.get("community", "Unknown")
        cid = re.sub(r'[^a-zA-Z0-9]', '-', name).lower()
        count = len(comm.get("conversations", []))
        h.append(f'<li><a href="#{cid}">{html_mod.escape(name)}</a>'
                 f'<span class="count"> ({count})</span></li>')
    h.append('</ul></div>')

    # Separator
    h.append('<hr style="border:none;border-top:2px solid #e0e2e8;margin:28px 0;">')

    # Community sections
    for comm in communities:
        name = comm.get("community", "Unknown")
        cid = re.sub(r'[^a-zA-Z0-9]', '-', name).lower()
        convs = comm.get("conversations", [])
        color = get_community_color(name)

        h.append(f'<div class="community-section" id="{cid}" '
                 f'style="border-left: 4px solid {color};">')
        h.append(f'<h2 class="community-header">{html_mod.escape(name)}'
                 f'<span class="badge">{len(convs)}</span></h2>')

        if not convs:
            h.append('<p class="no-conversations">No recent conversations.</p>')
        else:
            for i, conv in enumerate(convs, 1):
                conv_type = conv.get("type", "discussion")
                title = conv.get("title", "Untitled")
                thread_url = conv.get("thread_url", "")
                has_images = conv.get("has_images", False)
                summary_lines = conv.get("summary_lines", [])

                label, type_color = TYPE_LABELS.get(conv_type,
                                                     ("Discussion", "#3498db"))

                h.append('<div class="conversation">')

                # Title line with type badge
                title_esc = html_mod.escape(title)
                type_badge = (f'<span class="type-badge" '
                              f'style="background-color:{type_color};">'
                              f'{label}</span>')

                if thread_url:
                    url_esc = html_mod.escape(thread_url)
                    h.append(f'<p class="conv-title">{type_badge}'
                             f'<a href="{url_esc}" target="_blank">'
                             f'{title_esc}</a>')
                else:
                    h.append(f'<p class="conv-title">{type_badge}{title_esc}')

                if has_images:
                    h.append('<span class="images-tag">&#128206; Images</span>')
                h.append('</p>')

                # Summary body
                if summary_lines:
                    h.append('<div class="conv-body">')
                    for line in summary_lines:
                        h.append(f'<p style="margin:3px 0;">{line}</p>')
                    h.append('</div>')

                h.append('</div>')  # conversation

        h.append('<a href="#toc" class="back-link">&uarr; Back to index</a>')
        h.append('</div>')  # community-section

    h.append('<div class="footer-bar">Generated by Viva Engage Digest Pipeline</div>')
    h.append('</div>')  # wrapper
    h.append('</body></html>')

    return '\n'.join(h)


def parse_args():
    parser = argparse.ArgumentParser(description="Build Viva Engage digest HTML.")
    parser.add_argument("--input", default="",
                        help="Path to JSON input file (default: stdin)")
    return parser.parse_args()


def main():
    args = parse_args()

    if args.input:
        with open(args.input, encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = json.load(sys.stdin)

    html_content = build_html(data)

    # Write to output directory
    date_from = data.get("date_from", "")
    date_to = data.get("date_to", "")
    if date_from and date_to:
        from_dot = date_from.replace("-", ".")
        to_dot = date_to.replace("-", ".")
    else:
        date_label = data.get("date_label", datetime.now().strftime("%Y-%m-%d"))
        from_dot = date_label.replace("-", ".")
        to_dot = from_dot
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"Viva_Engage-Digest-From-{from_dot}-To-{to_dot}.html"
    html_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(html_path):
        for seq in range(2, 100):
            filename = f"Viva_Engage-Digest-From-{from_dot}-To-{to_dot}-{seq:02d}.html"
            html_path = os.path.join(OUTPUT_DIR, filename)
            if not os.path.exists(html_path):
                break
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    result = {"html_path": html_path, "filename": filename}
    print(json.dumps(result), flush=True)


if __name__ == "__main__":
    main()
