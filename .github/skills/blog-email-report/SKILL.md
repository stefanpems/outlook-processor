---
name: blog-email-report
description: "Build an HTML digest of blog posts from the SharePoint BlogPosts list grouped by topic, and send it via email. Triggered by prompts like 'invia il digest dei blog' or 'send blog digest email' or 'manda il report dei blog per email'."
argument-hint: "Optionally specify recipients, date range, and tech filter, e.g.: 'invia il digest dei blog dal 2026.03.01 al 2026.03.15 a user@example.com filtrato per Sentinel,Entra'"
---

# Blog Email Report (Digest)

## Purpose

Query the SharePoint BlogPosts list for items in a date range, build an HTML digest grouped by topic, and send it via email. Each topic section lists articles sorted by publication date (descending) then title (ascending), showing clickable title, date, technologies, and summary.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections used by this skill:

- `sharepoint.blog_list_api` — SP REST API base for the BlogPosts list
- `sharepoint.blog_list_url` — SP list URL (used to establish auth context)
- `edge_cdp.url` — CDP endpoint for Playwright connection
- `topic_color_palette` — pastel colors for topic-based styling

## Parameters

The user may provide the following. Parse them from the prompt (accept any date format):

| Parameter | Description | Default |
|-----------|-------------|---------|
| **Recipients** | Semicolon-separated email addresses | from `config.json` → `email_report.default_recipients` |
| **Date range** | Start and end dates for the `field_0` (published date) filter | Yesterday only (full day) |
| **Tech filter** | Comma-separated technology labels to filter items by | *(empty — include all items)* |

If dates are ambiguous or missing, compute defaults (yesterday) and confirm with the user.

## Pipeline Script

| Script | Purpose | CLI |
|--------|---------|-----|
| `pipeline_email_report.py` | Fetch SP items, build HTML digest, save to `output/` | `python pipeline_email_report.py --from-date YYYY-MM-DD --to-date YYYY-MM-DD [--tech "tech1,tech2"] [--recipients "a@x.com;b@x.com"]` |

**Output:** JSON to stdout with `html_path`, `subject`, `recipients`, `total_items`, `topics_count`, `date_from`, `date_to`.

If `total_items` is 0, inform the user that no items matched and **do not send the email**.

## Procedure

### Step 1 — Parse Parameters

Extract from the user's prompt:
- **Recipients**: list of email addresses. Default: value from `config.json` → `email_report.default_recipients`
- **Date range**: convert to `YYYY-MM-DD` format. Default: yesterday for both start and end.
- **Tech filter**: comma-separated technology names (must match labels in `tech_map` from `config.json`). Default: empty string.

### Step 2 — Generate the HTML Digest

```bash
python pipeline_email_report.py --from-date {DATE_FROM} --to-date {DATE_TO} --tech "{TECH_FILTER}"
```

Omit `--tech` if no tech filter was specified.

Parse the JSON output. If `total_items` is 0, inform the user and stop.

### Step 3 — Send the Email

Read the HTML file at `html_path` returned in Step 2.

Use the **send_email** MCP tool with:
- `emailAddresses`: the recipients (semicolon-separated)
- `subject`: the `subject` value from the script output
- `htmlBody`: the full HTML content read from the file

### Step 4 — Confirm to User

Report:
- Number of articles included
- Number of topics
- Date range
- Recipients the email was sent to
