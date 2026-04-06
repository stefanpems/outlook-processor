---
name: combined-digest
description: "Run both the blog email digest and the Viva Engage conversations digest in sequence for a given time window. Triggered by generic digest requests like 'crea un digest degli ultimi 3 giorni', 'create a digest for the last 5 days', 'digest settimanale', WITHOUT specifying blog or Viva Engage explicitly."
argument-hint: "Specify the number of days, e.g.: 'crea un digest degli ultimi 3 giorni' or 'digest for the last 7 days'"
---

# Combined Digest (Blog + Viva Engage)

## Purpose

When the user asks to "create a digest" for the last X days **without specifying** which content (blog or Viva Engage), this skill orchestrates both digest pipelines in sequence:

1. **Blog email digest** — queries SP BlogPosts for the date range and sends the HTML digest email
2. **Viva Engage conversations digest** — reads communities, summarizes conversations, and sends the HTML digest email

## When to Use This Skill

Use this skill when the user's prompt matches **all** of these conditions:

- Asks to **create / send a digest** (e.g. "crea un digest", "create a digest", "digest settimanale", "manda il digest degli ultimi 5 giorni")
- Specifies a **time window** (e.g. "ultimi 3 giorni", "last 7 days", "della settimana")
- Does **NOT** specify a particular content type (does not mention "blog", "Viva Engage", "community", "conversazioni", "email di notifica")

If the user explicitly mentions "blog" or "Viva Engage" / "community", use the corresponding individual skill instead.

## Date Range Calculation

Given "last X days":

- **End date (to):** yesterday (today minus 1 day)
- **Start date (from):** yesterday minus (X − 1) days

This makes the range X days inclusive ending yesterday.

Examples (assuming today is 2026-04-06):
- "ultimi 3 giorni" → from 2026-04-03 to 2026-04-05
- "ultimo giorno" / "ieri" → from 2026-04-05 to 2026-04-05
- "ultima settimana" → from 2026-03-30 to 2026-04-05

## Procedure

### Step 1 — Parse Parameters and Read Config

Extract the number of days from the prompt. Compute `date_from` and `date_to` as YYYY-MM-DD.

Read `config.json` to get:
- `email_report.default_recipients` (blog digest recipients)
- `viva_engage.default_recipients` (Viva Engage digest recipients)
- `viva_engage.communities` (list of communities)

### Step 2 — Blog Email Digest

Execute the **blog-notifications** skill in **Digest mode (standard)** for the computed date range. This ensures all blog notification emails in the period are registered in SharePoint (creating or completing SP items as needed) before generating and sending the HTML digest.

Follow the full blog-notifications Digest mode procedure (Steps 0–5). The digest email is sent automatically as part of Step 5.

If no blog items exist in SP for the date range after processing, inform the user and proceed to Step 3 (do not send an empty blog digest).

### Step 3 — Viva Engage Conversations Digest

Follow the **vivaengage-conversations** skill procedure (Mode 2 — email digest):

1. For each community in `viva_engage.communities`, run:
   ```bash
   python engage_read_conversations.py "<community_name>" <days>
   ```
2. Summarize conversations (LLM task per the vivaengage-conversations skill rules)
3. Build the HTML report:
   ```bash
   python engage_build_html.py --input _ve_summaries.json
   ```
4. Send the email via `send_email` MCP tool to `viva_engage.default_recipients`

If no conversations are found across all communities, inform the user (do not send an empty digest).

### Step 4 — Final Summary

Report to the user:
- **Blog digest:** number of articles, topics, date range, recipients (or "skipped — no items")
- **Viva Engage digest:** number of conversations, communities covered, recipients (or "skipped — no conversations")
