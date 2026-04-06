---
name: blog-notifications
description: "Process blog notification emails from Outlook: fetch content, generate summaries, create SP items, categorize and move emails. Triggered by prompts like 'processa le notifiche di blog ricevute via email dal giorno X al giorno Y' or 'process blog notifications from date X to date Y'."
argument-hint: "Specify start and end dates, e.g.: 'dal 2026.03.01 al 2026.03.15' or 'from 2026-03-01 to 2026-03-15'"
---

# Blog Notification Processing

## Purpose

Process blog post notification emails from Outlook. For each email: fetch and summarize the linked blog content, deduplicate within the session and against SharePoint, create a new SP list item, assign an Outlook category, move the email to a target folder, and track everything in synchronized XLSX + HTML reports — all updated after each individual email.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections:

- `outlook.sender` — sender to filter emails by
- `outlook.subject_prefix` — subject prefix to filter (e.g. `[Blog-`)
- `outlook.processed_category` — Outlook category assigned after processing
- `outlook.target_folder` — Outlook folder to move processed emails into
- `outlook.exclude_terms` — list of terms always added as negative filters to every Outlook search (e.g. `["PescoPedia"]` → appends `-PescoPedia` to every query)
- `sharepoint.blog_list_api` — SP REST API base for the BlogPosts list
- `sharepoint.blog_list_entity_type` — entity type for POST/MERGE
- `sharepoint.site_base` — SP site base URL for REST calls
- `source_map` — topic name → SP SourceNew lookup ID
- `tech_map` — technology label → SP Tech lookup ID
- `topic_color_palette` — pastel colors for topic-based styling

## Pipeline Python Scripts

The workspace contains purpose-built Python scripts for each pipeline activity. **Use these scripts. Do NOT create new scripts that duplicate their functionality.** All scripts read `config.json` at startup.

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `pipeline_init.py` | Initialize session: create XLSX + HTML templates with progressive naming | `python pipeline_init.py` |
| `pipeline_retrieve.py` | Retrieve matching emails from Outlook Web via CDP | `python pipeline_retrieve.py YYYY-MM-DD YYYY-MM-DD [--include-processed]` |
| `pipeline_fetch_blog.py` | Fetch blog content, resolve final URL, extract metadata | `python pipeline_fetch_blog.py <url>` |
| `pipeline_check_dup.py` | Check duplicates (session + SP) | `python pipeline_check_dup.py <title> <final_url>` |
| `pipeline_sp_create.py` | Create new SP BlogPosts item via REST API | `echo '{...}' \| python pipeline_sp_create.py -` |
| `pipeline_email_actions.py` | Categorize and/or move email in Outlook | `python pipeline_email_actions.py both <title>` |
| `pipeline_update_reports.py` | Rebuild XLSX + HTML from session_state.json | `python pipeline_update_reports.py` |
| `pipeline_fetch_blogposts.py` | Fetch all existing SP BlogPosts → `sp_blogposts.json` | `python pipeline_fetch_blogposts.py` |
| `pipeline_sweep_inbox.py` | Sweep all unprocessed blog emails from Inbox: categorize + move, loop until zero remain | `python pipeline_sweep_inbox.py YYYY-MM-DD YYYY-MM-DD` |

### Supporting Scripts

| File | Purpose |
|------|---------|
| `pipeline_cache_blogs.py` | Bulk blog content fetcher/cacher. |
| `pipeline_update_sp_summaries.py` | Bulk-update Summary field on existing SP items. |

## Input

The user provides a date range. Accept any format — `YYYY.MM.DD`, `YYYY-MM-DD`, `DD/MM/YYYY`, natural language. Parse both dates into `YYYY-MM-DD` format and confirm before starting.

The user may also explicitly request to **reprocess already-processed emails**. This is referred to as **reprocess mode** throughout this document.

**Reprocess mode is activated when ANY of these conditions are met:**
- The Italian verb has a **"Ri-" prefix**: "Riprocessa", "Rielabora", "Rivaluta", "Ricontrolla", "Rifai", "Rianalizza", etc.
- The English verb has a **"Re-" prefix**: "Reprocess", "Re-evaluate", "Redo", "Rerun", "Regenerate", etc.
- Explicit phrases: "anche le già processate", "include already processed", "rigenera i summary".

In short: if the user's action verb starts with "Ri" (Italian) or "Re" (English) prefix implying repetition, treat it as reprocess mode. If the verb has no repetition prefix (e.g. "Processa", "Elabora", "Process"), use normal mode.

### Digest Mode

**Digest mode is activated** when the user asks to **create a digest** (e.g. "Crea digest dei blog", "Create blog digest", "Genera il digest dei blog post dal X al Y"). The key verbs are "crea/create/genera" (not "invia/send" — which triggers the `blog-email-report` skill for SP-only queries without email processing).

Digest mode has two sub-cases:

**Digest + Reprocess:** If the digest request ALSO contains a reprocess indicator ("ri-"/"re-" prefix, such as "Crea digest riprocessando", "Regenera il digest"), combine reprocess logic with mandatory email sending. Process all emails (including already-categorized), regenerate all summaries, then send the digest.

**Digest (standard):** If no reprocess indicator is present:
- Retrieve ALL emails in the period, **including** already-processed ones (same retrieval flag as reprocess mode).
- For each email, check if a complete SP item already exists. If complete, just ensure the email is categorized+moved, and collect the SP item data for the digest. If incomplete or missing, fetch content, create/update the SP item, then collect data.
- After processing all emails, **always** generate and send the HTML digest email.

**Required SP fields (completeness check for Blogs):** Published, Title, Tech, Link, Source, Summary. An SP item is considered "complete" when `sp_has_summary` is true (the pipeline always populates all fields at creation time, so Summary presence reliably indicates full field completeness).

### Summary of Execution Modes

| Prompt pattern | Mode | Email retrieval | SP logic | Digest email |
|----------------|------|-----------------|----------|---------------|
| "Processa le email..." | Standard | Exclude processed | Create if missing, update Summary if empty, skip if complete | Only if explicitly requested |
| "Riprocessa le email..." | Reprocess | Include all | Create if missing, always regenerate+update | Only if explicitly requested |
| "Crea digest..." | Digest | Include all | Create if missing, fill gaps if incomplete, skip if complete | **Always** sent |
| "Crea digest... riprocessa..." | Digest + Reprocess | Include all | Create if missing, always regenerate+update | **Always** sent |

If dates are missing, ask the user.

## Procedure

Execute the steps below **strictly in order**. For each step, use the specified Python script via `run_in_terminal`.

---

### Step 0 — Initialize Session

```bash
python pipeline_init.py
```

This creates:
- An XLSX file: `output/YYYY.MM.DD-NN-ProcessedEmails.xlsx` with header row
- An HTML file: `output/YYYY.MM.DD-NN-ProcessedEmails.html` with template
- A session state file: `session_state.json`

The script outputs JSON with the paths. Note the `xlsx_path` and `html_path`.

---

### Step 1 — Fetch Existing SP Items (for deduplication)

```bash
python pipeline_fetch_blogposts.py
```

This downloads all existing SP BlogPosts items to `sp_blogposts.json`. Needed for Step 3.3 (SP dedup check).

---

### Step 2 — Retrieve Emails from Outlook Web

**Normal mode (default):**

```bash
python pipeline_retrieve.py {DATE_FROM} {DATE_TO}
```

**Reprocess mode** (only if user explicitly requested it):

**Digest mode** (both sub-cases — always includes processed emails):

```bash
python pipeline_retrieve.py {DATE_FROM} {DATE_TO} --include-processed
```

This script:
1. Connects to Edge via CDP
2. Navigates to Outlook Web
3. Searches: `from:{sender} subject:{prefix} received:{date_from}..{date_to} -"{processed_category}" -{exclude_term_1} -{exclude_term_2} ...` (in normal mode the negative filter excludes already-processed emails; in reprocess mode the category filter is omitted but `exclude_terms` are always applied)
4. Scrolls the virtualized list to find ALL results
5. Extracts subject, date, blog URL from each email's reading pane
6. **In normal mode:** skips emails that already have the processed category (visual check). **In reprocess mode:** includes all emails regardless of category.
7. **Deduplicates by email identity** (subject + date + link) to remove duplicates from Outlook's "Risultati principali" section, which mirrors emails shown in later sections
8. Saves all remaining emails to `session_state.json`. Distinct emails with the same subject (different dates/links) are kept — duplicates are handled by session/SP dup checks in Step 3.

**Output:** JSON with total count.

**CRITICAL — If 0 emails are found (normal mode):** Stop the entire pipeline and inform the user. Do **NOT** retry without the negative filter. Do **NOT** fall back to reprocess mode. The only way to include already-processed emails is if the user **expressly** requested it in the original prompt.

**If 0 emails are found (Digest mode):** Skip the per-email processing loop (Step 3) but **still execute Step 5** — the digest report is generated from SP data and may contain items registered in previous sessions.

---

### Step 3 — Process Each Email (Per-Email Loop)

Loop over the emails in `session_state.json`. For EACH email, perform Steps 3.1 through 3.7 in sequence. After each email, update the session state and reports.

**Digest mode (standard) — modified step order:** Perform Step 3.2 (dup check) **before** Step 3.1 (content fetch). Use the email's original `blog_link` URL and the `title` extracted from the email subject for the dup check. This allows skipping the expensive content fetch for SP items that are already complete. See the Digest mode table in Step 3.2 for the decision logic.

**Digest + Reprocess mode:** Follow the same step order as Reprocess mode (3.1 → 3.2 → 3.4 → 3.5 → 3.6 → 3.7).

#### 3.1 — Fetch Blog Content & Resolve Final URL

```bash
python pipeline_fetch_blog.py "<blog_url>"
```

This returns JSON with: `final_url`, `title`, `published_date`, `content`, `content_length`.

- The `final_url` is the canonical URL after redirects.
- The `published_date` is extracted from page metadata or URL.
- The `content` is the full article text (up to 8000 chars).

If the script fails (e.g. 403), use `fetch_webpage` as fallback for public pages.

**Important:** If `published_date` is empty, fall back to the email's received date.

#### 3.2 — Check for Session Duplicate

```bash
python pipeline_check_dup.py "<title>" "<final_url>"
```

This returns JSON with: `dup_session`, `dup_sp`, `sp_id`, `sp_has_summary`.

**Session duplicate** (`dup_session` true): another email in this session already had the same title+URL. Skip Steps 3.4-3.5 entirely and go straight to 3.6 (categorize+move).

**SP item logic — Normal mode** (based on `dup_sp` and `sp_has_summary`):

| `dup_sp` | `sp_has_summary` | Action |
|----------|-----------------|--------|
| false | — | Generate summary (3.4) + Create new SP item (3.5) |
| true | false | Generate summary (3.4) + Update existing SP item summary (3.5) |
| true | true | **Skip** summary and SP entirely — go to 3.6 |

**SP item logic — Reprocess mode** (based on `dup_sp`):

| `dup_sp` | Action |
|----------|--------|
| false | Generate summary (3.4) + Create new SP item (3.5) |
| true | **Always** regenerate summary (3.4) + Update/overwrite existing SP item summary (3.5), regardless of `sp_has_summary` |

**SP item logic — Digest mode (standard)** (based on `dup_sp` and `sp_has_summary`):

Remember: in this mode, Step 3.2 runs **before** Step 3.1.

| `dup_sp` | `sp_has_summary` | Action |
|----------|-----------------|--------|
| false | — | Fetch blog content (3.1) + Generate summary (3.4) + Create new SP item (3.5) |
| true | false | Fetch blog content (3.1) + Generate summary (3.4) + Update existing SP item (3.5) |
| true | true | **Skip** content fetch (3.1), summary (3.4), and SP update (3.5) — item is complete |

In all three cases, the SP item data (title, published_date, tech, link, source/topic, summary) will be included in the digest report generated in Step 5. For pre-existing complete items, `pipeline_email_report.py` reads these fields directly from SP.

**Digest + Reprocess mode:** Same table as Reprocess mode above (always regenerate, always update).

**Always perform** categorize + move (Step 3.6) for every email, regardless of duplicates.

#### 3.4 — Generate Summary (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Generate a summary only when needed (see table above). Read the `content` returned by Step 3.1:

- **Language:** Always English, regardless of source language.
- **Length:** 100–150 words.
- **Format:** Use `<b>` for keywords, `<i>` for key phrases, `<ul><li>` for bullet points when useful.
- **Coverage:** The summary must mention **every** concept, feature, or technique explained in the article's core content. Nothing should be silently omitted — at minimum, each concept must be cited by name even if not elaborated.
- **Skip preambles:** Most articles open with generic introductory paragraphs (context-setting, motivation, broad statements). Ignore these — focus exclusively on the substantive content that is the real subject of the article.
- **Style:** Concrete, informative. Report only what the author communicates. No filler.

Store the summary in the email's `summary` field in session state.

#### 3.5 — Classify Technology & SP Item (Create or Update)

**Technology classification** is also a Copilot LLM task. Using `tech_map` from `config.json`, assign technology tags:
- If a sub-technology matches (e.g. `Azure / AKV / MHSM`), do NOT also include the parent.
- Comma-separate multiple values.
- Only use values from the map. If no match, leave empty.

Store the tech classification in the email's `tech` field.

**If creating a new SP item** (`dup_sp` was false):

```bash
echo '{"title":"...","published_date":"...","summary":"...","topic":"...","tech":"...","blog_link":"..."}' | python pipeline_sp_create.py -
```

**If updating summary on existing SP item** (`dup_sp` true, `sp_has_summary` false):

```bash
echo '{"summary":"...","title":"..."}' | python pipeline_sp_create.py --update-summary <sp_id> -
```

The JSON must contain at minimum: `summary`. Do NOT create temporary files.

#### 3.6 — Categorize + Move Email in Outlook

```bash
python pipeline_email_actions.py both "<email_title>"
```

This script:
1. Searches Outlook for the specific email (sender + `[Blog-` + first 6 words)
2. Selects ONLY rows containing `[Blog-` (NEVER Ctrl+A)
3. Right-clicks → Categorizza → selects the configured category
4. Clicks Sposta (top toolbar, y < 200) → searches folder → ArrowDown + Enter

If successful, mark `categorized = "Yes"` and `moved = "Yes"` in session state.

#### 3.7 — Update Reports & Session State

After each email (whether duplicate or new):

1. **Register the title in session's processed_titles** (for future dedup checks within session)
2. **Update `session_state.json`** with all status fields for this email
3. **Run report update:**

```bash
python pipeline_update_reports.py
```

This rebuilds both XLSX and HTML from the current session state.

---

### Step 4 — Final Summary

After all emails are processed, print a summary:
- Total emails retrieved
- Session duplicates found
- SP duplicates found
- SP items created
- Emails categorized
- Emails moved
- Any errors encountered

---

### Step 5 — Send Email Report

**When this step is executed:**
- **Standard / Reprocess mode:** Only if the user explicitly requested sending the report (e.g. "e invia il report", "and send the report by email").
- **Digest mode (any sub-case):** **Always** executed — sending the digest is the primary goal of Digest mode.

Use the **same HTML digest template and sending logic** as the `blog-email-report` skill.

#### 5.1 — Generate the HTML Digest

Run `pipeline_email_report.py` with the same date range used for email retrieval:

```bash
python pipeline_email_report.py --from-date {DATE_FROM} --to-date {DATE_TO}
```

Parse the JSON output. If `total_items` is 0, inform the user and skip sending.

#### 5.2 — Send the Email

Read the HTML file at `html_path` returned in Step 5.1.

Use the **send_email** MCP tool with:
- `emailAddresses`: recipients from `config.json` → `email_report.default_recipients` (or as specified by the user)
- `subject`: the `subject` value from the script output
- `htmlBody`: the full HTML content read from the file

#### 5.3 — Confirm to User

Report:
- Number of articles included in the digest
- Number of topics
- Date range
- Recipients the email was sent to

---

## Session State Schema

The `session_state.json` file tracks the entire session. Each email in the `emails` array has:

```json
{
  "subject": "[Blog-Topic] Article Title",
  "received_date": "2026-04-04T14:32:00",
  "blog_link": "https://original-url.com/...",
  "title": "Article Title",
  "topic": "Topic",
  "final_url": "https://final-url.com/...",
  "published_date": "2026-04-04",
  "content": "Full article text...",
  "tech": "Technology Tag",
  "summary": "<b>Summary</b> with HTML formatting...",
  "dup_session": "",
  "dup_sp": "",
  "sp_created": "",
  "categorized": "",
  "moved": ""
}
```

The `processed_titles` dict maps `title → final_url` for session dedup.

## Report Formats

### XLSX Report

Columns: Published, Title, Tech, Link, Topic, Summary, Formatted Summary, Dup Session, Dup SP, SP Created, Categorized, Moved.

- Header: dark background (#2D3748), white bold text
- Rows colored by topic from `topic_color_palette`
- Status fields show "Yes" in bold green
- Sorted by topic (asc) then published_date (desc)

### HTML Report

- Professional, clean design on light background
- Stats bar at top with counts
- Table of contents linking to topic sections
- Each topic is a section with articles as sub-sections
- Each article shows: linked title, date, tech tags, rich-text summary
- "Back to index" link after each topic section
- Duplicate articles marked with a badge

---

## Lessons Learned (Mandatory Rules)

1. **Outlook search MUST ALWAYS filter by sender AND subject prefix.** Never run a broad search. Never read emails not returned by filtered search.

2. **NEVER use Ctrl+A to select search results.** Results may include unrelated emails. Always iterate rows, check for `[Blog-` text, Ctrl+click only verified matches.

3. **SP Link field cannot be set in POST.** `pipeline_sp_create.py` handles this with POST + MERGE.

4. **Outlook "Sposta" (Move) button exists in multiple places.** Always use the one in top toolbar (y < 200). Use keyboard navigation (ArrowDown + Enter).

5. **Refresh SP digest token every ~40 items** to avoid 403 errors.

6. **Long titles break Outlook search.** `pipeline_email_actions.py` shortens to first 6 words automatically.

7. **ElementHandle.click can timeout.** Scripts fall back to bounding-box clicks.

8. **Menu items use Italian labels:** "Categorizza", "Sposta", "Cerca una cartella".

9. **`openpyxl` cannot render HTML in cells.** Summary column is plain text; Formatted Summary keeps HTML tags.

10. **NEVER use `-category:` in Outlook Web search.** It silently returns 0 results. Instead, `pipeline_retrieve.py` visually checks each email's category in the reading pane.

11. **Outlook Web uses a virtualized list** — only ~8-10 items rendered. `pipeline_retrieve.py` scrolls with `mouse.wheel` and waits until 3 stable rounds.

12. **Use `page.evaluate("window.location.href = ...")` for SPA navigation** in Outlook. Direct `page.goto()` may cause ERR_ABORTED.

## Tools Required

- **Python 3.12+** with `playwright`, `openpyxl` packages
- **Microsoft Edge** launched with `--remote-debugging-port=9222` and authenticated debug profile
- **`fetch_webpage`** tool for reading public blog pages (fallback if Python fetch fails)
- **`run_in_terminal`** for executing pipeline Python scripts
