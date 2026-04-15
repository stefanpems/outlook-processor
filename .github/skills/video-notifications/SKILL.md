---
name: video-notifications
description: "Process video notification emails from Outlook: fetch YouTube metadata, download transcripts, generate abstracts, create SP items, categorize and move emails. Triggered by prompts like 'processa le notifiche video ricevute via email dal giorno X al giorno Y' or 'process video notifications from date X to date Y'."
argument-hint: "Specify start and end dates, e.g.: 'dal 2026.03.01 al 2026.03.15' or 'from 2026-03-01 to 2026-03-15'"
---

# Video Notification Processing

## Purpose

Process video notification emails from Outlook. For each email: fetch video metadata from YouTube (title, duration, published date, description, chapters), download the transcript, generate an abstract, deduplicate within the session and against SharePoint, create a new SP VideoPosts item, assign an Outlook category, move the email to a target folder, and track everything in synchronized XLSX + HTML reports — all updated after each individual email.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections:

- `video_outlook.sender` — sender to filter emails by
- `video_outlook.subject_prefix` — subject prefix to filter (e.g. `[Video-`)
- `video_outlook.processed_category` — Outlook category assigned after processing (`By agent - Video`)
- `video_outlook.target_folder` — Outlook folder to move processed emails into
- `video_outlook.exclude_terms` — list of terms always added as negative filters (e.g. `["PescoPedia"]`)
- `video_sharepoint.list_api` — SP REST API base for the VideoPosts list
- `video_sharepoint.list_entity_type` — entity type for POST/MERGE
- `video_sharepoint.fields` — field internal name mapping (published, abstract, duration, yt_id)
- `sharepoint.site_base` — SP site base URL for REST calls
- `source_map` — topic name → SP SourceNew lookup ID
- `tech_map` — technology label → SP Tech lookup ID

## Pipeline Python Scripts

The workspace contains purpose-built Python scripts for each pipeline activity. **Use these scripts. Do NOT create new scripts that duplicate their functionality.** All scripts read `config.json` at startup.

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `pipeline_init.py` | Initialize session: create XLSX + HTML templates | `python pipeline_init.py --type video --from-date YYYY-MM-DD --to-date YYYY-MM-DD` |
| `pipeline_video_retrieve.py` | Retrieve matching video emails from Outlook Web via CDP | `python pipeline_video_retrieve.py YYYY-MM-DD YYYY-MM-DD [--include-processed]` |
| `pipeline_fetch_video.py` | Fetch YouTube video metadata (title, date, duration, description, chapters) | `python pipeline_fetch_video.py <url>` |
| `yt_transcript.py` | Download YouTube transcript to `yt_transcripts/yt_<VIDEO_ID>.txt` | `python yt_transcript.py <video_id_or_url> [--clean]` |
| `pipeline_video_check_dup.py` | Check duplicates by title and yt_id (session + SP) | `python pipeline_video_check_dup.py <title> [<yt_id>]` |
| `pipeline_video_sp_create.py` | Create/update VideoPosts SP item via REST API | `python pipeline_video_sp_create.py _sp_input.json` |
| `pipeline_video_email_actions.py` | Categorize and/or move video email in Outlook | `python pipeline_video_email_actions.py both <title>` |
| `pipeline_update_reports.py` | Rebuild XLSX + HTML from session_state.json | `python pipeline_update_reports.py` |
| `pipeline_fetch_videoposts.py` | Fetch all existing SP VideoPosts → `sp_videoposts.json` | `python pipeline_fetch_videoposts.py` |
| `pipeline_video_email_report.py` | Build HTML video digest from SP VideoPosts (grouped by topic) | `python pipeline_video_email_report.py --from-date YYYY-MM-DD --to-date YYYY-MM-DD` |

## Input

The user provides a date range. Accept any format — `YYYY.MM.DD`, `YYYY-MM-DD`, `DD/MM/YYYY`, natural language. Parse both dates into `YYYY-MM-DD` format and confirm before starting.

The user may also explicitly request to **reprocess already-processed emails**. This is referred to as **reprocess mode** throughout this document.

**Reprocess mode is activated when ANY of these conditions are met:**
- The Italian verb has a **"Ri-" prefix**: "Riprocessa", "Rielabora", "Rivaluta", etc.
- The English verb has a **"Re-" prefix**: "Reprocess", "Re-evaluate", "Redo", etc.
- Explicit phrases: "anche le già processate", "include already processed".

If dates are missing, ask the user.

### Digest Mode

**Digest mode is activated** when the user asks to **create a digest** (e.g. "Crea digest dei video", "Create video digest", "Genera il digest dei video dal X al Y"). The key verbs are "crea/create/genera" (not "invia/send" — which would trigger a report-only query without email processing).

Digest mode has two sub-cases:

**Digest + Reprocess:** If the digest request ALSO contains a reprocess indicator ("ri-"/"re-" prefix, such as "Crea digest riprocessando", "Regenera il digest"), combine reprocess logic with mandatory email sending. Process all emails (including already-categorized), regenerate all abstracts, then send the digest.

**Digest (standard):** If no reprocess indicator is present:
- Retrieve ALL emails in the period, **including** already-processed ones (same retrieval flag as reprocess mode).
- For each email, check if a complete SP item already exists. If complete, just ensure the email is categorized+moved, and collect the SP item data for the digest. If incomplete or missing, fetch video metadata, create/update the SP item, then collect data.
- After processing all emails, **always** generate and send the HTML digest email.

**Required SP fields (completeness check for Videos):** Published, Title, Tech, Link, Source Blog, Abstract, yt_ID. An SP item is considered "complete" when `sp_has_abstract` is true (the pipeline always populates all fields at creation time, so Abstract presence reliably indicates full field completeness).

### Summary of Execution Modes

| Prompt pattern | Mode | Email retrieval | SP logic | Digest email |
|----------------|------|-----------------|----------|---------------|
| "Processa le email video..." | Standard | Exclude processed | Create if missing, update Abstract if empty, skip if complete | Only if explicitly requested |
| "Riprocessa le email video..." | Reprocess | Include all | Create if missing, always regenerate+update | Only if explicitly requested |
| "Crea digest dei video..." | Digest | Include all | Create if missing, fill gaps if incomplete, skip if complete | **Always** sent |
| "Crea digest dei video... riprocessa..." | Digest + Reprocess | Include all | Create if missing, always regenerate+update | **Always** sent |

## Procedure

Execute the steps below **strictly in order**. For each step, use the specified Python script via `run_in_terminal`.

---

### Step 0 — Initialize Session

```bash
python pipeline_init.py --type video --from-date {date_from} --to-date {date_to}
```

This creates:
- An XLSX file: `output/Video_Notifications-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.xlsx` with header row
- An HTML file: `output/Video_Notifications-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.html` with template
- A session state file: `session_state.json`

The script outputs JSON with the paths. Note the `xlsx_path` and `html_path`.

---

### Step 1 — Fetch Existing SP VideoPosts (for deduplication)

```bash
python pipeline_fetch_videoposts.py
```

This downloads all existing SP VideoPosts items to `sp_videoposts.json`. Needed for Step 3.3 (SP dedup check).

---

### Step 2 — Retrieve Emails from Outlook Web

**Normal mode (default):**

```bash
python pipeline_video_retrieve.py {DATE_FROM} {DATE_TO}
```

**Reprocess mode** (only if user explicitly requested it):

**Digest mode** (both sub-cases — always includes processed emails):

```bash
python pipeline_video_retrieve.py {DATE_FROM} {DATE_TO} --include-processed
```

This script:
1. Connects to Edge via CDP
2. Navigates to Outlook Web
3. Builds the search query from `config.json` values:
   - **Normal mode:** `from:{sender} subject:[Video- received:{date_from}..{date_to} -"{processed_category}" -{exclude_term_1} -{exclude_term_2} ...` — the negative filter on the processed category (`"By agent - Video"`) **excludes** already-processed emails; `exclude_terms` (e.g. `PescoPedia`) are **always** applied.
   - **Reprocess mode:** Same query but the `-"{processed_category}"` filter is **omitted** (so already-processed emails are included). The `exclude_terms` negative filters are **still always applied**.
4. Scrolls the virtualized list to find ALL results
5. Extracts subject, date, video URL from each email's reading pane (prioritizes YouTube URLs)
6. **In normal mode:** skips emails that already have the processed category (visual check). **In reprocess mode:** includes all emails regardless of category.
7. Deduplicates by email identity (subject + date + link)
8. Saves all remaining emails to `session_state.json`

**Output:** JSON with total count.

**CRITICAL — If 0 emails are found (normal mode):** Stop the entire pipeline and inform the user. Do **NOT** retry without the negative filter. Do **NOT** fall back to reprocess mode. The only way to include already-processed emails is if the user **expressly** requested it in the original prompt.

**If 0 emails are found (Digest mode):** Skip the per-email processing loop (Step 3) but **still execute Step 5** — the digest report is generated from SP data and may contain items registered in previous sessions.

---

### Step 3 — Process Each Email (Per-Email Loop)

Loop over the emails in `session_state.json`. For EACH email, perform Steps 3.0 through 3.8 in sequence. After each email, update the session state and reports.

**Digest mode (standard) — modified step order:** After Step 3.0 (hashtag skip) and Step 3.1 (determine video ID), perform Step 3.3 (dup check) **before** Step 3.2 (video metadata fetch). Use the email's `title` and `yt_id` from Step 3.1 for the dup check. This allows skipping the expensive metadata fetch for SP items that are already complete. See the Digest mode table in Step 3.3 for the decision logic.

**Digest + Reprocess mode:** Follow the same step order as Reprocess mode (3.0 → 3.1 → 3.2 → 3.3 → 3.4 → 3.5 → 3.6 → 3.7 → 3.8).

#### 3.0 — Early Skip: Hashtag-only notifications (JS / MMech)

**Before any other processing**, check the email's `topic` and `subject`:

If `topic` is **"JS"** or **"MMech"** AND the `subject` contains the character **"#"**: this is a social-reach / hashtag notification, not a proper video notification. **Skip** all video processing (Steps 3.1–3.6) entirely. Only perform:
- **Step 3.7** — Categorize + Move the email
- **Step 3.8** — Update session state and reports

Mark the email as `skipped_hashtag = "Yes"` in session state. Do NOT access the video link, do NOT create or update any SP item.

#### 3.1 — Determine Video ID

From the `video_link` field, extract the YouTube video ID (11-char alphanumeric).

- If the URL is a YouTube URL (`youtube.com/watch?v=...` or `youtu.be/...`): extract the ID. This ID is the primary key for dedup and for the SP `yt_ID` field.
- If the URL is NOT a YouTube URL: set `yt_id = None`. Proceed without transcript download.

#### 3.2 — Fetch Video Metadata from YouTube Page

```bash
python pipeline_fetch_video.py "<video_url>"
```

This returns JSON with: `video_id`, `url`, `title`, `published_date`, `duration_seconds`, `duration_formatted`, `description`, `chapters`.

- `published_date`: extracted from YouTube metadata. If empty, fall back to the email's received date.
- `duration_formatted`: already in `Xh YYm` or `Ym` format.
- `chapters`: array of `{title, time, seconds, url}` objects.
- `description`: full description text from the video page.

#### 3.3 — Check for Duplicates

```bash
python pipeline_video_check_dup.py "<title>" "<yt_id>"
```

(Omit `<yt_id>` argument if the video is not from YouTube.)

This returns JSON with: `dup_session`, `dup_sp`, `sp_id`, `sp_has_abstract`.

**Session duplicate** (`dup_session` true): another email in this session already had the same title. Skip Steps 3.4-3.6 entirely and go straight to 3.7 (categorize+move).

**SP item logic — Normal mode**:

| `dup_sp` | `sp_has_abstract` | Action |
|----------|-------------------|--------|
| false | — | Generate abstract (3.4) + Download transcript (3.5) + Create new SP item (3.6) |
| true | false | Generate abstract (3.4) + Download transcript (3.5) + Update existing SP item abstract (3.6) |
| true | true | **Skip** abstract/transcript/SP — go to 3.7 |

**SP item logic — Reprocess mode**:

| `dup_sp` | Action |
|----------|--------|
| false | Generate abstract (3.4) + Download transcript (3.5) + Create new SP item (3.6) |
| true | **Always** regenerate abstract (3.4) + Download transcript (3.5) + Update existing SP item (3.6) |

**SP item logic — Digest mode (standard)** (based on `dup_sp` and `sp_has_abstract`):

Remember: in this mode, Step 3.3 runs **before** Step 3.2.

| `dup_sp` | `sp_has_abstract` | Action |
|----------|-------------------|--------|
| false | — | Fetch video metadata (3.2) + Generate abstract (3.4) + Download transcript (3.5) + Create new SP item (3.6) |
| true | false | Fetch video metadata (3.2) + Generate abstract (3.4) + Download transcript (3.5) + Update existing SP item (3.6) |
| true | true | **Skip** metadata fetch (3.2), abstract (3.4), transcript (3.5), and SP update (3.6) — item is complete |

In all three cases, the SP item data (title, published_date, tech, link, source/topic, abstract, yt_id) will be included in the digest report generated in Step 5. For pre-existing complete items, `pipeline_video_email_report.py` reads these fields directly from SP.

**Digest + Reprocess mode:** Same table as Reprocess mode above (always regenerate, always update).

**Always perform** categorize + move (Step 3.7) for every email, regardless of duplicates.

#### 3.4 — Generate Abstract (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Using the video `description` and `chapters` from Step 3.2, generate a formatted abstract:

- **Language:** Always English, regardless of source language.
- **Format:** HTML formatting with `<b>` for keywords, `<i>` for key phrases, `<ul><li>` for bullet points, `<a href="...">` for hyperlinks.
- **Content — description part:** Produce a concise summary of the video content based on the description. Focus exclusively on what the video covers — the substance.
- **Content — chapters part:** Reproduce the **complete** list of chapters with timestamps and chapter links. Format each chapter as a clickable link to the timestamp, e.g. `<a href="https://www.youtube.com/watch?v=ID&t=123s">0:02:03 - Chapter Title</a>`.
- **Exclusions:** Remove **all** information about the channel, related links, author bios, social media handles, sponsors, etc. Keep ONLY content directly related to the video's subject matter.
- **Style:** Concrete, informative. Report only what the video covers.

Store the abstract in the email's `abstract` field in session state.

#### 3.5 — Download Transcript (if YouTube)

Only if `yt_id` is not None:

```bash
python yt_transcript.py <yt_id>
```

This downloads the transcript to `yt_transcripts/yt_<VIDEO_ID>.txt`.

**Skip if file already exists:** If the file `yt_transcripts/yt_<VIDEO_ID>.txt` already exists (from a previous run or session), do NOT re-download. This applies in both normal and reprocess mode.

If the transcript download fails (no subtitles available), log a warning but **continue** — it is not a blocking error.

#### 3.6 — Classify Technology & SP Item (Create or Update)

**Technology classification** is a Copilot LLM task. Using `tech_map` from `config.json`, assign technology tags:
- If a sub-technology matches (e.g. `Azure / AKV / MHSM`), do NOT also include the parent.
- Comma-separate multiple values.
- Only use values from the map. If no match, leave empty.

Store the tech classification in the email's `tech` field.

**If creating a new SP item** (`dup_sp` was false):

```powershell
Set-Content -Path _sp_input.json -Value '<JSON>' -Encoding utf8
python pipeline_video_sp_create.py _sp_input.json
Remove-Item _sp_input.json
```

The JSON fields:
- `title`: video title
- `published_date`: in `YYYY-MM-DD` format (converted to `YYYY.MM.DD` by the script)
- `abstract`: generated HTML abstract
- `topic`: **MANDATORY** — the value extracted by `pipeline_video_retrieve.py` from the email subject `[Video-Topic]`, already stored in `session_state.json` for each email. This maps to the SP `SourceNew` field via `config.json → source_map`. **Never omit this field and never try to infer it — always use the exact `topic` value from the email's entry in `session_state.json`.**
- `tech`: comma-separated tech tags
- `video_link`: canonical YouTube URL (or original URL if not YouTube)
- `duration`: formatted duration string (e.g. `1h 23m` or `45m`)
- `yt_id`: YouTube video ID (empty string if not YouTube)

**If updating abstract on existing SP item** (`dup_sp` true, `sp_has_abstract` false):

```powershell
Set-Content -Path _sp_input.json -Value '{"abstract":"...","title":"..."}' -Encoding utf8
python pipeline_video_sp_create.py --update-abstract <sp_id> _sp_input.json
Remove-Item _sp_input.json
```

Do NOT pipe JSON via `echo` in PowerShell — non-ASCII characters (hyphens, dashes, accented letters) will be corrupted. Instead, **always write JSON to a temp file with UTF-8 encoding** and pass the file path:

```powershell
Set-Content -Path _sp_input.json -Value '{"title":"...","published_date":"...","abstract":"...","topic":"...","tech":"...","video_link":"...","duration":"...","yt_id":"..."}' -Encoding utf8
python pipeline_video_sp_create.py _sp_input.json
Remove-Item _sp_input.json
```

For updates:

```powershell
Set-Content -Path _sp_input.json -Value '{"abstract":"...","title":"..."}' -Encoding utf8
python pipeline_video_sp_create.py --update-abstract <sp_id> _sp_input.json
Remove-Item _sp_input.json
```

#### 3.7 — Categorize + Move Email in Outlook

```bash
python pipeline_video_email_actions.py both "<email_title>"
```

This script:
1. Searches Outlook for the specific email (sender + `[Video-` + first 6 words)
2. Selects ONLY rows containing `[Video-` (NEVER Ctrl+A)
3. Right-clicks → Categorizza → selects "By agent - Video"
4. Clicks Sposta (top toolbar, y < 200) → searches folder → ArrowDown + Enter

If successful, mark `categorized = "Yes"` and `moved = "Yes"` in session state.

#### 3.8 — Update Reports & Session State

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
- Transcripts downloaded
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

Run `pipeline_video_email_report.py` with the same date range used for email retrieval:

```bash
python pipeline_video_email_report.py --from-date {DATE_FROM} --to-date {DATE_TO}
```

Parse the JSON output. If `total_items` is 0, inform the user and skip sending.

#### 5.2 — Send the Email

Read the HTML file at `html_path` returned in Step 5.1.

Use the **send_email** MCP tool with:
- `emailAddresses`: recipients from `config.json` → `email_report.default_recipients` (or as specified by the user)
- `subject`: the `subject` value from the script output
- `htmlBody`: **depends on file size** —
  - Read `config.json → email_report.max_html_body_size_kb` (default 12) for the size threshold.
  - If the HTML file is **≤ threshold**: use the **complete, unmodified** HTML content read from the file. Do NOT summarize, truncate, or alter it — paste the entire file content verbatim.
  - If the HTML file is **> threshold**: use a **short HTML message** inviting the reader to open the attached report, e.g.: `<html><body style="font-family:'Segoe UI',sans-serif;"><h2 style="color:#4361ee;">PescoPedia Video Digest</h2><p>Period: <b>{date_from} – {date_to}</b></p><p><b>{total_items} videos</b> across <b>{topics_count} topics</b></p><p style="margin-top:20px;">The full report is attached as an HTML file. Please open the attachment for the complete digest.</p></body></html>`
- `attachmentName`: the HTML file name (e.g. `Video_Notifications-Digest-From-2026.04.03-To-2026.04.06.html`)
- `attachmentContent`: the HTML file content encoded as **base64** (always the full file, regardless of body choice)

#### 5.3 — Confirm to User

Report:
- Number of videos included in the digest
- Number of topics
- Date range
- Recipients the email was sent to

---

## Session State Schema

The `session_state.json` file tracks the entire session. Each email in the `emails` array has:

```json
{
  "subject": "[Video-Topic] Video Title",
  "received_date": "2026-04-04T14:32:00",
  "video_link": "https://www.youtube.com/watch?v=...",
  "title": "Video Title",
  "topic": "Topic",
  "yt_id": "abc123XYZ_-",
  "published_date": "2026-04-04",
  "duration_formatted": "1h 23m",
  "description": "Full description text...",
  "chapters": [],
  "tech": "Technology Tag",
  "abstract": "<b>Abstract</b> with HTML formatting...",
  "transcript_file": "yt_transcripts/yt_abc123XYZ_-.txt",
  "dup_session": "",
  "dup_sp": "",
  "sp_created": "",
  "categorized": "",
  "moved": ""
}
```

The `processed_titles` dict maps `title → video_link` for session dedup.

## Report Formats

### XLSX Report

Columns: Published, Title, Tech, Link, Topic, Duration, Abstract, Dup Session, Dup SP, SP Created, Categorized, Moved.

- Header: dark background (#2D3748), white bold text
- Rows colored by topic from `topic_color_palette`
- Status fields show "Yes" in bold green
- Sorted by topic (asc) then published_date (desc)

### HTML Report

- Professional, clean design on light background
- Stats bar at top with counts
- Table of contents linking to topic sections
- Each topic is a section with videos as sub-sections
- Each video shows: linked title, date, tech tags, duration, rich-text abstract
- "Back to index" link after each topic section
- Duplicate articles marked with a badge

---

## Lessons Learned (Mandatory Rules)

1. **Outlook search MUST ALWAYS filter by sender AND subject prefix.** Never run a broad search.

2. **NEVER use Ctrl+A to select search results.** Always iterate rows, check for `[Video-` text, Ctrl+click only verified matches.

3. **SP Link field cannot be set in POST.** `pipeline_video_sp_create.py` handles this with POST + MERGE.

4. **Outlook "Sposta" (Move) button exists in multiple places.** Always use the one in top toolbar (y < 200). Use keyboard navigation (ArrowDown + Enter).

5. **Refresh SP digest token every ~40 items.**

6. **Long titles break Outlook search.** `pipeline_video_email_actions.py` shortens to first 6 words automatically.

7. **Menu items use Italian labels:** "Categorizza", "Sposta", "Cerca una cartella".

8. **NEVER use `-category:` in Outlook Web search.** It silently returns 0 results. Instead, `pipeline_video_retrieve.py` visually checks each email's category in the reading pane.

9. **Outlook Web uses a virtualized list** — only ~8-10 items rendered. `pipeline_video_retrieve.py` scrolls with `mouse.wheel` and waits until 3 stable rounds.

10. **Use `page.evaluate("window.location.href = ...")` for SPA navigation** in Outlook.

11. **NEVER create temp files for pipeline_video_sp_create.py** — pipe JSON via stdin.

12. **YouTube transcript download may fail** (no subtitles). Log warning but continue processing — transcript is supplementary, not required for SP item creation.

## HTML Digest Markers Contract

`pipeline_video_email_report.py` is under the **HTML Digest Structural Markers** contract (see `copilot-instructions.md`). The generated HTML must contain markers M1, M2, M3 exactly once each. **After any edit to this script**, run `python verify_html_markers.py` to confirm.

## Tools Required

- **Python 3.12+** with `playwright`, `openpyxl` packages
- **Microsoft Edge** launched with `--remote-debugging-port=9222` and authenticated debug profile
- **`run_in_terminal`** for executing pipeline Python scripts
- **`yt_transcript.py`** for downloading YouTube transcripts (uses Edge CDP)
