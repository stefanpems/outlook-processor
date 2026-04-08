---
name: vivaengage-notifications
description: "Create a digest from Viva Engage notification emails in Outlook: search, analyze, read thread replies, categorize+move emails, build HTML digest and send by email. Triggered by prompts like 'Crea digest per le notifiche di Viva Engage da ... a ...' or 'Create digest for Viva Engage notifications from date X to date Y'."
argument-hint: "Specify start and end dates, e.g.: 'dal 2026.04.01 al 2026.04.06' or 'from 2026-04-01 to 2026-04-06'"
---

# Viva Engage Notification Digest

## Purpose

Process Viva Engage notification emails from Outlook. For each relevant notification: read the email content, open the Viva Engage thread to read replies/comments, generate structured summaries, categorize and move the email, then build an HTML digest grouped by community and send it by email.

## STRICT RULE: No New Python Code

**NEVER create new Python code** — not as files, not as inline `python -c "..."`, not as heredoc `'@ | python -`. Always use the existing `.py` scripts listed below for every operation. If an existing script does not cover a need, propose modifying an existing script or creating a new permanent script — wait for user approval.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections:

| Key | Purpose |
|-----|---------|
| `viva_engage.notification_sender` | Sender to filter VE notification emails |
| `viva_engage.communities` | List of community names to consider (filter) |
| `viva_engage.processed_category` | Outlook category assigned after processing (`By agent - Viva Engage`) |
| `viva_engage.target_folder` | Outlook folder to move processed emails into (`Social Networks`) |
| `viva_engage.exclude_terms` | List of terms always added as negative filters to every Outlook search (e.g. `["PescoPedia"]`) |
| `viva_engage.default_recipients` | Default email recipients for digest |
| `edge_cdp.url` | CDP endpoint (default `http://localhost:9222`) |

## Pipeline Python Scripts

The workspace contains purpose-built Python scripts for each pipeline step. **Use ONLY these scripts. Do NOT create new scripts that duplicate their functionality.** All scripts read `config.json` at startup.

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `ve-notifications-retrieve.py` | Retrieve matching VE notification emails from Outlook Web via CDP | `python ve-notifications-retrieve.py YYYY-MM-DD YYYY-MM-DD [--include-processed]` |
| `ve-notifications-analyze.py` | Extract detailed content from a specific email in search results | `python ve-notifications-analyze.py <index>` |
| `ve-notifications-process.py` | Open a VE thread URL and read all replies/comments | `python ve-notifications-process.py <thread_url>` |
| `ve-notifications-email-actions.py` | Search, Ctrl+A, categorize and move VE notification emails | `python ve-notifications-email-actions.py "<title>" "<author>" "<community>"` or `--batch-file <json>` |
| `ve-notifications-build-html.py` | Build HTML digest from summaries JSON file | `python ve-notifications-build-html.py --input summaries.json` |

## Input

The user provides a date range. Accept any format — `YYYY.MM.DD`, `YYYY-MM-DD`, `DD/MM/YYYY`, natural language. Parse both dates into `YYYY-MM-DD` format and confirm before starting.

**Default date range:** If no dates are specified, use the full day of the day before execution (from `YYYY-MM-DD` of yesterday to `YYYY-MM-DD` of yesterday).

**Reprocess mode:** If the user requests to reprocess already-processed notifications (via "ri-"/"re-" prefix verbs or explicit phrases like "anche le già processate"), pass `--include-processed` to the retrieve script.

## Procedure

Execute the steps below **strictly in order**. For each step, use the specified Python script via `run_in_terminal`.

---

### Phase 1 — Retrieve VE Notification Emails from Outlook

**Normal mode (default):**

```bash
python ve-notifications-retrieve.py {DATE_FROM} {DATE_TO}
```

**Reprocess mode:**

```bash
python ve-notifications-retrieve.py {DATE_FROM} {DATE_TO} --include-processed
```

The script:
1. Connects to Edge via CDP
2. Navigates to Outlook Web
3. Searches: `from:{notification_sender} received:{date_from}..{date_to} -{exclude_term_1} -{exclude_term_2} ...` (plus `-"By agent - Viva Engage"` in normal mode; exclude terms from `viva_engage.exclude_terms`)
4. Scrolls the virtualized list to find ALL results
5. Extracts subject, date, post type, post title, community, thread URL from each email's reading pane
6. Deduplicates by (subject + date)
7. Saves all to `ve_notifications_cache.json`

**Output:** JSON with total count.

**If 0 emails found:** Inform the user and stop.

---

### Phase 2 — Analyze Email Content

For each email in `ve_notifications_cache.json`, the Copilot agent reads the `body_text` field and extracts/annotates in the cache:

- **received_date**: Date of the email notification
- **subject**: Full email subject
- **post_type**: The type before the first ":" in subject (Question, Announcement, Praise, etc.)
- **post_title**: The part after the first ":" in subject
- **community_name**: Name of the community (from email body, "Pubblicato in ..." link)
- **community_url**: URL of the community
- **author**: Author of the post
- **post_date**: Date the post was published
- **thread_url**: Link to the VE conversation thread

For each email, generate a **summary in English** of the post text (max 100 words, well-formatted using `<b>`, `<i>`, `<ul><li>`, URLs). **Ignore** images, tags, and replies at this stage.

The analysis data is recorded in the cache by the Copilot agent (reading and updating `ve_notifications_cache.json`).

---

### Phase 3 — Process Cached Notifications

Loop through each notification in the cache. Apply the following logic:

#### 3.1 — Community Filter

Check if `community_name` is among the communities listed in `config.json → viva_engage.communities`. If NOT, skip to the next item. If YES, mark the item as `"processed": true` in the cache and proceed with steps 3.2 and 3.3.

#### 3.2 — Read Thread Replies (Viva Engage)

Open the thread URL in the browser to read replies and comments:

```bash
python ve-notifications-process.py "<thread_url>"
```

The script:
1. Opens the thread URL in Edge via CDP
2. Expands all collapsed replies/comments (clicks "see more", "N replies", "Show X more answers")
3. Reads the full thread text
4. Returns JSON with `full_text`

**Copilot LLM Task — Summarize Replies:**

Read the `full_text` and generate structured summaries based on post type:

**If Question:**
Classify contributions as:
- A. Additional doubts from the same or other users on the topic
- B. Elements of response/answer
- C. Follow-up actions committed to
- D. User mentions/tagging (completely ignore these)

Generate these summary fields (in English, well-formatted with `<b>`, `<i>`, `<ul><li>`, URLs):
- **"Additional contributions to the question"** (omit if none)
- **"Answer"** — synthesize all substantive answer elements
- **"Follow-up"** — concrete actions committed to (omit if none)

**If Announcement, Discussion, Praise, or other:**
Classify substantive contributions as "comments" or "other information". Generate:
- **"Comments"** — well-formatted synthesis of all valuable contributions (omit if none)
- **"Follow-up"** — concrete actions committed to (omit if none)

**Skip irrelevant contributions:** Never report comments that only tag/mention people, pure administrative messages, or "+1" responses.

Store the reply summaries in the cache for Phase 4.

#### 3.3 — Categorize and Move Emails in Outlook

Search for **all** notification emails matching this item, select ALL with Ctrl+A, categorize as "By agent - Viva Engage", and move to "Social Networks":

```bash
python ve-notifications-email-actions.py "<notification_title>" "<author>" "<community_name>"
```

The script:
1. Searches the **entire mailbox** for `from:{sender} subject:(<title>) <author> <community>`
2. Clicks the first result row, then Ctrl+A to select ALL results
3. Right-click → Categorizza → selects "By agent - Viva Engage"
4. Clicks Sposta (top toolbar, y < 200) → searches "Social Networks" folder → ArrowDown + Enter

For batch processing, prepare a JSON file and use:

```bash
python ve-notifications-email-actions.py --batch-file <json_file>
```

Where the JSON is a list: `[{"notification_title": "...", "author": "...", "community_name": "..."}]`

---

### Phase 4 — Build and Send HTML Digest

#### 4.1 — Prepare Summaries JSON

Consider only items with `"processed": true` in the cache. Group them by community. Prepare a JSON file (`_ve_notif_summaries.json`) with this structure:

```json
{
  "date_from": "2026-04-03",
  "date_to": "2026-04-06",
  "communities": [
    {
      "community": "Community Name",
      "conversations": [
        {
          "type": "question",
          "title": "Post title",
          "thread_url": "https://engage.cloud.microsoft/...",
          "has_images": false,
          "author": "Author Name",
          "date": "2026-04-05",
          "summary_lines": [
            "<b>Question:</b> Summary of the question...",
            "<b>Author:</b> Author Name — Apr 5",
            "<b>Answer:</b> Summary of answers...",
            "<b>Main responder(s):</b> Name1, Name2",
            "<b>Follow-up:</b> Action committed..."
          ]
        }
      ]
    }
  ]
}
```

The `summary_lines` array uses the same summary format as the `vivaengage-conversations` skill:

**For Questions:**
- `<b>Question:</b>` — well-formatted synthesis (up to 100 words), `<b>` for key terms
- `<b>Author:</b>` — name + date
- `<b>📎 Images attached</b>` — only if actual images present
- `<b>Additional doubts:</b>` — further doubts raised (omit if none)
- `<b>Answer:</b>` — well-formatted synthesis (100-150 words), aggregating ALL answers
- `<b>Main responder(s):</b>` — names (omit if no answer)
- `<b>Follow-up:</b>` — only if concrete action committed (omit otherwise)

**For Announcements/Discussions/Praise:**
- `<b>Post:</b>` — well-formatted synthesis (up to 100 words), `<b>` for key terms
- `<b>Author:</b>` — name + date
- `<b>📎 Images attached</b>` — only if actual images present
- `<b>Comments:</b>` — well-formatted synthesis (100-150 words) (omit if none)
- `<b>Follow-up:</b>` — only if concrete action committed (omit otherwise)

#### 4.2 — Build HTML Report

```bash
python ve-notifications-build-html.py --input _ve_notif_summaries.json
```

This produces an HTML file in `output/Viva_Engage-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.html` and outputs JSON with the `html_path`.

#### 4.3 — Send Email

Read the HTML file at `html_path` returned in Step 4.2.

Use the **send_email** MCP tool with:
- **To:** `config.json → viva_engage.default_recipients` (or user-specified)
- **Subject:** `PescoPedia Viva Engage Digest - From-YYYY.MM.DD To-YYYY.MM.DD`
- **Body:** the **complete, unmodified** HTML content read from the file — verbatim
- **Attachment:** the HTML file (base64-encoded), same filename

#### 4.4 — Cleanup

Remove the temporary files:

```bash
Remove-Item _ve_notif_summaries.json, _ve_batch_actions.json, _ve_attachment_b64.txt -ErrorAction SilentlyContinue
```

The `ve_notifications_cache.json` file is kept for debugging. It will be overwritten on next run.

#### 4.5 — Final Report

Report to the user:
- Number of notifications processed
- Number of communities included
- Number of notifications skipped (community not in config)
- Date range
- Recipients the email was sent to

---

## Summarization Rules

- **Language:** Always English, regardless of the language of the user's prompt or conversation content.
- **Skip irrelevant contributions:** Never report comments that only tag/mention people, "+1", or purely administrative messages.
- **Consistency:** Apply the full summary structure uniformly to every conversation.
- **Follow-up extraction is mandatory:** After summarizing, always re-scan for commitments to action.
- **Coverage:** The summary must mention every substantial concept — nothing should be silently omitted.
- **Style:** Concrete, informative. Report only what is communicated. No filler.

## Cache File Schema

The `ve_notifications_cache.json` file tracks the session:

```json
{
  "date_from": "2026-04-03",
  "date_to": "2026-04-06",
  "emails": [
    {
      "subject": "Question: How to configure X",
      "received_date": "2026-04-05T10:30:00",
      "post_type": "Question",
      "post_title": "How to configure X",
      "community_name": "Defender for Cloud (Internal)",
      "community_url": "https://engage.cloud.microsoft/...",
      "author": "John Doe",
      "post_date": "Apr 5, 2026",
      "thread_url": "https://engage.cloud.microsoft/.../thread/...",
      "body_text": "Full email body text (up to 3000 chars)",
      "post_summary": "<b>Summary</b> of the post...",
      "reply_summary": {
        "additional_doubts": "...",
        "answer": "...",
        "follow_up": "...",
        "comments": "..."
      },
      "processed": true
    }
  ]
}
```

## Tools Required

- **Python 3.12+** with `playwright` package
- **Microsoft Edge** launched with `--remote-debugging-port=9222` and the authenticated debug profile
- **`run_in_terminal`** for executing Python scripts
- **`send_email` MCP tool** for sending the digest email
