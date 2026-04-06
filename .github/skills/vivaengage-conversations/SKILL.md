---
name: vivaengage-conversations
description: "Summarize recent conversations from Viva Engage communities: questions, announcements, discussions and their replies. Triggered by prompts like 'riassumi le conversazioni delle community Viva Engage degli ultimi 3 giorni' or 'summarize Viva Engage conversations from the last 7 days' or 'novità dalla community Defender for Cloud'."
argument-hint: "Specify the number of days, e.g.: 'ultimi 3 giorni' or 'last 7 days'"
---

# Viva Engage Conversations Summary

## Purpose

Read and summarize recent conversations from one or more Viva Engage communities configured in `config.json`. For each community, browse the feed sorted by **Recent activity**, read each conversation (post + all replies/comments), and produce a structured summary. Stop reading when a conversation's dates (post date + all reply dates) are **all** older than the specified time window.

## Configuration

All parameters are in `config.json` at the workspace root:

| Key | Purpose |
|-----|---------|
| `viva_engage.communities` | List of community names to read (must match sidebar link text exactly) |
| `viva_engage.default_days` | Default number of days to look back (typically 1) |
| `viva_engage.default_recipients` | Default email recipients for digest mode (semicolon-separated) |
| `edge_cdp.url` | CDP endpoint (default `http://localhost:9222`) |
| `edge_cdp.profile_name` | Edge profile directory name for the debug profile |

## Python Scripts

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `engage_read_conversations.py` | Read conversations from a single community via CDP | `python engage_read_conversations.py "<community_name>" <days>` |
| `engage_build_html.py` | Build HTML report from conversation summaries JSON | `python engage_build_html.py --input summaries.json` |

### Script Output

`engage_read_conversations.py` writes JSON to **stdout** and progress to **stderr**:

```json
{
  "community": "Community Name",
  "days": 3,
  "cutoff_date": "2026-04-03",
  "reference_date": "2026-04-06 15:30",
  "total_conversations": 4,
  "conversations": [
    {
      "type": "question",
      "heading": "Author, Topic text...",
      "thread_url": "https://engage.cloud.microsoft/main/groups/.../thread/...",
      "has_images": false,
      "raw_text": "Full text of post + all replies...",
      "dates": ["2026-04-04 15:19", "2026-04-04 17:27"],
      "most_recent": "2026-04-04"
    }
  ]
}
```

## Input

The user may provide:
- A **number of days** (or a date range) — how far back to look for activity.
- One or more **community names** — which communities to read.

**Defaults** (when the user does not specify):
- **Communities:** all communities listed in `config.json → viva_engage.communities`
- **Days:** `config.json → viva_engage.default_days` (typically 1 = yesterday's full day)

If the user names a specific community, operate on that one only.  Otherwise, operate on **all** communities from config.

## HTML Template Color Scheme

The HTML digest uses the following colors for conversation type badges (defined in `engage_build_html.py → TYPE_LABELS`):

| Element | Color | Hex |
|---------|-------|-----|
| Question badge | Red | `#e74c3c` |
| Announcement badge | Green | `#2ecc71` |
| Discussion badge | Purple | `#9b59b6` |
| Community item count badge | Blue | `#4361ee` |

## Output Modes

This skill has two output modes depending on how the user triggers it:

### Mode 1 — Process conversations (HTML file)

**Trigger phrases:** "processa le conversazioni", "process conversations", "riassumi le conversazioni", "summarize conversations", "novità dalla community"

Produces a well-formatted **HTML file** saved to the `output/` directory (named `YYYY-MM-DD-VivaEngageDigest.html`). The agent presents the summaries in the chat AND generates the HTML file.

### Mode 2 — Create digest / Send email

**Trigger phrases:** "crea digest", "create digest", "invia digest", "send digest", "manda il report", "send Viva Engage email"

Same processing as Mode 1, but **additionally sends the HTML as the body of an email** to the default recipients from `config.json → viva_engage.default_recipients`. Uses the `send_email` MCP tool.

Email details:
- **To:** `config.json → viva_engage.default_recipients` (can be overridden if user specifies recipients)
- **Subject:** `PescoPedia Viva Engage Digest - From: YYYY.MM.DD To: YYYY.MM.DD` (using the date range of the digest window)
- **Body:** the full HTML content produced by `engage_build_html.py`

## Procedure

### Prerequisites

- Edge must be running with the debug profile and CDP port (see workspace setup instructions)
- The user must be authenticated on Viva Engage in the debug profile

### Step 0 — Read Configuration

Read `config.json` to get:
- The list of communities from `viva_engage.communities`
- The CDP URL from `edge_cdp.url`

### Step 1 — For Each Community, Retrieve Conversations

For each community, run:

```bash
python engage_read_conversations.py "<community_name>" <days>
```

The script:
1. Connects to Edge via CDP
2. Navigates to the community in Viva Engage
3. Verifies filters: **All conversations** + **Recent activity** sort
4. Scrolls the virtualized feed, reading each conversation's full text (post + expanded replies)
5. Parses dates from each conversation
6. **Stops** when it finds the first conversation where ALL dates are older than X days
7. Outputs JSON to stdout

**Capture the JSON output** for use in Step 2.

If the script reports 0 conversations, inform the user and move to the next community.

### Step 2 — Summarize Conversations (Copilot LLM Task)

For each conversation in the JSON output, read the `raw_text` field and generate a summary following the format below.

The title of each conversation is a **markdown link** to the thread. Use the `thread_url` field from the JSON output.

#### If the conversation type is `question`:

- **Question:** well-formatted rich-text synthesis (up to **100 words**; can be shorter only if all key concepts are captured). Use `<b>` for key terms. Must cover every core concept from the original question — do not oversimplify.
- **Author:** name + creation date + last modified date (e.g. "John Doe — Apr 3 (last activity: Apr 5)"). The last modified date is the most recent date found in the conversation's `dates` array.
- **📎 Images attached** — only include this line when there are **actual embedded images** (screenshots, diagrams, photos) in the conversation post or replies. Do NOT include it when the conversation only contains hyperlinks to external pages. The `has_images` field from the JSON output is a hint but must be verified by checking the `raw_text` — if it only mentions URLs/links without actual image content, omit this line.
- **Additional doubts:** further doubts or related questions raised by other users in the thread (if any; omit this line if none)
- **Answer:** well-formatted rich-text synthesis (**100–150 words**), aggregating relevant information from ALL reply contributions. Dry, evidence-focused style. Use `<b>`, `<i>`, and `<ul><li>` bullet points for readability. **Every concept that is part of the answer must be at least mentioned** — do not drop information to save space. If no substantive answer has been provided yet, write: *"No answer yet."*
- **Main responder(s):** names of main responders (omit if no answer yet)
- **Follow-up:** only include this line when someone committed to a concrete action (e.g. "Diana Grigore will investigate with the team"). Omit for pure acknowledgements or tagging.

#### If the conversation type is `announcement` or `discussion`:

- **Post:** well-formatted rich-text synthesis (max 100 chars). Use `<b>` for key terms.
- **Author:** name + creation date + last modified date (e.g. "Yuri Diogenes — Apr 3 (last activity: Apr 5)"). The last modified date is the most recent date found in the conversation's `dates` array.
- **📎 Images attached** — only include this line when there are **actual embedded images** (screenshots, diagrams, photos) in the conversation post or replies. Do NOT include it when the conversation only contains hyperlinks to external pages.
- **Comments:** well-formatted rich-text synthesis (**100–150 words**), aggregating relevant comment contributions. Dry, evidence-focused style. Use `<b>`, `<i>`, and `<ul><li>` bullet points for readability. **Every concept that is part of the comments must be at least mentioned.** Omit this line if no substantive comments.
- **Follow-up:** only include this line when someone committed to a concrete action. Omit for pure acknowledgements or tagging.

#### Summarization rules

- **Language:** Always English, regardless of the language of the user's prompt or the conversation content.
- **Skip irrelevant contributions:** never report comments that only tag/mention people (e.g. "Adding @X to track this", "+1", "@person"), or comments that are purely administrative (e.g. "share your tenantID privately").
- Focus only on **substantive content contributions** that add information, answers, or context.
- **Consistency across all conversations:** Apply the full summary structure (Question/Post, Author, Images, Additional doubts, Answer/Comments, Main responders, Follow-up) uniformly to **every single conversation** from the first to the last. Do NOT degrade quality or skip fields (especially Follow-up) for later conversations. Every conversation deserves the same level of analysis.
- **Follow-up extraction is mandatory for every conversation:** After summarizing the answer/comments, always re-scan the full `raw_text` looking for commitments to action (e.g. "I will check", "will track", "will investigate", "filing a bug", "opening a ticket"). If found, include the Follow-up line. This check must happen for ALL conversations, not just the first few.

### Step 3 — Build HTML Report

After summarizing all communities, prepare a JSON file with the following structure and save it to a temporary file (e.g. `_ve_summaries.json`):

```json
{
  "date_label": "2026-04-06",
  "days": 1,
  "communities": [
    {
      "community": "Community Name",
      "conversations": [
        {
          "type": "question",
          "title": "Short title for the conversation",
          "thread_url": "https://engage.cloud.microsoft/...",
          "has_images": false,
          "author": "Author Name",
          "date": "2026-04-05",
          "summary_lines": [
            "<b>Question:</b> How to configure X in environment Y?",
            "<b>Author:</b> John Doe — Apr 5",
            "<b>Answer:</b> Use setting Z in the portal. <b>Key:</b> ensure AAD is synced first.",
            "<b>Main responder(s):</b> Jane Smith"
          ]
        }
      ]
    }
  ]
}
```

The `summary_lines` array contains the formatted summary lines from Step 2 (each line as an HTML fragment: `<b>Question:</b> ...`, `<b>Answer:</b> ...`, etc.). Include the 📎 Images line and Follow-up line only when applicable.

Then run:

```bash
python engage_build_html.py --input _ve_summaries.json
```

This produces an HTML file in `output/YYYY-MM-DD-VivaEngageDigest.html` and outputs JSON with the `html_path`.

### Step 4 — Deliver Results

#### Mode 1 (Process conversations)

Present the summaries in the chat grouped by community (same format as before), and inform the user that the HTML report has been saved:

```
## Community Name 1

### 1. [Question] [Short title](thread_url)
- **Question:** ...
- **Author:** ... — date
- **Answer:** ...
- **Main responder(s):** ...

### 2. [Announcement] [Short title](thread_url) 📎
- **Post:** ...
- **Author:** ... — date
- **Comments:** ...
- **Follow-up:** ...

## Community Name 2

### 1. ...

---
HTML report saved to `output/YYYY-MM-DD-VivaEngageDigest.html`.
```

#### Mode 2 (Create digest / Send email)

Same as Mode 1, but additionally:

1. Read the generated HTML file content
2. Use the `send_email` MCP tool to send the email:
   - **To:** recipients from `config.json → viva_engage.default_recipients` (or user-specified)
   - **Subject:** `PescoPedia Viva Engage Digest - From: YYYY.MM.DD To: YYYY.MM.DD`
   - **Body:** the full HTML content
3. Confirm to the user that the email was sent and the HTML file was saved

## How It Works — Technical Details

### Virtualized Feed

Viva Engage renders only 2-3 threads at a time in the DOM. The script scrolls a specific overflow container (identified by `scrollHeight > 5000`) to trigger lazy loading.  Old threads are removed from the DOM as new ones appear, so each thread's text is captured immediately when its heading first appears.

### Thread Identification

Thread headings have DOM id `heading-thread-*`.  The script uses these as anchors to:
1. Detect new threads as they appear in the viewport
2. Delimit each thread's text portion within the full page text

### Date Parsing

Viva Engage shows relative dates.  The script handles:
- `Xh` / `Xm` / `Xd` — relative time ago
- `Yesterday at H:MM AM/PM`
- `Day at H:MM AM/PM` — weekday-based (e.g. "Fri at 3:19 PM")
- `Mon DD` — month and day of the current/previous year
- `Mon DD, YYYY` — full date

### Content Expansion

Before reading a thread, the script clicks:
- `see more` buttons to reveal truncated post bodies
- `N replies` / `Show X more answers` buttons to expand collapsed reply chains
- It **never** clicks `see less`, `collapse`, or `hide` buttons

### Stopping Logic

Since the feed is sorted by **Recent activity**, the most recently active threads come first. The script stops at the first thread where **all** parsed dates (post date + every reply/comment date) are older than the cutoff.  This guarantees that:
- Threads with old posts but recent replies are included
- No active thread within the window is missed

## Lessons Learned

1. **Click via JS** — Viva Engage overlays can block Playwright clicks. Use `element.evaluate("e => e.click()")`.
2. **Dismiss popups** — Press `Escape` after navigation to close any "Catch me up" banners.
3. **SPA navigation works** — Unlike Outlook, `page.goto()` works fine for Engage (no need for `window.location.href` workaround).
4. **Community must be in sidebar** — The script locates communities by matching link text. Ensure the community is in Favorites or a visible section.
5. **Edge must stay open** — Playwright connects via CDP; closing Edge breaks the connection.

## Tools Required

- **Python 3.12+** with `playwright` package
- **Microsoft Edge** launched with `--remote-debugging-port=9222` and the authenticated debug profile
- **`run_in_terminal`** for executing the Python script
