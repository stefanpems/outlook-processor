---
name: teams-meeting-recording
description: "Register a Teams meeting recording in the SharePoint VideosMSInt list from its Stream URL. Fetches metadata (title, date, duration), downloads the transcript, generates a summary, classifies tech, creates the SP item. Triggered by prompts like 'registra questo meeting recording' or 'register this Teams meeting recording' or 'aggiungi questa registrazione Teams'."
argument-hint: "Provide the SharePoint Stream URL of the Teams meeting recording."
---

# Teams Meeting Recording — Register in SharePoint

## Purpose

Given a Teams meeting recording URL (hosted on SharePoint Stream), fetch the video metadata, download the transcript, generate a rich-text summary, classify the technology topic, and create a new item in the SharePoint **VideosMSInt** list.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections:

- `teams_meeting.list_api` — SP REST API path for the VideosMSInt list
- `teams_meeting.list_entity_type` — entity type for POST/MERGE (auto-discovered at runtime)
- `teams_meeting.list_url` — SP list URL (used for auth context navigation)
- `teams_meeting.fields` — field internal name mapping: `published`, `summary`, `duration`, `sha256_id`, `long_link`
- `sharepoint.site_base` — SP site base URL for REST calls
- `tech_map` — technology label → SP Tech lookup ID

## Pipeline Python Scripts

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `pipeline_fetch_teams_meeting.py` | Fetch metadata + transcript from a Stream page via CDP | `python pipeline_fetch_teams_meeting.py "<url>"` |
| `pipeline_teams_sp_create.py` | Create a new VideosMSInt SP item via REST API | `python pipeline_teams_sp_create.py <input.json>` |

**Do NOT create new scripts that duplicate their functionality.**

## Input

The user provides one or more SharePoint Stream URLs pointing to Teams meeting recordings. URLs follow this pattern:

```
https://{tenant}-my.sharepoint.com/personal/{user}/_layouts/15/stream.aspx?id=...
```

**Meeting Sender convention:** If the user places text in parentheses immediately before a URL, use that text as the `meeting_sender` value for the SP item. This maps to the `SourceNew` lookup field via `config.json → source_map`. Example:

```
(MSSec) https://microsoft.sharepoint.com/...
```

→ `"meeting_sender": "MSSec"` in the input JSON → `SourceNewId` is set from `source_map["MSSec"]`.

If no parenthesized text precedes the URL, omit `meeting_sender` from the input JSON.

**Auto-detection:** `pipeline_fetch_teams_meeting.py` scans the page text and URL for known keywords (currently: "LevelUp"). If found, the output JSON includes `"detected_meeting_sender": "LevelUp"`. Use this value as `meeting_sender` when calling `pipeline_teams_sp_create.py` **unless** the user explicitly provided a parenthesized sender (explicit always wins).

## Procedure

Execute the steps below **strictly in order**. For multiple URLs, repeat the full procedure for each.

**Same-title rule:** When processing multiple URLs in a single batch and two or more recordings share the same title (as returned by `pipeline_fetch_teams_meeting.py`), append ` - Part X` to each title before creating the SP item, where X is a sequential number (1, 2, 3 …) in the order the URLs were provided by the user. Example: three recordings all titled "Security Workshop" become "Security Workshop - Part 1", "Security Workshop - Part 2", "Security Workshop - Part 3".

---

### Step 1 — Fetch Meeting Metadata and Transcript

```bash
python pipeline_fetch_teams_meeting.py "<url>"
```

This script:
1. Connects to Edge via CDP (auto-launches Edge if needed)
2. Navigates to the Stream page with **all media muted** (volume 0, muted attribute)
3. Extracts:
   - **Title** — from the page DOM or the filename in the URL
   - **Published date** — from page metadata or from the filename date pattern `YYYYMMDD_HHMMSS`
   - **Duration** — from the `<video>` element. If not available, briefly plays the video (muted) to load metadata, then pauses.
4. Opens the transcript panel and scrapes all transcript text
5. Computes the **SHA256 hex digest** of the URL as-is (not URL-decoded)
6. Saves the transcript to `teams_transcripts/<sha256>.txt`

**Output:** JSON to stdout:
```json
{
  "title": "Workshop Wednesday GSA – Explicit Forward Proxy",
  "published_date": "2026-04-08",
  "duration_seconds": 3720,
  "duration_formatted": "1h 02m",
  "sha256_id": "a1b2c3d4...",
  "transcript_path": "teams_transcripts/a1b2c3d4....txt",
  "transcript_length": 28450,
  "url": "https://...",
  "detected_meeting_sender": "LevelUp"
}
```

> `detected_meeting_sender` is only present when a known keyword (e.g. "LevelUp") is found on the page.

**If transcript extraction fails:** The script saves a debug screenshot (`debug_teams_meeting.png`). Inspect the screenshot to identify the correct DOM selectors for the transcript panel, then modify the script if needed. Do not proceed to Step 2 without a transcript — the summary depends on it.

**If duration is 0:** The page might need more time to load the video metadata. Retry by navigating manually or check the recording's properties page.

Parse the JSON output. Store all values for use in subsequent steps.

---

### Step 2 — Read the Transcript

Read the transcript file saved in Step 1:
- Path: value of `transcript_path` from Step 1 output
- The transcript contains speaker-attributed text with timestamps

Store the full transcript text for use in Steps 3 and 4.

---

### Step 3 — Generate Summary (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Using the transcript from Step 2, generate a formatted summary:

- **Length:** 100–200 words. Be comprehensive — include ALL concepts cited in the transcript.
- **Language:** Always English, regardless of the transcript's language.
- **Format:** HTML rich text with:
  - `<b>` for **keywords** (product names, technologies, key terms)
  - `<i>` for *key concepts* and *important phrases*
  - `<ul><li>` for bullet-point lists (schematic decomposition of concepts)
- **Content:** Synthesize the entire transcript. Cover every significant topic, decision, demo, or concept mentioned. Do NOT omit topics to save space — use concise phrasing instead.
- **Exclusions:** Remove pleasantries, filler, administrative scheduling talk. Keep ONLY substantive content.
- **Style:** Concrete, informative, structured. Open with a one-sentence overview, then use bullet points for the detailed breakdown.

Store the summary HTML string for Step 5.

---

### Step 4 — Classify Technology (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Using the transcript content and `tech_map` from `config.json`, assign technology tags:

- **Rules:**
  - Match technologies mentioned in the transcript to values in `tech_map`.
  - If a sub-technology matches (e.g. `Azure / AKV / MHSM`), do NOT also include the parent (`Azure / AKV` or `Azure`).
  - Comma-separate multiple values.
  - Only use values that exist in `tech_map`. If no match, leave empty.
- **Source:** Base classification on the actual transcript content, not just the title.

Store the tech classification string for Step 5.

---

### Step 5 — Create SharePoint Item

First, format the data:
- `published_date`: from Step 1, in `YYYY-MM-DD` format (the script converts to `YYYY.MM.DD`)
- `title`: from Step 1
- `summary`: from Step 3
- `tech`: from Step 4
- `duration`: **MANDATORY** — from Step 1 (`duration_formatted`). Format: `Xh YYm` (e.g. `1h 06m`) or `Ym` if minutes are single-digit (e.g. `8m`). Must always be provided in the input JSON as either `"duration"` or `"duration_formatted"`.
- `sha256_id`: from Step 1
- `video_link`: the original URL provided by the user
- `meeting_sender`: (optional) the parenthesized text preceding the URL, if provided by the user. Maps to `SourceNewId` via `config.json → source_map`. **Priority:**
  1. Explicit parenthesized text from user
  2. `detected_meeting_sender` from Step 1 output
  3. **Technology-based fallback:** If neither (1) nor (2) is available, attempt to infer a source from the meeting's primary technology. Look at the `source_map` keys in `config.json` and consider **only technology-oriented entries** (e.g. `Entra`, `MDC`, `MDE`, `MDVM`, `MDO`, `MDXDR`, `Sentinel`, `Purview`, `SCP`, `AzNet`, `AzMon`, `AzArc`, `ConfComp`, `MSSec`, `AAIFoundry`, `MDA`, `MDI`), **excluding community/series entries** (e.g. `LevelUp`, `Ninja`, `CCP`, `SSA`, `EEC`, `IIC`, `ECS`, `IdAdv`, `JS`, `CPS`, `Other`). Using the transcript content and the tech classification from Step 4, identify the single primary technology of the meeting and match it to the closest technology source key. If a clear match exists, use it as `meeting_sender`.
  4. Omit — if no match is found at any level, leave `meeting_sender` empty.

Write the JSON to a temp file and run:

```powershell
Set-Content -Path _sp_input.json -Value '<JSON>' -Encoding utf8
python pipeline_teams_sp_create.py _sp_input.json
Remove-Item _sp_input.json
```

Do NOT pipe JSON via `echo` — non-ASCII characters get corrupted. Always use `Set-Content -Encoding utf8`.

The script:
1. Navigates to the SP list page (for auth context)
2. Auto-discovers the entity type from the list metadata
3. **Checks for duplicates** by querying for existing items with the same `ID_SHA256` field value
4. If a duplicate exists → returns `{"ok": false, "duplicate": true, ...}`. Inform the user and stop.
5. Creates the new SP item via POST
6. Sets the Link field via a separate MERGE request (`SP.FieldUrlValue`)

**Long URLs (>255 chars):** SharePoint `SP.FieldUrlValue` (Hyperlink) fields have a 255-char limit. Teams meeting URLs frequently exceed this. When the URL is longer than 255 characters, the script:
- Stores the **full original URL** in the `LongLink` field (a multi-line plain text field, internal name `field_11`) during the POST
- **Skips** the MERGE on the `Link` Hyperlink field entirely (no truncation, no broken links)
- Returns `link_skipped` in the output JSON to signal this

When the URL is ≤ 255 characters, the `Link` Hyperlink field is set normally via MERGE, and `LongLink` is not populated.

**Output:** JSON with `ok`, `id`, `title`, and optionally `link_warning`.

---

### Step 6 — Confirm to User

Report:
- Title of the recording
- Published date
- Duration
- Technologies classified
- SP item ID created
- Transcript file path and size
- Summary length (words)

---

## Error Handling

| Error | Action |
|-------|--------|
| CDP not reachable | `cdp_helper.py` auto-launches Edge. If it still fails, ask user to close Edge and relaunch manually (Step 1 in copilot-instructions). |
| Transcript extraction fails | Inspect `debug_teams_meeting.png`. The DOM selectors may need updating for the specific Stream page version. |
| Duration is 0 | Try reloading the page or check the video properties panel. |
| SP duplicate detected | Inform the user. The recording was already registered. |
| SP entity type mismatch | The script auto-discovers the entity type. If the config value was wrong, it logs the correct one. Update `config.json` if needed. |
| SP field name error | Check `config.json → teams_meeting.fields` against actual SP list column internal names. |

## Notes

- The SHA256 digest is computed on the **raw URL as provided** (not URL-decoded). This ensures uniqueness even if the same video is shared with different URL parameters.
- Transcript files are saved in `teams_transcripts/` with the SHA256 as filename. They persist across sessions.
- The `tech_map` in `config.json` is the same one used by the video-notifications skill. Both skills share the same technology taxonomy.
