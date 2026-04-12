---
name: youtube-video
description: "Register one or more YouTube videos in the SharePoint VideoPosts list from their URLs. Fetches metadata, downloads transcript, generates abstract, classifies tech, creates SP items. Triggered by prompts like 'registra questo video YouTube' or 'register this YouTube video' or 'aggiungi questo video alla lista'."
argument-hint: "Provide one or more YouTube video URLs, e.g.: 'registra https://www.youtube.com/watch?v=bEP3upJcurQ'"
---

# Register YouTube Video

## Purpose

Register one or more YouTube videos into the SharePoint VideoPosts list by providing their URLs directly. For each URL: fetch video metadata from YouTube (title, duration, published date, description, chapters), download the transcript, generate a formatted abstract, classify the technology, check for duplicates, and create the SP list item.

This skill is independent of the email notification pipeline — it works with any YouTube video URL.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections:

- `video_sharepoint.list_api` — SP REST API base for the VideoPosts list
- `video_sharepoint.list_entity_type` — entity type for POST/MERGE
- `video_sharepoint.list_url` — SP list URL
- `video_sharepoint.fields` — field internal name mapping (published, abstract, duration, yt_id)
- `sharepoint.site_base` — SP site base URL for REST calls
- `source_map` — topic name → SP SourceNew lookup ID (this skill uses `Other` by default, or a user-specified topic)
- `tech_map` — technology label → SP Tech lookup ID

## Pipeline Python Scripts Used

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `pipeline_fetch_video.py` | Fetch YouTube video metadata (title, date, duration, description, chapters) | `python pipeline_fetch_video.py <url>` |
| `yt_transcript.py` | Download YouTube transcript to `yt_transcripts/yt_<VIDEO_ID>.txt` | `python yt_transcript.py <video_id_or_url>` |
| `pipeline_video_check_dup.py` | Check duplicates by title and yt_id (session + SP) | `python pipeline_video_check_dup.py <title> [<yt_id>]` |
| `pipeline_video_sp_create.py` | Create/update VideoPosts SP item via REST API | `python pipeline_video_sp_create.py _sp_input.json` |
| `pipeline_fetch_videoposts.py` | Fetch all existing SP VideoPosts → `sp_videoposts.json` | `python pipeline_fetch_videoposts.py` |

**Do NOT create new scripts that duplicate their functionality.**

## Input

The user provides one or more YouTube URLs. Extract all URLs from the user's message. Confirm the list before starting.

The user may optionally specify a **topic** (one of the keys in `config.json → source_map`). If not specified, default to `"Other"`.

## Procedure

Execute the steps below **strictly in order**. For multiple URLs, repeat Steps 1.1–1.7 for each.

---

### Step 0 — Fetch Existing SP VideoPosts (for deduplication)

```bash
python pipeline_fetch_videoposts.py
```

Downloads all existing SP VideoPosts items to `sp_videoposts.json`. Needed for duplicate checks.

---

### Step 1 — Process Each URL

For EACH URL provided, perform Steps 1.1 through 1.7 in sequence.

#### 1.1 — Extract Video ID and Fetch Metadata

Extract the YouTube video ID (11-char alphanumeric) from the URL (`youtube.com/watch?v=...` or `youtu.be/...`).

```bash
python pipeline_fetch_video.py "<url>"
```

This script:
1. Connects to Edge via CDP
2. Navigates to the YouTube page with **all video elements muted** via `MutationObserver` init script (prevents audio playback)
3. Extracts metadata from `ytInitialPlayerResponse` and DOM

Returns JSON with: `video_id`, `url`, `title`, `published_date`, `duration_seconds`, `duration_formatted`, `description`, `chapters`.

- `published_date`: extracted from YouTube metadata. **If empty**, ask the user.
- `duration_formatted`: already in `Xh YYm` or `Ym` format.
- `chapters`: array of `{title, time, seconds, url}` objects.
- `description`: full description text from the video page.

#### 1.2 — Check for Duplicates

```bash
python pipeline_video_check_dup.py "<title>" "<yt_id>"
```

Returns JSON with: `dup_session`, `dup_sp`, `sp_id`, `sp_has_abstract`.

| `dup_sp` | `sp_has_abstract` | Action |
|----------|-------------------|--------|
| false | — | Proceed to 1.3 + 1.4 + 1.5 + 1.6 (create new SP item) |
| true | false | Proceed to 1.3 + 1.4 + 1.5, then **update** existing SP item abstract (1.6) |
| true | true | **Skip** — inform user the video is already registered with abstract. Ask if they want to regenerate. |

**Session duplicate** (`dup_session` true): another URL in this batch already had the same title. **Skip** entirely.

#### 1.3 — Download Transcript

```bash
python yt_transcript.py <yt_id>
```

Downloads the transcript to `yt_transcripts/yt_<VIDEO_ID>.txt`.

**Skip if file already exists:** If the file `yt_transcripts/yt_<VIDEO_ID>.txt` already exists, do NOT re-download.

If the transcript download fails (no subtitles available), log a warning but **continue** — transcript is supplementary, not required for SP item creation. Generate the abstract from description and chapters only.

#### 1.4 — Generate Abstract (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Using the video `description`, `chapters` from Step 1.1, and the **transcript** from Step 1.3 (if available), generate a formatted abstract:

- **Language:** Always English, regardless of source language.
- **Format:** HTML formatting with `<b>` for keywords, `<i>` for key phrases, `<ul><li>` for bullet points, `<a href="...">` for hyperlinks.
- **Content — description part:** Produce a concise summary of the video content based on the description and transcript. Focus exclusively on what the video covers — the substance.
- **Content — chapters part:** Reproduce the **complete** list of chapters with timestamps and chapter links. Format each chapter as a clickable link to the timestamp, e.g. `<a href="https://www.youtube.com/watch?v=ID&t=123s">0:02:03 - Chapter Title</a>`.
- **Exclusions:** Remove **all** information about the channel, related links, author bios, social media handles, sponsors, etc. Keep ONLY content directly related to the video's subject matter.
- **Style:** Concrete, informative. Report only what the video covers.

#### 1.5 — Classify Technology (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Using `tech_map` from `config.json`, assign technology tags based on the video content (description, transcript, chapters):

- If a sub-technology matches (e.g. `Azure / AKV / MHSM`), do NOT also include the parent (e.g. `Azure`).
- Comma-separate multiple values if the video covers multiple technologies.
- Only use values that exist as keys in `tech_map`. If no match, use `** Other Tech **`.

#### 1.6 — Create or Update SP Item

**If creating a new SP item** (`dup_sp` was false):

```powershell
Set-Content -Path _sp_input.json -Value '<JSON>' -Encoding utf8
python pipeline_video_sp_create.py _sp_input.json
Remove-Item _sp_input.json
```

The JSON must contain:
- `title`: video title (from Step 1.1)
- `published_date`: in `YYYY-MM-DD` format (from Step 1.1)
- `abstract`: generated HTML abstract (from Step 1.4)
- `topic`: the topic specified by the user, or `"Other"` if not specified — maps to SP SourceNew via `config.json → source_map`
- `tech`: comma-separated tech tags (from Step 1.5)
- `video_link`: canonical YouTube URL
- `duration`: formatted duration string (from Step 1.1, e.g. `1h 23m` or `45m`)
- `yt_id`: YouTube video ID

**If updating abstract on existing SP item** (`dup_sp` true, `sp_has_abstract` false):

```powershell
Set-Content -Path _sp_input.json -Value '{"abstract":"...","title":"..."}' -Encoding utf8
python pipeline_video_sp_create.py --update-abstract <sp_id> _sp_input.json
Remove-Item _sp_input.json
```

**IMPORTANT:** Do NOT pipe JSON via `echo` in PowerShell — non-ASCII characters will be corrupted. **Always write JSON to a temp file** with `Set-Content -Encoding utf8` and pass the file path.

#### 1.7 — Confirm to User

For each video, report:
- Title
- Published date
- Duration
- Technologies classified
- Transcript file path and size (if downloaded)
- SP action taken (created / updated / skipped)
- SP item ID

---

### Step 2 — Final Summary

After all URLs are processed, print a summary table:

| URL | Title | Published | Duration | Tech | SP Action | SP ID |
|-----|-------|-----------|----------|------|-----------|-------|

Include: total processed, duplicates found, SP items created, SP items updated, transcripts downloaded, any errors.

---

## Error Handling

| Error | Action |
|-------|--------|
| CDP not reachable | `cdp_helper.py` auto-launches Edge. If it still fails, ask user to close Edge and relaunch manually. |
| Not a YouTube URL | Reject with error. This skill only handles YouTube videos. |
| Video metadata fetch fails | Retry once. If still failing, report error and skip this URL. |
| Transcript download fails | Log warning, continue. Generate abstract from description/chapters only. |
| SP duplicate with abstract | Inform user. Ask if they want to regenerate. |
| SP creation fails | Report the error details. Check digest token freshness. |

## Notes

- **Audio is always muted:** `pipeline_fetch_video.py` uses a `MutationObserver` init script to mute all `<video>` elements before they start playing. No audio will play during metadata extraction.
- **Transcripts persist:** Files in `yt_transcripts/` persist across sessions. If a transcript was already downloaded, it is reused.
- **The `tech_map` in `config.json`** is shared with the video-notifications and teams-meeting-recording skills — same technology taxonomy.
- **Topic defaults to `"Other"`** unless the user explicitly provides one. Valid topics are the keys in `config.json → source_map`.
