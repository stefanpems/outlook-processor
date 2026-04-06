---
name: youtube-transcript-downloader
description: "Download the transcript/subtitles of a YouTube video to a local text file. Triggered by prompts like 'scarica il transcript del video YouTube' or 'download YouTube video transcript' or 'estrai i sottotitoli da YouTube'."
argument-hint: "Provide a YouTube video URL or ID, e.g.: 'https://www.youtube.com/watch?v=2C6G9M1aOko' or '2C6G9M1aOko'"
---

# YouTube Transcript Downloader

## Purpose

Download the transcript (auto-generated or manual subtitles) from a YouTube video and save it as a text file in the workspace root. Works for any video that has a transcript available on YouTube.

## How It Works

The script uses **Playwright CDP connected to the Edge debug instance** to:

1. Navigate to the YouTube video page inside the authenticated Edge browser
2. Open the transcript panel (via description expand, three-dot menu, or JS fallback)
3. Scrape all `ytd-transcript-segment-renderer` elements from the DOM
4. Save the text to `yt_<VIDEO_ID>.txt`

This approach bypasses the YouTube `timedtext` API rate limiting (429 errors) because the transcript data is loaded by YouTube's own frontend within the browser session — no additional API calls are made.

## Prerequisites

- **Edge running with CDP**: `--remote-debugging-port=9222` (see main copilot-instructions.md for setup)
- **Playwright for Python**: `pip install playwright`

## Script

| File | Purpose | CLI Usage |
|------|---------|-----------|
| `yt_transcript.py` | Download transcript from a YouTube video | `python yt_transcript.py <video_id_or_url> [--clean]` |

### Arguments

- `video_id_or_url` — YouTube video ID (e.g. `2C6G9M1aOko`) or full URL (e.g. `https://www.youtube.com/watch?v=2C6G9M1aOko`)
- `--clean` — Optional. Strip timestamp lines, output only spoken text.

### Output

- File: `yt_transcripts/yt_<VIDEO_ID>.txt` — naming convention: `yt_` prefix + video ID
- The `yt_transcripts/` directory is created automatically and listed in `.gitignore`
- Last line of stdout is a JSON object: `{"video_id": "...", "title": "...", "file": "...", "chars": ...}`

## Procedure

### Step 1 — Ensure Edge CDP is Running

Verify CDP connectivity:

```
Invoke-RestMethod -Uri "http://localhost:9222/json/version"
```

If Edge is not running, follow the setup in the main copilot-instructions.md.

### Step 2 — Run the Script

```
python yt_transcript.py <VIDEO_URL_OR_ID>
```

For clean text without timestamps:

```
python yt_transcript.py <VIDEO_URL_OR_ID> --clean
```

### Step 3 — Verify Output

Read the first and last lines of the transcript file to confirm completeness.

## Important Notes

- **Works with auto-generated subtitles** — most YouTube videos have these even without manual captions.
- **Language**: the script downloads whatever language the transcript panel shows (typically the video's original language). YouTube auto-selects the default.
- **Already-open panel**: if the transcript panel is already open from a previous run/session on the same tab, the script detects this and skips the panel-opening step.
- **Long videos** (1h+): fully supported — YouTube loads all transcript segments into the DOM at once.
- **No transcript available**: if the video has no subtitles at all, the script exits with an error and saves a debug screenshot.
- **Rate limiting**: this approach does NOT hit the `timedtext` API, so 429 errors do not apply.
