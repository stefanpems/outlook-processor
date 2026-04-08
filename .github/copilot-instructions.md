# Copilot Workspace Instructions — Authenticated Edge Browser Automation

## Overview

This workspace uses **Microsoft Edge with remote debugging** to allow GitHub Copilot to automate browser interactions on authenticated corporate pages (Outlook Web, SharePoint, etc.) that require Conditional Access compliance.

The VS Code integrated browser (Simple Browser/Chromium) **cannot** pass corporate Conditional Access policies. Instead, we launch Edge with the user's authenticated profile and connect to it via **Chrome DevTools Protocol (CDP)** using Playwright.

## Workspace Python Files — USE THESE, DO NOT RECREATE

The workspace contains tested, working Python scripts for the blog notification pipeline. **Always use these files. Never create new scripts that duplicate their functionality.** All scripts load parameters from `config.json` at startup.

### STRICT RULE: No New Python Code

**NEVER create new Python code** — not as files, not as inline `python -c "..."`, not as heredoc `'@ | python -`. Always use the existing `.py` scripts listed below for every operation: searching, checking, processing, debugging.

If an existing script does not cover a need:

1. **Propose modifying an existing script** — explain what change is needed and why.
2. **If no existing script fits**, propose creating a **new permanent script** to be maintained and used systematically — justify why it's needed. Never create throwaway or temporary scripts.
3. **Wait for user approval** before making any change or creating any new file.

### Pipeline Scripts (per-email processing)

| File | Purpose |
|------|---------|
| `pipeline_init.py` | Initialize session: create XLSX + HTML templates. Usage: `python pipeline_init.py --type blog|video --from-date YYYY-MM-DD --to-date YYYY-MM-DD`. Naming: `{Type}_Notifications-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.{ext}` (adds `-NN` suffix on collision). |
| `pipeline_retrieve.py` | Retrieve matching emails from Outlook Web via Playwright CDP for a date range. |
| `pipeline_fetch_blog.py` | Fetch blog content from a URL: resolve final URL, extract title, date, article text. |
| `pipeline_check_dup.py` | Check if a blog post is a duplicate (session + SP). |
| `pipeline_sp_create.py` | Create a new BlogPosts SP item via REST API (POST + MERGE for Link). |
| `pipeline_email_actions.py` | Categorize and/or move a single email in Outlook Web (safe search, Italian labels). |
| `pipeline_update_reports.py` | Rebuild XLSX + HTML reports from `session_state.json`. |
| `pipeline_sweep_inbox.py` | Sweep all unprocessed blog emails from Inbox: scroll search results, categorize + move each, loop until zero remain. |
| `pipeline_batch.py` | Batch processor for blog notification emails: groups by title, performs bulk operations (email_actions, sp_create, fetch, dupcheck) in a single CDP session. |

### Video Pipeline Scripts (per-email processing)

| File | Purpose |
|------|----------|
| `pipeline_video_retrieve.py` | Retrieve matching video notification emails from Outlook Web via Playwright CDP for a date range. |
| `pipeline_fetch_video.py` | Fetch YouTube video metadata via CDP: title, published date, duration, description, chapters. |
| `pipeline_video_check_dup.py` | Check if a video post is a duplicate (session + SP), matches by title and yt_id. |
| `pipeline_video_sp_create.py` | Create a new VideoPosts SP item via REST API (POST + MERGE for Link). |
| `pipeline_video_email_actions.py` | Categorize and/or move a single video email in Outlook Web (safe search, Italian labels). |
| `pipeline_video_email_report.py` | Build HTML video digest from SP VideoPosts (grouped by topic), save to `output/` for email sending. |

### Teams Meeting Recording Pipeline Scripts

| File | Purpose |
|------|---------|
| `pipeline_fetch_teams_meeting.py` | Fetch metadata + transcript from a SharePoint Stream page via CDP. Extracts title, date, duration, transcript; computes SHA256 of URL. |
| `pipeline_teams_sp_create.py` | Create a new VideosMSInt SP item via REST API (POST + MERGE for Link). Includes built-in dedup check by SHA256. |

### VE Notification Pipeline Scripts

| File | Purpose |
|------|---------|
| `ve-notifications-retrieve.py` | Retrieve VE notification emails from Outlook Web via CDP for a date range. |
| `ve-notifications-analyze.py` | Extract detailed content from a specific VE notification email in search results. |
| `ve-notifications-process.py` | Open a VE thread URL via CDP and read all replies/comments (expand all). |
| `ve-notifications-email-actions.py` | Search, Ctrl+A select all, categorize and move VE notification emails in Outlook. |
| `ve-notifications-build-html.py` | Build HTML digest from VE notification summaries JSON → `output/`. |

### Supporting Scripts

| File | Purpose |
|------|---------|
| `pipeline_fetch_blogposts.py` | Fetches all existing BlogPosts from SP list → `sp_blogposts.json`. |
| `pipeline_fetch_sp_list.py` | Fetches "Ref Technologies New" list from SP → `tech_list.json`. |
| `pipeline_cache_blogs.py` | Bulk blog content fetcher/cacher → `blog_cache/`. |
| `pipeline_update_sp_summaries.py` | Bulk-updates Summary field on existing SP items via REST MERGE. |
| `pipeline_email_report.py` | Build HTML blog digest from SP BlogPosts (grouped by topic), save to `output/` for email sending. |
| `engage_read_conversations.py` | Read conversations from a Viva Engage community via CDP → JSON to stdout. |
| `engage_build_html.py` | Build HTML digest from Viva Engage conversation summaries JSON → `output/`. |
| `pipeline_ve_email_actions.py` | Categorize and move VE notification emails in Outlook (search by subject, categorize as "By agent - Viva Engage", move to "Social Networks"). |
| `yt_transcript.py` | Download YouTube video transcript via CDP → `yt_<VIDEO_ID>.txt`. |
| `pipeline_fetch_videoposts.py` | Fetches all existing VideoPosts from SP list → `sp_videoposts.json`. |
| `cdp_helper.py` | Shared helper: checks if Edge CDP is reachable, auto-launches Edge with debug profile if not. Imported by all CDP-dependent scripts. |

### Configuration & Data Files

| File / Directory | Purpose |
|------------------|---------|
| `config.json` | Central configuration: email sender, subject prefix, SP list URLs, Outlook category/folder, topic→tech lookup maps. **All scripts read this at startup.** |
| `session_state.json` | Current session state: list of processed emails with status, SP item IDs, errors. Created by `pipeline_init.py`, updated by pipeline steps. |
| `sp_blogposts.json` | Cached snapshot of all existing SP BlogPosts items. Generated by `pipeline_fetch_blogposts.py`, used for deduplication. |
| `sp_videoposts.json` | Cached snapshot of all existing SP VideoPosts items. Generated by `pipeline_fetch_videoposts.py`, used for deduplication. |
| `tech_list.json` | Cached snapshot of the "Ref Technologies New" SP list. Generated by `pipeline_fetch_sp_list.py`, used for topic→tech lookups. |
| `blog_cache/` | Directory of cached blog article texts (`.txt` files). Populated by `pipeline_fetch_blog.py` and `pipeline_cache_blogs.py`. |
| `yt_transcripts/` | Directory of downloaded YouTube video transcripts (`yt_<VIDEO_ID>.txt`). Populated by `yt_transcript.py`. |
| `teams_transcripts/` | Directory of downloaded Teams meeting recording transcripts (`<SHA256>.txt`). Populated by `pipeline_fetch_teams_meeting.py`. |
| `config.json.template` | Template for `config.json` with placeholder values. Copy and fill in to create a working `config.json`. |
| `ve_notifications_cache.json` | Cache of VE notification emails being processed by the `vivaengage-notifications` skill. Overwritten on each new run. |
| `output/` | Directory for generated reports. Naming convention: `Blog_Notifications-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.{ext}` (blog), `Video_Notifications-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.{ext}` (video), `Viva_Engage-Digest-From-YYYY.MM.DD-To-YYYY.MM.DD.html` (Viva Engage). If file exists, `-NN` suffix is added before the extension. |

## Prerequisites

1. **Microsoft Edge** installed at `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
2. **Python 3.12+** available in PATH
3. **Playwright for Python**: install with `pip install playwright` (no need to run `playwright install` — we connect to Edge, not a bundled browser)

## Setup Procedure

> **Why a copy?** Chrome/Edge 136+ blocks `--remote-debugging-port` on the default user data directory for security. A copy preserves all cookies, sessions, and corporate SSO tokens.

### Step 1 — Close Edge and Launch with Debug Profile

The debug profile at `$env:LOCALAPPDATA\Edge-Debug-Profile` **persists across reboots**. Most of the time you can just launch Edge directly with it — no copy needed.

```powershell
Stop-Process -Name msedge -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" `
    --remote-debugging-port=9222 `
    --user-data-dir="$env:LOCALAPPDATA\Edge-Debug-Profile" `
    --profile-directory="<value from config.json → edge_cdp.profile_name>" `
    --no-first-run `
    --no-default-browser-check
```

Verify at `edge://version` that the profile is correct and try accessing an authenticated page (e.g. Outlook). If authentication works → **skip Step 2 entirely**.

> **Important:** Do NOT have another Edge instance using the same `--user-data-dir` running simultaneously.

### Step 2 — Refresh Profile (only if auth expired or first-time setup)

Only needed when SSO tokens have expired (you get login prompts on corporate pages) or the debug profile doesn't exist yet.

1. Re-authenticate in regular Edge first (open Outlook, SharePoint, etc.)
2. Close all Edge windows
3. Run an **incremental sync** (copies only changed files — takes seconds, not minutes):

```powershell
Stop-Process -Name msedge -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$src = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
$dst = "$env:LOCALAPPDATA\Edge-Debug-Profile"
robocopy $src $dst /MIR /XO /XD "Crashpad" "ShaderCache" "GrShaderCache" "Service Worker" "Code Cache" "GPUCache" /XF "lockfile" "LOCK" /NFL /NDL /NJH /NJS /NP
```

> `/XO` skips files that haven't changed → incremental sync in seconds. Only the first-ever copy is slow.

Then re-launch Edge with Step 1.

### Step 3 — Verify CDP Connectivity (optional)

```powershell
Invoke-RestMethod -Uri "http://localhost:9222/json/version" | ConvertTo-Json
```

Should return browser metadata with a `webSocketDebuggerUrl`.

## How Copilot Should Automate Edge

### Connecting to Edge

Use `run_in_terminal` to execute Python scripts that connect to the running Edge instance:

```python
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    context = browser.contexts[0]  # reuse the existing authenticated context
    page = context.new_page()      # or context.pages[0] for an existing tab
    page.goto("https://outlook.office.com/mail/")
    # ... automate as needed
```

### Key Patterns

#### Navigate to a URL
```python
page.goto("https://outlook.office.com/mail/")
page.wait_for_load_state("networkidle")
```

#### Read page content
```python
content = page.content()          # full HTML
text = page.inner_text("body")    # visible text
title = page.title()
```

#### Click elements
```python
page.click("text=Inbox")
page.click("[aria-label='Search']")
page.click("button:has-text('Categorize')")
```

#### Type text
```python
page.fill("input[aria-label='Search']", "search query here")
page.press("input[aria-label='Search']", "Enter")
```

#### Wait for elements
```python
page.wait_for_selector("div[role='listbox']", timeout=10000)
```

#### Take a screenshot (for debugging)
```python
page.screenshot(path="debug-screenshot.png")
```

#### Extract links
```python
links = page.eval_on_selector_all("a[href]", "els => els.map(e => ({text: e.innerText, href: e.href}))")
```

### Running Multi-Step Scripts

Copilot should compose Python scripts and execute them via `run_in_terminal`. For complex workflows:

1. Write the script to a `.py` file in the workspace using `create_file`
2. Execute with `python <script>.py`
3. Have the script output structured data (JSON) to stdout or to a file
4. Read the output and process it

### Important Notes

- **Do NOT use VS Code integrated browser tools** (`open_browser_page`, `read_page`, `click_element`, etc.) for authenticated corporate pages — they use Chromium, which is blocked by Conditional Access.
- **Do use `run_in_terminal`** with Python + Playwright `connect_over_cdp` for all authenticated page interactions.
- **`fetch_webpage`** can still be used for public pages (e.g., public blog posts) that don't require authentication.
- **Session freshness:** If SSO tokens expire, re-authenticate in regular Edge, then run the incremental profile sync (Step 2) and re-launch. The debug profile persists across reboots — no need to re-copy if auth is still valid.
- **Edge must stay open** during automation — do not close the Edge window launched in Step 1.
- **No input focus required** — Playwright controls Edge via CDP regardless of which window has focus.

## Quick Reference — Session Startup Checklist

```
[ ] Close all Edge windows
[ ] Launch Edge with debug profile (Step 1)
[ ] Check if authenticated pages load (Outlook, SharePoint)
[ ] If auth failed → re-login in regular Edge, incremental sync (Step 2), relaunch
[ ] Verify CDP: Invoke-RestMethod http://localhost:9222/json/version
[ ] Ready for Copilot automation via Playwright CDP
```

## Workspace Configuration

This workspace uses a `config.json` file at the root for all configurable parameters (email sender, subject prefix, SP list URLs, Outlook category/folder, lookup maps). **Always read `config.json` at the start of any automation task.** Never hard-code these values in scripts or instructions.

## Outlook Web Automation — Safety Rules

When automating Outlook Web via Playwright, these rules are **mandatory**:

1. **Every search MUST include both `from:{sender}` AND `subject:({prefix})`** from `config.json`. Never run a broad search. Never read or act on emails not returned by a properly filtered search.

2. **Never use Ctrl+A to select all search results.** Results may include unrelated emails. Instead, iterate `div[role='option']` rows, check each for the subject prefix (e.g. `[Blog-`) in its text, and Ctrl+click only verified matches.

3. **The "Sposta" (Move) button** exists in multiple places. Always select the one in the top toolbar (bounding box `y < 200`). Use keyboard navigation (ArrowDown + Enter) in the folder search, not direct folder clicks.

4. **ElementHandle.click may timeout.** Use `timeout=5000` and fall back to a bounding-box click via `page.mouse.click(x, y)`.

5. **Menu items use Italian labels** in this environment: "Categorizza", "Sposta", "Cerca una cartella".

6. **Never use `-category:` in Outlook Web search queries.** Outlook Web search does not support negative category filters — they silently fail and return zero results. To exclude already-processed emails, omit the category filter from the search query and instead **visually inspect** each email's category labels in the reading pane after clicking on it. If the email already has the processed category, skip it.

## SharePoint REST API — Key Patterns

When creating or updating items in SP lists via the REST API:

1. **Digest token**: Obtain via `_api/contextinfo` POST. Refresh every ~40 operations.
2. **Link field**: Cannot be set in the initial POST. Use a separate MERGE request with `SP.FieldUrlValue`.
3. **Lookup fields**: `SourceNewId` (single int), `TechId` (`{"results": [int, ...]}`).
4. **Entity type**: Read from `config.json` (`sharepoint.blog_list_entity_type`).
5. **All requests** use `odata=verbose` in Accept/Content-Type headers.
