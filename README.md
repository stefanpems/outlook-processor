# Outlook Processor — Skills & Scripts Reference

This workspace contains a set of Copilot skills that automate the processing of email notifications, blog posts, video metadata, and Viva Engage conversations. All skills connect to Microsoft Edge via Chrome DevTools Protocol (CDP) using Playwright, and interact with Outlook Web, SharePoint REST APIs, and YouTube pages.

All configurable parameters (senders, subject prefixes, categories, folders, SP endpoints, lookup maps) are centralized in `config.json`. Skills and scripts never hardcode these values.

---

## Skills

### 1. blog-notifications

**What it does:**
Processes blog notification emails from Outlook for a given date range. For each email: retrieves matching emails from Outlook Web via CDP, fetches the linked blog article content (resolving redirects, extracting title and publication date), checks for duplicates both within the current session and against the SharePoint BlogPosts list, generates an English HTML summary (100–150 words) using the Copilot LLM, classifies the article's technology using the configured tech map, creates or updates the corresponding SP list item, assigns the configured Outlook category to the email, moves it to the configured target folder, and updates synchronized XLSX + HTML session reports after each email. Optionally sends an HTML digest email at the end.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on Outlook Web and SharePoint.
- Emails present in the mailbox from the configured sender, with the configured subject prefix (e.g. `[Blog-Topic]`), not yet bearing the configured processed category (unless reprocess mode is used).
- The SharePoint BlogPosts list accessible via the configured REST API endpoint.

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Processa le notifiche blog dal X al Y"` | **Standard mode** — retrieves only unprocessed emails, creates/updates SP items as needed, categorizes and moves emails. |
| `"Riprocessa le notifiche blog dal X al Y"` | **Reprocess mode** — includes already-processed emails, always regenerates summaries and updates SP items. Activated by "Ri-" (IT) or "Re-" (EN) prefix on the action verb. |
| `"Crea digest dei blog dal X al Y"` | **Digest mode** — retrieves all emails (including processed), fills gaps in SP items, skips complete items, then generates and sends the HTML digest email. |
| `"Crea digest riprocessando i blog dal X al Y"` | **Digest + Reprocess** — retrieves all, regenerates all summaries, sends digest. |

Dates can be in any format (`YYYY-MM-DD`, `YYYY.MM.DD`, `DD/MM/YYYY`, natural language). If missing, the skill asks for them.

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `pipeline_init.py` | Initialize session: create XLSX + HTML report templates |
| `pipeline_retrieve.py` | Search and retrieve matching blog emails from Outlook Web |
| `pipeline_fetch_blog.py` | Fetch blog article content, resolve final URL, extract metadata |
| `pipeline_check_dup.py` | Check for session and SP duplicates |
| `pipeline_sp_create.py` | Create or update BlogPosts SP item via REST API |
| `pipeline_email_actions.py` | Categorize and move each email in Outlook |
| `pipeline_update_reports.py` | Rebuild XLSX + HTML reports from session state |
| `pipeline_fetch_blogposts.py` | Download all existing SP BlogPosts for deduplication |
| `pipeline_sweep_inbox.py` | Bulk sweep of unprocessed blog emails (categorize + move) |
| `pipeline_email_report.py` | Generate the HTML blog digest for email sending |

---

### 2. video-notifications

**What it does:**
Processes video notification emails from Outlook for a given date range. For each email: retrieves matching emails from Outlook Web, fetches YouTube video metadata (title, duration, published date, description, chapters) by opening the video page via CDP, downloads the video transcript to a local file, generates an English HTML abstract using the Copilot LLM (including a clickable chapter list with timestamps), checks for duplicates within the session and against the SharePoint VideoPosts list, creates or updates the SP item, assigns the configured Outlook category, moves the email to the configured folder, and updates session reports. Emails that are hashtag-only social-reach notifications (detected by topic and `#` in subject) are skipped for SP processing but still categorized and moved. Optionally sends an HTML digest email.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on Outlook Web, SharePoint, and YouTube.
- Emails present in the mailbox from the configured video sender, with the configured video subject prefix (e.g. `[Video-Topic]`), not yet bearing the configured video processed category (unless reprocess mode is used).
- The SharePoint VideoPosts list accessible via the configured REST API endpoint.

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Processa le notifiche video dal X al Y"` | **Standard mode** — unprocessed emails only, create/update SP items, categorize and move. |
| `"Riprocessa le notifiche video dal X al Y"` | **Reprocess mode** — includes processed emails, regenerates all abstracts. |
| `"Crea digest dei video dal X al Y"` | **Digest mode** — all emails, fill SP gaps, send digest. |
| `"Crea digest dei video riprocessando dal X al Y"` | **Digest + Reprocess** — all emails, regenerate all, send digest. |

Reprocess is activated by "Ri-"/"Re-" prefix verbs. Dates in any format.

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `pipeline_init.py` | Initialize session (with `--type video`) |
| `pipeline_video_retrieve.py` | Search and retrieve matching video emails from Outlook Web |
| `pipeline_fetch_video.py` | Fetch YouTube video metadata via CDP |
| `yt_transcript.py` | Download YouTube video transcript via CDP |
| `pipeline_video_check_dup.py` | Check for session and SP duplicates (by title and yt_id) |
| `pipeline_video_sp_create.py` | Create or update VideoPosts SP item via REST API |
| `pipeline_video_email_actions.py` | Categorize and move each video email in Outlook |
| `pipeline_update_reports.py` | Rebuild XLSX + HTML reports from session state |
| `pipeline_fetch_videoposts.py` | Download all existing SP VideoPosts for deduplication |
| `pipeline_video_email_report.py` | Generate the HTML video digest for email sending |

---

### 3. blog-email-report

**What it does:**
Queries the SharePoint BlogPosts list for items in a date range, builds an HTML digest grouped by topic, and sends it via email. Each topic section lists articles sorted by publication date (descending) then title (ascending), showing clickable title, date, technologies, and summary. Does not process emails — it only reads from SP and generates the report.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on SharePoint.
- BlogPosts SP items existing in the list for the requested date range.

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Invia il digest dei blog"` | Uses yesterday as default date range, sends to configured default recipients. |
| `"Invia il digest dei blog dal X al Y a user@example.com"` | Uses specified date range and recipients. |
| `"Invia il digest dei blog filtrato per Sentinel,Entra"` | Filters SP items by the specified technology labels (must match keys in the configured tech map). |

If no items match the query, the skill informs the user and does not send the email.

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `pipeline_email_report.py` | Fetch SP BlogPosts, build HTML digest, save to `output/` |

---

### 4. register-blog-post

**What it does:**
Registers one or more blog articles into the SharePoint BlogPosts list directly from their URLs, independent of the email pipeline. For each URL: fetches the page content (resolving redirects), extracts title and publication date, generates an English HTML summary via the Copilot LLM, classifies the technology using the configured tech map, checks for duplicates against the SP list, and creates or updates the SP item. If an article is already registered with a summary, the skill asks whether to regenerate.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on SharePoint.
- The blog URLs must be publicly accessible (or accessible within the authenticated Edge session).

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Registra https://example.com/post"` | Registers a single blog post URL. |
| `"Registra https://url1 e https://url2"` | Registers multiple URLs in sequence. |

The skill always confirms the URL list before starting. The topic field is always set to the configured default value (the `Other` key in the source map).

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `pipeline_fetch_blog.py` | Fetch blog content, resolve final URL, extract metadata |
| `pipeline_check_dup.py` | Check for SP duplicates |
| `pipeline_sp_create.py` | Create or update BlogPosts SP item |
| `pipeline_fetch_blogposts.py` | Download all existing SP BlogPosts for deduplication |

---

### 5. vivaengage-conversations

**What it does:**
Reads and summarizes recent conversations from one or more Viva Engage communities. For each community: navigates to the community feed via CDP (sorted by recent activity), scrolls the virtualized feed reading each conversation (post + all expanded replies), stops when all conversation dates are older than the requested time window. Generates structured English summaries per conversation type (question, announcement, discussion) with formatted Q&A, comments, author attribution, follow-up actions. Builds an HTML report saved to `output/`. Optionally categorizes and moves related Viva Engage notification emails in Outlook. In digest mode, also sends the HTML report by email.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on Viva Engage.
- The configured community names must match the sidebar link text in Viva Engage exactly.

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Riassumi le conversazioni Viva Engage degli ultimi 3 giorni"` | **Mode 1 (HTML file)** — reads all configured communities, summarizes, saves HTML report. |
| `"Novità dalla community Defender for Cloud"` | Mode 1 scoped to a single named community. |
| `"Crea digest Viva Engage degli ultimi 5 giorni"` | **Mode 2 (email digest)** — same as Mode 1 but also sends the HTML as email to configured recipients. |

If a specific community is named, only that community is processed. If no community is named, all communities from the configuration are processed. If no number of days is specified, the configured default is used.

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `engage_read_conversations.py` | Read conversations from a single community via CDP |
| `engage_build_html.py` | Build HTML report from conversation summaries JSON |
| `pipeline_ve_email_actions.py` | Categorize and move VE notification emails in Outlook |

---

### 6. vivaengage-notifications

**What it does:**
Creates a digest from Viva Engage notification emails in Outlook. For each notification email in the date range: retrieves emails from Outlook Web via CDP, reads the email body to extract post type, title, community, author, and thread URL, opens the Viva Engage thread via CDP to read all replies and comments (expanding collapsed content), generates structured English summaries (with Q&A format for questions, comments format for announcements/discussions), categorizes and moves each email in Outlook, then builds an HTML digest grouped by community and sends it by email.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on Outlook Web and Viva Engage.
- Viva Engage notification emails present in the mailbox from the configured notification sender, not yet bearing the configured VE processed category (unless reprocess mode is used).
- Community names in the emails must match the configured community list (emails from other communities are skipped).

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Crea digest per le notifiche Viva Engage dal X al Y"` | Retrieves unprocessed VE notification emails, processes threads, builds and sends digest. |
| `"Crea digest per le notifiche Viva Engage di ieri"` | Uses yesterday as both start and end date. |
| `"Riprocessa le notifiche Viva Engage dal X al Y"` | **Reprocess mode** — includes already-processed emails. |

If no dates are specified, defaults to yesterday. Reprocess activated by "Ri-"/"Re-" prefix verbs.

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `ve-notifications-retrieve.py` | Retrieve VE notification emails from Outlook Web via CDP |
| `ve-notifications-analyze.py` | Extract detailed content from a specific email (fallback tool) |
| `ve-notifications-process.py` | Open a VE thread URL and read all replies/comments via CDP |
| `ve-notifications-email-actions.py` | Search, select, categorize and move VE notification emails |
| `ve-notifications-build-html.py` | Build HTML digest from notification summaries JSON |

---

### 7. combined-digest

**What it does:**
Orchestrates the blog, video, and Viva Engage notification digests in sequence for a given time window. First runs the blog-notifications skill in Digest mode (ensuring all blog notification emails are processed and SP items are complete), then generates and sends the blog HTML digest. Next, runs the video-notifications skill in Digest mode (same logic for video emails and SP VideoPosts items), generating and sending the video HTML digest. Finally, runs the vivaengage-notifications skill (retrieving VE notification emails, reading thread replies, summarizing, categorizing and moving emails), generating and sending the VE HTML digest. Each digest is sent independently to its own configured recipients.

**Prerequisites:**
- All prerequisites of `blog-notifications`, `video-notifications`, and `vivaengage-notifications` skills combined.

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Crea un digest degli ultimi 3 giorni"` | Runs blog + video + VE notification digests for the last 3 days (ending yesterday). |
| `"Digest settimanale"` | Last 7 days ending yesterday. |
| `"Crea un digest dell'ultimo giorno"` | Yesterday only. |

This skill activates only when the user asks for a generic digest **without** specifying "blog", "video", or "Viva Engage". If a specific type is mentioned, the corresponding individual skill is used instead. The date range ends at yesterday; "last X days" means X days inclusive ending yesterday.

**Python scripts used:**
All scripts from `blog-notifications`, `video-notifications`, and `vivaengage-notifications` (see those skills for details).

---

### 8. teams-meeting-recording

**What it does:**
Registers a Teams meeting recording in the SharePoint VideosMSInt list from its SharePoint Stream URL. Connects to Edge via CDP, navigates to the Stream page with all media muted, extracts metadata (title, published date, duration) from the page DOM and video element, opens the transcript panel and scrapes all transcript entries by scrolling the virtualized list, computes the SHA256 hex digest of the URL as a dedup key, saves the transcript to `teams_transcripts/<sha256>.txt`. The Copilot agent then generates a 100–200 word English HTML summary from the transcript and classifies the technology using the configured tech map. Finally, creates the SP item via REST API (POST + MERGE for Link field), with built-in duplicate detection by SHA256.

**Prerequisites:**
- Edge running with CDP debug profile, user authenticated on SharePoint.
- The recording must be hosted on SharePoint Stream and accessible in the authenticated Edge session.
- The transcript panel must be visible on the Stream page (the script scrolls the virtualized list to collect all entries).

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Registra questo meeting recording https://..."` | Registers a single Teams meeting recording from its Stream URL. |
| `"Registra queste registrazioni Teams: https://url1 e https://url2"` | Registers multiple recordings in sequence. |

The skill always confirms the URL list before starting. For each URL, the full pipeline runs in order: fetch metadata + transcript → generate summary → classify tech → create SP item.

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `pipeline_fetch_teams_meeting.py` | Fetch metadata + transcript from a SharePoint Stream page via CDP |
| `pipeline_teams_sp_create.py` | Create a new VideosMSInt SP item via REST API (with built-in dedup check by SHA256) |

---

### 9. youtube-transcript-downloader

**What it does:**
Downloads the transcript (auto-generated or manual subtitles) of a YouTube video and saves it as a text file. Connects to Edge via CDP, navigates to the YouTube video page, opens the transcript panel, scrapes all transcript segments from the DOM, and saves the text to `yt_transcripts/yt_<VIDEO_ID>.txt`. This approach avoids YouTube API rate limiting (429 errors) because the transcript is loaded by YouTube's own frontend.

**Prerequisites:**
- Edge running with CDP debug profile.
- The YouTube video must have a transcript/subtitles available.

**Prompt structure and behavior variants:**

| Prompt pattern | Behavior |
|----------------|----------|
| `"Scarica il transcript del video https://youtube.com/watch?v=ABC123"` | Downloads transcript with timestamps. |
| `"Scarica il transcript pulito del video ABC123"` | Uses `--clean` flag: strips timestamps, outputs only spoken text. Accepts video ID or full URL. |

**Python scripts used:**

| Script | Role in this skill |
|--------|--------------------|
| `yt_transcript.py` | Download transcript from YouTube video page via CDP |

---

## Python Scripts Reference

The table below lists all Python scripts in the workspace, their purpose, and which instruction files reference them.

| Script | Purpose | Referenced in |
|--------|---------|---------------|
| `cdp_helper.py` | Shared helper: checks if Edge CDP is reachable, auto-launches Edge with debug profile if not. Imported by all CDP-dependent scripts. | `copilot-instructions.md` |
| `pipeline_init.py` | Initialize a processing session: create XLSX + HTML report templates and `session_state.json`. Accepts `--type blog\|video` and date range. | `copilot-instructions.md`, `blog-notifications`, `video-notifications` |
| `pipeline_retrieve.py` | Retrieve matching blog notification emails from Outlook Web via CDP for a date range. Supports `--include-processed` flag. | `copilot-instructions.md`, `blog-notifications` |
| `pipeline_fetch_blog.py` | Fetch blog content from a URL: resolve final URL via redirects, extract title, publication date, and article text (up to 8000 chars). | `copilot-instructions.md`, `blog-notifications`, `register-blog-post` |
| `pipeline_check_dup.py` | Check if a blog post is a duplicate within the current session and against the SP BlogPosts list. Returns `dup_session`, `dup_sp`, `sp_id`, `sp_has_summary`. | `copilot-instructions.md`, `blog-notifications`, `register-blog-post` |
| `pipeline_sp_create.py` | Create a new BlogPosts SP item via REST API (POST + MERGE for Link field), or update summary on an existing item (`--update-summary`). | `copilot-instructions.md`, `blog-notifications`, `register-blog-post` |
| `pipeline_email_actions.py` | Categorize and/or move a single blog email in Outlook Web. Uses safe search, Italian UI labels, keyboard-based folder selection. | `copilot-instructions.md`, `blog-notifications` |
| `pipeline_update_reports.py` | Rebuild XLSX + HTML session reports from `session_state.json`. | `copilot-instructions.md`, `blog-notifications`, `video-notifications` |
| `pipeline_sweep_inbox.py` | Sweep all unprocessed blog emails from Inbox: scroll search results, categorize + move each, loop until zero remain. | `copilot-instructions.md`, `blog-notifications` |
| `pipeline_batch.py` | Batch processor for blog notification emails: groups by title, performs bulk operations in a single CDP session. | `copilot-instructions.md` |
| `pipeline_fetch_blogposts.py` | Fetch all existing BlogPosts from SP list → `sp_blogposts.json`. Used for deduplication. | `copilot-instructions.md`, `blog-notifications`, `register-blog-post` |
| `pipeline_fetch_sp_list.py` | Fetch "Ref Technologies New" list from SP → `tech_list.json`. | `copilot-instructions.md` |
| `pipeline_cache_blogs.py` | Bulk blog content fetcher/cacher → `blog_cache/`. | `copilot-instructions.md`, `blog-notifications` |
| `pipeline_update_sp_summaries.py` | Bulk-update Summary field on existing SP BlogPosts items via REST MERGE. | `copilot-instructions.md`, `blog-notifications` |
| `pipeline_email_report.py` | Build HTML blog digest from SP BlogPosts (grouped by topic), save to `output/`. | `copilot-instructions.md`, `blog-notifications`, `blog-email-report` |
| `pipeline_video_retrieve.py` | Retrieve matching video notification emails from Outlook Web via CDP for a date range. Supports `--include-processed` flag. | `copilot-instructions.md`, `video-notifications` |
| `pipeline_fetch_video.py` | Fetch YouTube video metadata via CDP: title, published date, duration, description, chapters. | `copilot-instructions.md`, `video-notifications` |
| `pipeline_video_check_dup.py` | Check if a video post is duplicate (session + SP), matches by title and yt_id. | `copilot-instructions.md`, `video-notifications` |
| `pipeline_video_sp_create.py` | Create a new VideoPosts SP item via REST API, or update abstract on existing item (`--update-abstract`). | `copilot-instructions.md`, `video-notifications` |
| `pipeline_video_email_actions.py` | Categorize and/or move a single video email in Outlook Web. | `copilot-instructions.md`, `video-notifications` |
| `pipeline_video_email_report.py` | Build HTML video digest from SP VideoPosts (grouped by topic), save to `output/`. | `copilot-instructions.md`, `video-notifications` |
| `pipeline_fetch_videoposts.py` | Fetch all existing VideoPosts from SP list → `sp_videoposts.json`. Used for deduplication. | `copilot-instructions.md`, `video-notifications` |
| `yt_transcript.py` | Download YouTube video transcript via CDP → `yt_transcripts/yt_<VIDEO_ID>.txt`. Supports `--clean` flag. | `copilot-instructions.md`, `video-notifications`, `youtube-transcript-downloader` |
| `engage_read_conversations.py` | Read conversations from a single Viva Engage community via CDP. Outputs JSON to stdout. | `copilot-instructions.md`, `vivaengage-conversations`, `combined-digest` |
| `engage_build_html.py` | Build HTML digest from Viva Engage conversation summaries JSON → `output/`. | `copilot-instructions.md`, `vivaengage-conversations`, `combined-digest` |
| `pipeline_ve_email_actions.py` | Categorize and move VE notification emails in Outlook (search by subject, set configured VE category, move to configured VE folder). | `copilot-instructions.md`, `vivaengage-conversations` |
| `ve-notifications-retrieve.py` | Retrieve VE notification emails from Outlook Web via CDP for a date range. Supports `--include-processed` flag. | `copilot-instructions.md`, `vivaengage-notifications` |
| `ve-notifications-analyze.py` | Extract detailed content from a specific VE notification email in search results (fallback tool). | `copilot-instructions.md`, `vivaengage-notifications` |
| `ve-notifications-process.py` | Open a VE thread URL via CDP and read all replies/comments (expands collapsed content). | `copilot-instructions.md`, `vivaengage-notifications` |
| `ve-notifications-email-actions.py` | Search, select all, categorize and move VE notification emails in Outlook. Supports `--batch-file` for bulk operations. | `copilot-instructions.md`, `vivaengage-notifications` |
| `ve-notifications-build-html.py` | Build HTML digest from VE notification summaries JSON → `output/`. | `copilot-instructions.md`, `vivaengage-notifications` |
| `pipeline_fetch_teams_meeting.py` | Fetch metadata + transcript from a SharePoint Stream page via CDP. Extracts title, date, duration, transcript; computes SHA256 of URL. Saves transcript to `teams_transcripts/<sha256>.txt`. | `copilot-instructions.md`, `teams-meeting-recording` |
| `pipeline_teams_sp_create.py` | Create a new VideosMSInt SP item via REST API (POST + MERGE for Link). Includes built-in dedup check by SHA256. | `copilot-instructions.md`, `teams-meeting-recording` |
