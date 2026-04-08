---
name: register-blog-post
description: "Register one or more blog posts in the SharePoint BlogPosts list from their URLs. Fetches content, generates summary, classifies tech, creates SP items. Triggered by prompts like 'registra questo blog post' or 'register this blog post' or 'aggiungi questo articolo alla lista blog'."
argument-hint: "Provide one or more blog post URLs, e.g.: 'registra https://example.com/blog-post-1 e https://example.com/blog-post-2'"
---

# Register Blog Post

## Purpose

Register one or more blog articles into the SharePoint BlogPosts list by providing their URLs directly. For each URL: fetch page content, extract metadata (title, publication date), generate a well-formatted summary, classify the technology, check for duplicates, and create the SP list item.

This skill is independent of the email notification pipeline — it works with any publicly accessible blog URL.

## Configuration

All parameters are in `config.json` at the workspace root. Read it at the start of every run. Key sections:

- `sharepoint.blog_list_api` — SP REST API base for the BlogPosts list
- `sharepoint.blog_list_entity_type` — entity type for POST/MERGE
- `sharepoint.site_base` — SP site base URL for REST calls
- `source_map` — topic name → SP SourceNew lookup ID (this skill always uses the `Other` key)
- `tech_map` — technology label → SP Tech lookup ID

## Pipeline Python Scripts Used

| Script | Purpose | CLI Usage |
|--------|---------|-----------|
| `pipeline_fetch_blog.py` | Fetch blog content, resolve final URL, extract metadata | `python pipeline_fetch_blog.py <url>` |
| `pipeline_check_dup.py` | Check duplicates against session + SP | `python pipeline_check_dup.py <title> <final_url>` |
| `pipeline_sp_create.py` | Create new SP BlogPosts item via REST API | `python pipeline_sp_create.py _sp_input.json` |
| `pipeline_fetch_blogposts.py` | Fetch all existing SP BlogPosts → `sp_blogposts.json` | `python pipeline_fetch_blogposts.py` |

## Input

The user provides one or more URLs. Extract all URLs from the user's message. Confirm the list before starting.

## Procedure

Execute the steps below **strictly in order**.

---

### Step 0 — Fetch Existing SP Items (for deduplication)

```bash
python pipeline_fetch_blogposts.py
```

Downloads all existing SP BlogPosts items to `sp_blogposts.json`. Needed for duplicate checks.

---

### Step 1 — Process Each URL

For EACH URL provided, perform Steps 1.1 through 1.5 in sequence.

#### 1.1 — Fetch Blog Content

```bash
python pipeline_fetch_blog.py "<url>"
```

Returns JSON with: `final_url`, `title`, `published_date`, `content`, `content_length`.

- `final_url` is the canonical URL after redirects.
- `published_date` is extracted from page metadata or URL (format: `YYYY-MM-DD`).
- `content` is the full article text (up to 8000 chars).

If the script fails (e.g. 403/timeout), use `fetch_webpage` as fallback for public pages.

**If `published_date` is empty:** Try to infer from the URL path or page content. If still unavailable, ask the user.

#### 1.2 — Check for Duplicates

```bash
python pipeline_check_dup.py "<title>" "<final_url>"
```

Returns JSON with: `dup_session`, `dup_sp`, `sp_id`, `sp_has_summary`.

| `dup_sp` | `sp_has_summary` | Action |
|----------|-----------------|--------|
| false | — | Proceed to 1.3 + 1.4 + 1.5 (create new SP item) |
| true | false | Proceed to 1.3 + 1.4, then **update** existing SP item summary (1.5) |
| true | true | **Skip** — inform user the article is already registered with summary. Ask if they want to regenerate. |

#### 1.3 — Generate Summary (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Read the `content` returned by Step 1.1 and generate a summary:

- **Language:** Always English, regardless of source language.
- **Length:** 100–150 words.
- **Format:** Use `<b>` for keywords/product names, `<i>` for key concepts/phrases, `<ul><li>` for bullet-point lists. **Strongly recommended structure:** a short introductory sentence followed by a `<ul>` bullet list of the main points/concepts.
- **Coverage:** The summary must mention **every** concept, feature, or technique explained in the article. Nothing should be silently omitted — at minimum, each concept must be cited by name even if not elaborated.
- **Skip preambles:** Ignore generic introductory paragraphs (context-setting, motivation, broad statements). Focus exclusively on substantive content.
- **Style:** Concrete, informative. Report only what the author communicates. No filler, no speculation, no "the article discusses…" framing.

#### 1.4 — Classify Technology (Copilot LLM Task)

**This step is performed by the Copilot agent, not by a Python script.**

Using `tech_map` from `config.json`, assign technology tags based on the article content:

- If a sub-technology matches (e.g. `Azure / AKV / MHSM`), do NOT also include the parent (e.g. `Azure`).
- Comma-separate multiple values if the article covers multiple technologies.
- Only use values that exist as keys in `tech_map`. If no match, use `** Other Tech **`.

#### 1.5 — Create or Update SP Item

**If creating a new SP item** (`dup_sp` was false):

```powershell
Set-Content -Path _sp_input.json -Value '<JSON>' -Encoding utf8
python pipeline_sp_create.py _sp_input.json
Remove-Item _sp_input.json
```

The JSON must contain:
- `title`: blog post title (from Step 1.1)
- `published_date`: in `YYYY-MM-DD` format (from Step 1.1)
- `summary`: generated HTML summary (from Step 1.3)
- `topic`: always `"Other"` — maps to SP SourceNew via `config.json → source_map`
- `tech`: comma-separated tech tags (from Step 1.4)
- `blog_link`: the `final_url` (from Step 1.1)

**If updating summary on existing SP item** (`dup_sp` true, `sp_has_summary` false):

```powershell
Set-Content -Path _sp_input.json -Value '{"summary":"...","title":"..."}' -Encoding utf8
python pipeline_sp_create.py --update-summary <sp_id> _sp_input.json
Remove-Item _sp_input.json
```

**IMPORTANT:** Do NOT pipe JSON via `echo` in PowerShell — non-ASCII characters will be corrupted. **Always write JSON to a temp file** with `Set-Content -Encoding utf8` and pass the file path.

---

### Step 2 — Final Summary

After all URLs are processed, print a summary table:

| URL | Title | Published | Tech | SP Action | SP ID |
|-----|-------|-----------|------|-----------|-------|

Include: total processed, duplicates found, SP items created, SP items updated, any errors.
