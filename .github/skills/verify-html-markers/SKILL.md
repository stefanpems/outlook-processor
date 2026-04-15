---
name: verify-html-markers
description: "Verify that HTML digest files contain the three structural markers required by the external automation. Run on demand to check source code and/or generated output. Triggered by prompts like 'verifica i marker HTML' or 'check HTML digest markers' or 'controlla la struttura degli HTML'."
argument-hint: "Optionally specify --source or --output to limit the check scope."
---

# Verify HTML Digest Markers

## Purpose

An external automation injects content into the HTML digest files produced by this workspace. It locates injection points via three **exact text markers** that MUST appear exactly once in every digest HTML:

| Marker | Exact text | Location |
|--------|-----------|----------|
| **M1** | `<p style="font-size:15px;color:#555;margin:-12px 0 20px 0;">` | Subtitle paragraph, right before the "From... To..." date label |
| **M2** | `</p>` + newline + `<div class="stats-bar">` | End of subtitle `</p>` followed by a literal `\n` then the stats bar `<div>` |
| **M3** | `<div class="footer-bar">` | Footer bar at the bottom of the page |

These strings are searched **as-is** (literal text match). Any change to whitespace, attributes, ordering, or duplication will break the automation.

## Generator Files Under Contract

The following four Python scripts produce HTML digest files that MUST respect this contract:

| Script | Digest type |
|--------|-------------|
| `pipeline_email_report.py` | Blog digest |
| `pipeline_video_email_report.py` | Video digest |
| `engage_build_html.py` | Viva Engage conversational digest |
| `ve-notifications-build-html.py` | Viva Engage notification digest |

**Note:** `pipeline_update_reports.py` produces session-report HTML with a different template — it is NOT subject to this contract.

## Verification Script

`verify_html_markers.py` in the workspace root performs two checks:

1. **SOURCE CHECK** — scans each generator Python file, verifies that M1, M2, M3 strings exist exactly once, the list is joined with `'\n'`, and warns if any marker is inside a conditional block.
2. **OUTPUT CHECK** — scans all HTML files in `output/` matching digest naming patterns, verifies the three markers each appear exactly once. Session-report files (no markers at all) are automatically skipped.

## Usage

```bash
python verify_html_markers.py              # both checks
python verify_html_markers.py --source     # source code only
python verify_html_markers.py --output     # output files only
```

Exit code: `0` = all checks pass, `1` = at least one failure.

## When To Run

- **After any edit** to the four generator scripts listed above
- **Before sending a digest email** to confirm the HTML is well-formed
- **On demand** via user prompt (e.g. "verifica i marker HTML")

## Known Issue

In `ve-notifications-build-html.py`, the M1 `<p>` tag is wrapped in `if subtitle:`. When `date_from` and `date_to` are not provided, `subtitle` is empty and M1+M2 are absent. In normal pipeline usage dates are always provided, but this is a fragile spot.
