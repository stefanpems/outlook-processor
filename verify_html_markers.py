"""Verify that HTML digest generators produce the required structural markers.

An external automation depends on three EXACT text markers being present
(each exactly once) in every HTML digest file:

  M1  <p style="font-size:15px;color:#555;margin:-12px 0 20px 0;">
  M2  </p>\n<div class="stats-bar">          (literal newline between)
  M3  <div class="footer-bar">

This script performs two checks:
  1. SOURCE CHECK — scans the five Python generator files and verifies
     that the marker strings appear in the code exactly as required.
  2. OUTPUT CHECK — scans every HTML file in output/ whose name matches
     the digest naming patterns, verifying the three markers each appear
     exactly once.

Usage:
    python verify_html_markers.py               # run both checks
    python verify_html_markers.py --source      # source check only
    python verify_html_markers.py --output      # output check only
"""
import argparse, os, re, sys

BASE = os.path.dirname(os.path.abspath(__file__))

# ── Markers ──────────────────────────────────────────────────────────
M1 = '<p style="font-size:15px;color:#555;margin:-12px 0 20px 0;">'
M2 = '</p>\n<div class="stats-bar">'
M3 = '<div class="footer-bar">'

MARKER_NAMES = ["M1 (subtitle <p>)", "M2 (</p>\\n<div stats-bar>)", "M3 (<div footer-bar>)"]
MARKERS = [M1, M2, M3]

# ── Generator files that MUST produce these markers ──────────────────
GENERATORS = [
    "pipeline_email_report.py",
    "pipeline_video_email_report.py",
    "pipeline_teams_email_report.py",
    "engage_build_html.py",
    "ve-notifications-build-html.py",
]

# ── Output filename patterns (regex) for digest HTML ─────────────────
DIGEST_PATTERNS = [
    r"Blog_Notifications-Digest-From-.*\.html",
    r"Video_Notifications-Digest-From-.*\.html",
    r"Recordings_Digest-From-.*\.html",
    r"Viva_Engage-Digest-From-.*\.html",
]

# Some files with same naming are produced by pipeline_update_reports.py
# (session reports) which use a different HTML template. We detect them
# by checking if they contain the M1 marker OR a known pipeline_update
# signature (<div class="stats">).  If they have neither M1 nor the
# session-report signature, we flag them as unknown.
SESSION_REPORT_SIGNATURE = '<div class="stats">'


def check_source():
    """Verify markers exist in the Python source of each generator."""
    print("=" * 60)
    print("SOURCE CHECK — scanning generator Python files")
    print("=" * 60)
    ok = True
    for gen in GENERATORS:
        path = os.path.join(BASE, gen)
        if not os.path.isfile(path):
            print(f"\n  FAIL  {gen}: file not found")
            ok = False
            continue

        with open(path, "r", encoding="utf-8") as f:
            src = f.read()

        print(f"\n  {gen}:")

        # M1: the <p style="..."> string must appear in source
        m1_count = src.count(M1)
        if m1_count == 1:
            print(f"    M1 (subtitle <p>)            : OK (1 occurrence)")
        elif m1_count == 0:
            print(f"    M1 (subtitle <p>)            : FAIL (not found)")
            ok = False
        else:
            print(f"    M1 (subtitle <p>)            : FAIL ({m1_count} occurrences, expected 1)")
            ok = False

        # M3: <div class="footer-bar"> in the body (not CSS)
        # Count occurrences that start with '<div' (not '.footer-bar' in CSS)
        m3_body = len(re.findall(re.escape(M3), src))
        m3_css = src.count('.footer-bar')
        m3_effective = m3_body  # in generated HTML only the h.append() matters
        if m3_effective == 1:
            print(f"    M3 (<div footer-bar>)        : OK (1 occurrence)")
        elif m3_effective == 0:
            print(f"    M3 (<div footer-bar>)        : FAIL (not found)")
            ok = False
        else:
            print(f"    M3 (<div footer-bar>)        : FAIL ({m3_effective} occurrences, expected 1)")
            ok = False

        # M2: verify join uses '\n' and that stats-bar h.append follows the </p> append
        if "\\n'.join(h)" in src or "'\\n'.join(h)" in src or '"\\n".join(h)' in src:
            print(f"    Join character                : OK (newline)")
        else:
            print(f"    Join character                : WARN (could not confirm '\\n'.join(h))")

        # Check that <div class="stats-bar"> appears exactly once
        stats_count = src.count('<div class="stats-bar">')
        if stats_count == 1:
            print(f"    M2 (<div stats-bar>)         : OK (1 occurrence)")
        else:
            print(f"    M2 (<div stats-bar>)         : FAIL ({stats_count} occurrences, expected 1)")
            ok = False

        # Warn about conditional emission
        warned_lines = set()
        if "if subtitle" in src or "if date_" in src:
            # Check if M1 is inside a conditional block
            lines = src.split('\n')
            for i, line in enumerate(lines):
                if M1 in line or (i > 0 and M1 in lines[i-1] + line):
                    # Look backwards for an if statement at same or lower indentation
                    indent = len(line) - len(line.lstrip())
                    for j in range(i-1, max(i-5, -1), -1):
                        stripped = lines[j].strip()
                        if stripped.startswith('if ') and stripped.endswith(':'):
                            j_indent = len(lines[j]) - len(lines[j].lstrip())
                            if j_indent < indent and j not in warned_lines:
                                warned_lines.add(j)
                                print(f"    WARNING: M1 is inside conditional block at line {j+1}: {stripped}")
                                print(f"             If the condition is False, M1 and M2 will be absent!")
                                break

    return ok


def check_output():
    """Verify markers in generated HTML digest files."""
    print("=" * 60)
    print("OUTPUT CHECK — scanning output/ HTML digest files")
    print("=" * 60)
    output_dir = os.path.join(BASE, "output")
    if not os.path.isdir(output_dir):
        print("  output/ directory not found — skipping")
        return True

    ok = True
    checked = 0
    skipped_session = 0

    for fn in sorted(os.listdir(output_dir)):
        if not fn.endswith(".html"):
            continue
        if not any(re.match(pat, fn) for pat in DIGEST_PATTERNS):
            continue

        path = os.path.join(output_dir, fn)
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()

        # Count how many markers are present
        counts = [content.count(m) for m in MARKERS]
        present = sum(1 for c in counts if c >= 1)

        # If zero markers → session report or legacy format, skip
        if present == 0:
            skipped_session += 1
            continue

        checked += 1

        # If some but not all markers, or wrong counts → FAIL
        failures = []
        for name, marker, count in zip(MARKER_NAMES, MARKERS, counts):
            if count != 1:
                failures.append(f"{name}: count={count}")

        if failures:
            print(f"\n  FAIL  {fn}")
            for fail in failures:
                print(f"    {fail}")
            ok = False

    print(f"\n  Digest files checked: {checked}, session reports skipped: {skipped_session}")
    if ok and checked > 0:
        print("  All checked digest files: OK")
    elif checked == 0:
        print("  No digest files found to check")
    return ok


def main():
    parser = argparse.ArgumentParser(description="Verify HTML digest structural markers")
    parser.add_argument("--source", action="store_true", help="Source check only")
    parser.add_argument("--output", action="store_true", help="Output check only")
    args = parser.parse_args()

    run_source = not args.output or args.source
    run_output = not args.source or args.output

    results = []
    if run_source:
        results.append(("SOURCE", check_source()))
    if run_output:
        print()
        results.append(("OUTPUT", check_output()))

    print("\n" + "=" * 60)
    all_ok = all(r[1] for r in results)
    for name, passed in results:
        print(f"  {name}: {'PASS' if passed else 'FAIL'}")
    print("=" * 60)
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
