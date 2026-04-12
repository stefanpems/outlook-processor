"""
Batch processor for blog notification emails.

Reads session_state.json, groups emails by title, and performs operations
on each group using a single CDP connection.

Supported phases (run one at a time):
  python pipeline_batch.py email_actions   — categorize + move all emails
  python pipeline_batch.py sp_create       — create/update SP items for emails with summaries
  python pipeline_batch.py fetch           — fetch blog content for all unique URLs
  python pipeline_batch.py dupcheck        — run dedup checks for all emails

Design principles:
- Groups emails by TITLE so one Outlook search covers all emails with the same subject
- Single CDP/Playwright connection for the entire run
- Saves session_state.json after each group to preserve progress
- Idempotent: re-running skips already-completed work
"""
import json, re, os, sys, time, subprocess
from collections import defaultdict
from playwright.sync_api import sync_playwright
from cdp_helper import ensure_edge_cdp

sys.stdout = open(sys.stdout.fileno(), mode="w", buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
SESSION_FILE = os.path.join(BASE, "session_state.json")
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CDP_URL = CONFIG["edge_cdp"]["url"]
SENDER = CONFIG["outlook"]["sender"]
SUBJECT_PREFIX = CONFIG["outlook"]["subject_prefix"]
CATEGORY = CONFIG["outlook"]["processed_category"]
TARGET_FOLDER = CONFIG["outlook"]["target_folder"]


def load_session():
    with open(SESSION_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_session(session):
    with open(SESSION_FILE, "w", encoding="utf-8") as f:
        json.dump(session, f, indent=2, ensure_ascii=False)


def run_pipeline_script(args, stdin_data=None, timeout=120):
    """Run a pipeline_*.py script, return parsed JSON from its stdout."""
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    r = subprocess.run(
        [sys.executable] + args,
        input=stdin_data, capture_output=True, text=True,
        timeout=timeout, cwd=BASE, env=env, encoding="utf-8"
    )
    stdout = r.stdout.strip()
    # Brace-matching parser for multi-line JSON
    end = stdout.rfind("}")
    if end >= 0:
        depth = 0
        for i in range(end, -1, -1):
            if stdout[i] == "}":
                depth += 1
            elif stdout[i] == "{":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(stdout[i:end + 1])
                    except json.JSONDecodeError:
                        break
    return {"_raw": stdout[:500], "_stderr": r.stderr[:500] if r.stderr else ""}


# ═══════════════════════════════════════════════════════════════
#  Phase: email_actions  (categorize + move)
# ═══════════════════════════════════════════════════════════════

def do_search(page, title):
    """Search Outlook for blog emails by title. No negative category filter."""
    safe = re.sub(r'["\u2018\u2019\u201c\u201d\u2013\u2014\u00a0]', ' ', title)
    safe = re.sub(r'\s+', ' ', safe).strip()
    words = safe.split()
    short = " ".join(words[:6]) if len(words) > 6 else safe
    query = f'from:{SENDER} subject:({SUBJECT_PREFIX} {short})'

    try:
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)
    except Exception:
        page.evaluate("window.location.href = 'https://outlook.cloud.microsoft/mail/'")
        page.wait_for_timeout(5000)
        sb = page.wait_for_selector("#topSearchInput", timeout=10000)

    sb.click()
    page.wait_for_timeout(400)
    page.keyboard.press("Control+a")
    page.keyboard.press("Backspace")
    page.wait_for_timeout(200)
    page.keyboard.type(query, delay=5)
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(6000)


def get_blog_rows(page):
    rows = page.query_selector_all("div[role='option']")
    return [r for r in rows if "[Blog-" in (r.inner_text() or "")]


def has_category(page):
    try:
        rp = page.locator('[role="main"]').first.inner_text()
        return CATEGORY.lower() in rp.lower()
    except Exception:
        return False


def get_row_folder(row):
    try:
        lines = [l.strip() for l in row.inner_text().strip().split("\n") if l.strip()]
        if lines:
            parts = [p.strip() for p in lines[-1].replace("\t", "|").split("|") if p.strip()]
            if parts:
                return parts[-1]
    except Exception:
        pass
    return ""


def do_categorize(page, row):
    try:
        box = row.bounding_box()
        if not box:
            return False
        cx, cy = box["x"] + box["width"] / 2, box["y"] + box["height"] / 2
        page.mouse.click(cx, cy)
        page.wait_for_timeout(500)
        page.mouse.click(cx, cy, button="right")
    except Exception:
        return False
    page.wait_for_timeout(2000)

    cat_btn = None
    for item in page.query_selector_all("[role='menuitem']"):
        try:
            if "categorizza" in item.inner_text().lower():
                cat_btn = item
                break
        except Exception:
            pass
    if not cat_btn:
        page.keyboard.press("Escape")
        return False

    cat_btn.evaluate("el => el.click()")
    page.wait_for_timeout(1500)

    for sub in page.query_selector_all("[role='menuitemcheckbox'], [role='menuitem']"):
        try:
            txt = sub.inner_text().strip()
            if CATEGORY in txt:
                if sub.get_attribute("aria-checked") == "true":
                    page.keyboard.press("Escape")
                    return True
                sub.evaluate("el => el.click()")
                page.wait_for_timeout(800)
                return True
        except Exception:
            pass
    page.keyboard.press("Escape")
    return False


def go_home_tab(page):
    for el in page.query_selector_all("[role='tab']"):
        try:
            if "home" in el.inner_text().strip().lower():
                el.click()
                page.wait_for_timeout(1000)
                return True
        except Exception:
            pass
    return False


def do_move(page, row):
    try:
        box = row.bounding_box()
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            page.wait_for_timeout(400)
    except Exception:
        return False

    go_home_tab(page)
    page.wait_for_timeout(500)

    sposta_btn = None
    for btn in page.query_selector_all("button"):
        try:
            if not btn.is_visible():
                continue
            box = btn.bounding_box()
            if not box or box["y"] > 200:
                continue
            txt = btn.inner_text().strip()
            al = btn.get_attribute("aria-label") or ""
            if txt == "Sposta" or (al == "Sposta" and "alto" not in al):
                sposta_btn = btn
                break
        except Exception:
            pass
    if not sposta_btn:
        for btn in page.query_selector_all("button[aria-haspopup], button[aria-expanded]"):
            try:
                if not btn.is_visible():
                    continue
                box = btn.bounding_box()
                if not box or box["y"] > 200:
                    continue
                al = btn.get_attribute("aria-label") or ""
                if "sposta" in al.lower() and "alto" not in al.lower():
                    sposta_btn = btn
                    break
            except Exception:
                pass
    if not sposta_btn:
        return False

    try:
        sposta_btn.click(timeout=5000)
    except Exception:
        # Button found but not enabled — try bounding-box click
        try:
            sbox = sposta_btn.bounding_box()
            if sbox:
                page.mouse.click(sbox["x"] + sbox["width"] / 2,
                                 sbox["y"] + sbox["height"] / 2)
            else:
                return False
        except Exception:
            return False
    page.wait_for_timeout(2000)

    search_input = None
    for inp in page.query_selector_all("input"):
        try:
            if not inp.is_visible():
                continue
            ph = inp.get_attribute("placeholder") or ""
            if "cartella" in ph.lower() or "folder" in ph.lower() or "cerca" in ph.lower():
                search_input = inp
                break
        except Exception:
            pass
    if not search_input:
        page.keyboard.press("Escape")
        return False

    search_input.click()
    page.wait_for_timeout(300)
    search_input.fill(TARGET_FOLDER)
    page.wait_for_timeout(3000)
    page.keyboard.press("ArrowDown")
    page.wait_for_timeout(500)
    page.keyboard.press("Enter")
    page.wait_for_timeout(2000)
    return True


def phase_email_actions(session):
    """Categorize + move all emails, grouped by title."""
    emails = session["emails"]

    # Group indices by title
    groups = defaultdict(list)
    for i, em in enumerate(emails):
        if em.get("title"):
            groups[em["title"]].append(i)

    todo = {t: idxs for t, idxs in groups.items()
            if any(emails[i].get("categorized") != "Yes" or emails[i].get("moved") != "Yes"
                   for i in idxs)}

    print(f"Email actions: {len(todo)} title-groups ({sum(len(v) for v in todo.values())} emails)")
    if not todo:
        print("  Nothing to do!")
        return

    p = sync_playwright().start()
    try:
        ensure_edge_cdp()
        browser = p.chromium.connect_over_cdp(CDP_URL)
        ctx = browser.contexts[0]
        page = ctx.new_page()
        page.goto("https://outlook.cloud.microsoft/mail/",
                   wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(6000)

        stats = {"cat": 0, "mov": 0, "skip": 0, "err": 0}

        for gi, (title, indices) in enumerate(todo.items()):
            n = len(indices)
            print(f"\n[{gi + 1}/{len(todo)}] \"{title[:65]}\" ({n} email{'s' if n > 1 else ''})")

            do_search(page, title)
            blog_rows = get_blog_rows(page)
            print(f"  Found {len(blog_rows)} [Blog-] row(s)")

            if not blog_rows:
                print(f"  0 results — marking all {n} as done (already processed)")
                for i in indices:
                    emails[i].setdefault("categorized", "Yes")
                    emails[i].setdefault("moved", "Yes")
                save_session(session)
                continue

            # Process each visible row
            ri = 0
            while ri < len(blog_rows):
                row = blog_rows[ri]
                try:
                    box = row.bounding_box()
                    if not box:
                        ri += 1
                        continue
                    page.mouse.click(box["x"] + box["width"] / 2,
                                     box["y"] + box["height"] / 2)
                    page.wait_for_timeout(1500)
                except Exception:
                    ri += 1
                    continue

                try:
                    already_cat = has_category(page)
                    folder = get_row_folder(row)
                    in_target = TARGET_FOLDER.lower() in folder.lower()

                    if already_cat and in_target:
                        stats["skip"] += 1
                        ri += 1
                        continue

                    if not already_cat:
                        ok = do_categorize(page, row)
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(500)
                        if ok:
                            stats["cat"] += 1
                            print(f"    row {ri + 1}: categorized")
                        else:
                            stats["err"] += 1
                            print(f"    row {ri + 1}: categorize FAILED")

                    if not in_target:
                        blog_rows = get_blog_rows(page)
                        if ri < len(blog_rows):
                            row = blog_rows[ri]
                        elif blog_rows:
                            row = blog_rows[0]
                            ri = 0
                        else:
                            break

                        ok = do_move(page, row)
                        if ok:
                            stats["mov"] += 1
                            print(f"    row {ri + 1}: moved")
                            page.wait_for_timeout(2000)
                            blog_rows = get_blog_rows(page)
                            continue  # don't increment — row disappeared
                        else:
                            stats["err"] += 1
                            print(f"    row {ri + 1}: move FAILED")
                except Exception as exc:
                    stats["err"] += 1
                    print(f"    row {ri + 1}: ERROR {exc!r}")

                ri += 1

            # Mark ALL session emails with this title as done
            for i in indices:
                emails[i]["categorized"] = "Yes"
                emails[i]["moved"] = "Yes"
            save_session(session)
            time.sleep(0.5)

        page.close()
        done = sum(1 for e in emails if e.get("categorized") == "Yes" and e.get("moved") == "Yes")
        print(f"\nDone: cat={stats['cat']} mov={stats['mov']} skip={stats['skip']} err={stats['err']}")
        print(f"Session: {done}/{len(emails)} emails complete")
    finally:
        p.stop()


# ═══════════════════════════════════════════════════════════════
#  Phase: sp_create  (create new + update summaries)
# ═══════════════════════════════════════════════════════════════

def phase_sp_create(session):
    """Create/update SP items for emails that have summaries."""
    emails = session["emails"]
    targets = []
    for i, em in enumerate(emails):
        if not em.get("summary"):
            continue
        if em.get("dup_session"):
            continue
        if em.get("sp_created") == "Yes" or em.get("sp_updated") == "Yes":
            continue
        if em.get("dup_sp") == "Yes" and em.get("sp_has_summary"):
            continue
        targets.append(i)

    print(f"SP create/update: {len(targets)} emails to process")
    if not targets:
        print("  Nothing to do!")
        return

    for idx in targets:
        em = emails[idx]
        title = em["title"]

        if em.get("dup_sp") == "Yes" and not em.get("sp_has_summary"):
            sp_id = em.get("sp_id")
            print(f"  [{idx}] Update summary on SP ID={sp_id}: {title[:55]}")
            payload = json.dumps({"summary": em["summary"], "title": title}, ensure_ascii=False)
            r = run_pipeline_script(
                ["pipeline_sp_create.py", "--update-summary", str(sp_id), "-"],
                stdin_data=payload
            )
            if r.get("ok"):
                em["sp_updated"] = "Yes"
                print(f"       -> OK")
            else:
                print(f"       -> FAIL: {r}")
        else:
            print(f"  [{idx}] Create: {title[:55]}")
            payload = json.dumps({
                "title": title,
                "published_date": em.get("published_date", ""),
                "summary": em["summary"],
                "topic": em.get("topic", ""),
                "tech": em.get("tech", ""),
                "blog_link": em.get("final_url", em.get("blog_link", ""))
            }, ensure_ascii=False)
            r = run_pipeline_script(["pipeline_sp_create.py", "-"], stdin_data=payload)
            if r.get("ok"):
                em["sp_created"] = "Yes"
                em["sp_id"] = r.get("id")
                print(f"       -> OK (ID={r.get('id')})")
            else:
                print(f"       -> FAIL: {r}")

        save_session(session)

    created = sum(1 for e in emails if e.get("sp_created") == "Yes")
    updated = sum(1 for e in emails if e.get("sp_updated") == "Yes")
    print(f"\nTotal: {created} created, {updated} updated")


# ═══════════════════════════════════════════════════════════════
#  Phase: fetch  (fetch blog content for all unique URLs)
# ═══════════════════════════════════════════════════════════════

def phase_fetch(session):
    """Fetch blog content for all emails, deduplicating by URL."""
    emails = session["emails"]
    url_map = {}  # url -> first result
    fetched = 0
    cached = 0

    for i, em in enumerate(emails):
        url = em.get("blog_link", "")
        if not url:
            continue
        if em.get("final_url") and em.get("content_length", 0) > 0:
            # Already fetched — cache for other emails with same URL
            if url not in url_map:
                url_map[url] = em
            cached += 1
            continue
        if url in url_map:
            # Copy from previously fetched
            src = url_map[url]
            for k in ("final_url", "title", "published_date", "content_length", "cache_file"):
                if k in src:
                    em[k] = src[k]
            cached += 1
            continue

        print(f"  [{i}] Fetching: {url[:70]}")
        r = run_pipeline_script(["pipeline_fetch_blog.py", url])
        if r.get("final_url"):
            em["final_url"] = r["final_url"]
            if r.get("published_date"):
                em["published_date"] = r["published_date"]
            em["content_length"] = r.get("content_length", 0)
            url_map[url] = em
            fetched += 1
        else:
            print(f"       -> FAIL: {r.get('_raw', '')[:100]}")

        save_session(session)

    print(f"\nFetch: {fetched} new, {cached} cached/copied")


# ═══════════════════════════════════════════════════════════════
#  Phase: dupcheck  (check session + SP duplicates)
# ═══════════════════════════════════════════════════════════════

def phase_dupcheck(session):
    """Run dedup check for all emails."""
    emails = session["emails"]
    checked = 0

    for i, em in enumerate(emails):
        title = em.get("title", "")
        final_url = em.get("final_url", em.get("blog_link", ""))
        if not title or not final_url:
            continue
        if em.get("dup_sp") is not None and em.get("dup_sp") != "":
            continue  # already checked

        r = run_pipeline_script(["pipeline_check_dup.py", title, final_url])
        if "dup_session" in r:
            em["dup_session"] = r.get("dup_session", "")
            em["dup_sp"] = "Yes" if r.get("dup_sp") else ""
            em["sp_id"] = r.get("sp_id")
            em["sp_has_summary"] = r.get("sp_has_summary", False)
            checked += 1
            status = "DUP_SESSION" if r.get("dup_session") else ("DUP_SP" if r.get("dup_sp") else "NEW")
            print(f"  [{i}] {status}: {title[:55]}")
        else:
            print(f"  [{i}] FAIL: {r.get('_raw', '')[:100]}")

        save_session(session)

    print(f"\nDupcheck: {checked} checked")


# ═══════════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline_batch.py <phase>")
        print("Phases: fetch, dupcheck, sp_create, email_actions")
        sys.exit(1)

    phase = sys.argv[1]
    session = load_session()
    print(f"Session: {len(session['emails'])} emails")

    if phase == "fetch":
        phase_fetch(session)
    elif phase == "dupcheck":
        phase_dupcheck(session)
    elif phase == "sp_create":
        phase_sp_create(session)
    elif phase == "email_actions":
        phase_email_actions(session)
    else:
        print(f"Unknown phase: {phase}")
        print("Phases: fetch, dupcheck, sp_create, email_actions")
        sys.exit(1)


if __name__ == "__main__":
    main()
