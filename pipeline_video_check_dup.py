"""
Check if a video post is a duplicate — either within the current session
or in the SharePoint VideoPosts list.

Usage: python pipeline_video_check_dup.py <title> [<yt_id>]
Output: JSON to stdout with dup_session, dup_sp, sp_id, sp_has_abstract.

Reads session_state.json for session duplicates and sp_videoposts.json for SP duplicates.
Matches by title (primary) and yt_id (secondary).
"""
import json, os, sys, unicodedata, re

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))
SESSION_FILE = os.path.join(BASE, "session_state.json")
SP_VIDEOPOSTS_FILE = os.path.join(BASE, "sp_videoposts.json")


def normalize_title(title):
    """Normalize title for comparison."""
    t = title.lower().strip()
    try:
        t = t.encode('latin-1').decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass
    t = t.replace('\xa0', ' ').replace('\u00a0', ' ')
    t = re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]', '-', t)
    t = re.sub(r'[\u2018\u2019\u201a]', "'", t)
    t = re.sub(r'[\u201c\u201d\u201e]', '"', t)
    t = re.sub(r'[\U0001F300-\U0001FAFF]', '', t)
    t = ''.join(c for c in t if unicodedata.category(c) != 'So')
    t = ' '.join(t.split())
    return t


def check_session_duplicate(title):
    """Check if this title was already processed in the current session."""
    if not os.path.exists(SESSION_FILE):
        return False

    session = json.load(open(SESSION_FILE, encoding="utf-8"))
    processed = session.get("processed_titles", {})

    title_lower = title.lower().strip()

    for proc_title in processed.keys():
        if proc_title.lower().strip() == title_lower:
            return True
    return False


def check_sp_duplicate(title, yt_id=None):
    """Check if this title or yt_id already exists in the SP VideoPosts list.
    Returns (found: bool, sp_id: int|None, has_abstract: bool)."""
    if not os.path.exists(SP_VIDEOPOSTS_FILE):
        return False, None, False

    sp_items = json.load(open(SP_VIDEOPOSTS_FILE, encoding="utf-8"))

    norm_title = normalize_title(title)

    # First pass: match by yt_id if provided
    if yt_id:
        for item in sp_items:
            sp_yt_id = (item.get("yt_id") or "").strip()
            if sp_yt_id and sp_yt_id == yt_id:
                has_abstract = bool((item.get("abstract") or "").strip())
                return True, item.get("id"), has_abstract

    # Second pass: match by title
    for item in sp_items:
        sp_title = normalize_title(item.get("title") or "")
        if sp_title == norm_title:
            has_abstract = bool((item.get("abstract") or "").strip())
            return True, item.get("id"), has_abstract

    return False, None, False


def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline_video_check_dup.py <title> [<yt_id>]")
        sys.exit(1)

    title = sys.argv[1]
    yt_id = sys.argv[2] if len(sys.argv) > 2 else None

    dup_session = check_session_duplicate(title)
    dup_sp, sp_id, sp_has_abstract = check_sp_duplicate(title, yt_id)

    result = {
        "title": title,
        "yt_id": yt_id,
        "dup_session": dup_session,
        "dup_sp": dup_sp,
        "sp_id": sp_id,
        "sp_has_abstract": sp_has_abstract,
        "is_duplicate": dup_session or dup_sp,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
