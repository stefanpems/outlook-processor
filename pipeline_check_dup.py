"""
Check if a blog post is a duplicate — either within the current session
or in the SharePoint BlogPosts list.

Usage: python pipeline_check_dup.py <title> <final_url>
Output: JSON to stdout with dup_session, dup_sp, sp_id, sp_has_summary.

Reads session_state.json for session duplicates and sp_blogposts.json for SP duplicates.
"""
import json, os, sys
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))
SESSION_FILE = os.path.join(BASE, "session_state.json")
SP_BLOGPOSTS_FILE = os.path.join(BASE, "sp_blogposts.json")


def normalize_url(url):
    """Normalize URL for comparison: strip protocol, trailing slash, query params."""
    url = url.lower().strip()
    url = url.replace("https://", "").replace("http://", "")
    url = url.split("?")[0].split("#")[0]
    url = url.rstrip("/")
    return url


def check_session_duplicate(title, final_url):
    """Check if this title+URL was already processed in the current session."""
    if not os.path.exists(SESSION_FILE):
        return False

    session = json.load(open(SESSION_FILE, encoding="utf-8"))
    processed = session.get("processed_titles", {})

    # Check by title (case-insensitive)
    title_lower = title.lower().strip()
    norm_url = normalize_url(final_url)

    for proc_title, proc_url in processed.items():
        if proc_title.lower().strip() == title_lower:
            if normalize_url(proc_url) == norm_url:
                return True
    return False


def normalize_title(title):
    """Normalize title for comparison: lowercase, collapse whitespace, fix encoding,
    normalize Unicode hyphens/dashes, and strip emoji/symbol characters."""
    import unicodedata, re
    t = title.lower().strip()
    # Fix double-encoded UTF-8: Â\xa0 -> \xa0, â\x80\x93 -> \u2013, etc.
    try:
        t = t.encode('latin-1').decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass
    # Replace non-breaking spaces with regular spaces
    t = t.replace('\xa0', ' ').replace('\u00a0', ' ')
    # Normalize all Unicode hyphens/dashes to ASCII hyphen-minus (U+002D)
    # U+2010 hyphen, U+2011 non-breaking hyphen, U+2012 figure dash,
    # U+2013 en dash, U+2014 em dash, U+2015 horizontal bar, U+2212 minus sign
    t = re.sub(r'[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]', '-', t)
    # Normalize Unicode quotes to ASCII equivalents
    t = re.sub(r'[\u2018\u2019\u201a]', "'", t)  # single quotes
    t = re.sub(r'[\u201c\u201d\u201e]', '"', t)   # double quotes
    # Strip emoji and other symbol characters (So category) that may prefix titles
    t = re.sub(r'[\U0001F300-\U0001FAFF]', '', t)  # supplemental symbols, emoticons, etc.
    t = ''.join(c for c in t if unicodedata.category(c) != 'So')  # catch remaining symbols
    # Collapse whitespace
    t = ' '.join(t.split())
    return t


def check_sp_duplicate(title, final_url):
    """Check if this title+URL already exists in the SP BlogPosts list.
    Returns (found: bool, sp_id: int|None, has_summary: bool)."""
    if not os.path.exists(SP_BLOGPOSTS_FILE):
        return False, None, False

    sp_items = json.load(open(SP_BLOGPOSTS_FILE, encoding="utf-8"))

    norm_title = normalize_title(title)
    norm_url = normalize_url(final_url)

    for item in sp_items:
        sp_title = normalize_title(item.get("title") or "")
        sp_url = normalize_url(item.get("url") or "")

        # Match by title AND URL
        if sp_title == norm_title and sp_url and norm_url and sp_url == norm_url:
            has_summary = bool((item.get("summary") or "").strip())
            return True, item.get("id"), has_summary

        # Also match by title only (some URLs might differ slightly)
        if sp_title == norm_title:
            has_summary = bool((item.get("summary") or "").strip())
            return True, item.get("id"), has_summary

    return False, None, False


def main():
    if len(sys.argv) < 3:
        print("Usage: python pipeline_check_dup.py <title> <final_url>")
        sys.exit(1)

    title = sys.argv[1]
    final_url = sys.argv[2]

    dup_session = check_session_duplicate(title, final_url)
    dup_sp, sp_id, sp_has_summary = check_sp_duplicate(title, final_url)

    result = {
        "title": title,
        "final_url": final_url,
        "dup_session": dup_session,
        "dup_sp": dup_sp,
        "sp_id": sp_id,
        "sp_has_summary": sp_has_summary,
        "is_duplicate": dup_session or dup_sp,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
