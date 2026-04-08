"""Shared helper to ensure Edge is running with CDP (remote debugging) enabled.

Usage in any pipeline script:
    from cdp_helper import ensure_edge_cdp
    ensure_edge_cdp()  # call before connect_over_cdp
"""
import json, os, sys, subprocess, time
try:
    from urllib.request import urlopen, Request
    from urllib.error import URLError
except ImportError:
    pass

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))


def _cdp_is_reachable(cdp_url: str, timeout: float = 3) -> bool:
    """Check if Edge CDP endpoint responds."""
    try:
        req = Request(f"{cdp_url}/json/version")
        resp = urlopen(req, timeout=timeout)
        return resp.status == 200
    except Exception:
        return False


def ensure_edge_cdp() -> None:
    """Ensure Edge is running with remote debugging.

    Reads edge_cdp config from config.json. If CDP endpoint is already
    reachable, returns immediately. Otherwise, launches Edge with the
    debug profile and waits until CDP is ready (up to 15 seconds).
    """
    cfg = CONFIG["edge_cdp"]
    cdp_url = cfg["url"]

    if _cdp_is_reachable(cdp_url):
        return

    edge_path = cfg["edge_path"]
    profile_dir = os.path.expandvars(cfg["profile_dir_env"])
    profile_name = cfg["profile_name"]

    # Parse port from URL
    port = cdp_url.split(":")[-1].split("/")[0]

    print(f"[cdp_helper] Edge CDP not reachable at {cdp_url}, launching Edge...",
          file=sys.stderr, flush=True)

    subprocess.Popen(
        [
            edge_path,
            f"--remote-debugging-port={port}",
            f"--user-data-dir={profile_dir}",
            f"--profile-directory={profile_name}",
            "--no-first-run",
            "--no-default-browser-check",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    # Wait for CDP to become reachable
    max_wait = 15
    interval = 1
    elapsed = 0
    while elapsed < max_wait:
        time.sleep(interval)
        elapsed += interval
        if _cdp_is_reachable(cdp_url):
            print(f"[cdp_helper] Edge CDP ready after {elapsed}s.",
                  file=sys.stderr, flush=True)
            return

    print(f"[cdp_helper] WARNING: Edge CDP still not reachable after {max_wait}s. "
          "Proceeding anyway — connect_over_cdp may fail.",
          file=sys.stderr, flush=True)


def ensure_all_folders_scope(page) -> bool:
    """After an Outlook search, expand scope to 'Tutte le cartelle' if needed.

    Outlook Web defaults search scope to the current folder (usually Inbox).
    Emails auto-sorted by mail rules into other folders won't appear.
    This function checks the scope indicator near the top of the page and
    clicks to expand to all folders when necessary.

    Returns True if scope was already 'all' or was successfully expanded.
    """
    # Look for the scope indicator button near the top-left of the search bar
    for btn in page.query_selector_all("button"):
        try:
            if not btn.is_visible():
                continue
            box = btn.bounding_box()
            if not box or box["y"] > 50 or box["x"] > 350:
                continue
            txt = (btn.inner_text() or "").strip().lower()
            # Already on "all folders" scope
            if "tutte le cartelle" in txt or "all folder" in txt or "all mail" in txt:
                return True
            # Narrow scope detected — need to expand
            if any(k in txt for k in ["posta in arrivo", "inbox", "cartella corrente",
                                       "current folder", "tech-", "tech "]):
                print(f"[cdp_helper] Search scope is '{btn.inner_text().strip()}', expanding to all folders...",
                      file=sys.stderr, flush=True)
                btn.click()
                page.wait_for_timeout(2000)
                # Select "Tutte le cartelle" from dropdown
                for opt in page.query_selector_all(
                    "[role='option'], [role='menuitem'], [role='menuitemradio'], button"
                ):
                    try:
                        ot = (opt.inner_text() or "").strip()
                        if "tutte le cartelle" in ot.lower() or "all folder" in ot.lower():
                            opt.click()
                            page.wait_for_timeout(4000)
                            print("[cdp_helper] Scope expanded to 'Tutte le cartelle'.",
                                  file=sys.stderr, flush=True)
                            return True
                    except Exception:
                        pass
                # Dropdown didn't have the expected option — dismiss and continue
                page.keyboard.press("Escape")
                page.wait_for_timeout(500)
                print("[cdp_helper] WARNING: could not find 'Tutte le cartelle' option in dropdown.",
                      file=sys.stderr, flush=True)
                return False
        except Exception:
            pass
    # No scope button found — may already be fine or search UI is different
    return True
