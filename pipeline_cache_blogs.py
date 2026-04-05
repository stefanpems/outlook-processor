"""Extract clean article text from all blog URLs and cache locally.
This prepares content for LLM summarization."""
import json, glob, os, re, urllib.request, html as html_mod, sys

sys.stdout = open(sys.stdout.fileno(), mode='w', buffering=1)

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = json.load(open(os.path.join(BASE, "config.json"), encoding="utf-8"))

CACHE_DIR = os.path.join(BASE, "blog_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

OUTPUT_DIR = os.path.join(BASE, CONFIG["output"]["dir"])

# Collect all unique URLs
urls = {}
for f in sorted(glob.glob(os.path.join(OUTPUT_DIR, "emails_*.json"))):
    data = json.load(open(f, encoding="utf-8"))
    for em in data:
        url = em.get("blog_link", "")
        title = em.get("title", "")
        if url and url not in urls:
            urls[url] = title

print(f"Total unique URLs: {len(urls)}")

def extract_article_text(raw_html):
    """Extract clean article body text from HTML."""
    # Strategy 1: JSON-LD BlogPosting
    ld_match = re.search(
        r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>',
        raw_html, re.S
    )
    if ld_match:
        try:
            ld = json.loads(ld_match.group(1))
            body = ld.get("articleBody") or ld.get("description") or ""
            if len(body) > 200:
                return body
        except Exception:
            pass

    # Strategy 2: HTML content extraction
    text = re.sub(r'<script[^>]*>.*?</script>', '', raw_html, flags=re.S)
    text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.S)
    text = re.sub(r'<nav[^>]*>.*?</nav>', '', text, flags=re.S)
    text = re.sub(r'<header[^>]*>.*?</header>', '', text, flags=re.S)
    text = re.sub(r'<footer[^>]*>.*?</footer>', '', text, flags=re.S)

    article_match = re.search(r'<article[^>]*>(.*?)</article>', text, re.S)
    if article_match:
        text = article_match.group(1)
    else:
        main_match = re.search(r'<main[^>]*>(.*?)</main>', text, re.S)
        if main_match:
            text = main_match.group(1)

    text = re.sub(r'<[^>]+>', ' ', text)
    text = html_mod.unescape(text)
    text = re.sub(r'\s+', ' ', text).strip()

    # Remove footer noise
    for marker in ['Share this page', 'What\'s new Surface', 'California Consumer Privacy']:
        idx = text.find(marker)
        if idx > 200:
            text = text[:idx]

    return text


success = 0
fail = 0
skip = 0

for i, (url, title) in enumerate(urls.items(), 1):
    # Use URL hash as filename
    safe_name = re.sub(r'[^a-zA-Z0-9]', '_', url[-60:])
    cache_file = os.path.join(CACHE_DIR, f"{safe_name}.txt")
    
    if os.path.exists(cache_file):
        skip += 1
        continue

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=15)
        raw = resp.read().decode("utf-8", errors="replace")
        text = extract_article_text(raw)

        if len(text) > 100:
            with open(cache_file, "w", encoding="utf-8") as f:
                f.write(f"TITLE: {title}\n")
                f.write(f"URL: {url}\n")
                f.write(f"---\n")
                f.write(text[:5000])  # Cap at 5000 chars
            success += 1
            print(f"[{i}/{len(urls)}] OK: {title[:60]}")
        else:
            with open(cache_file, "w", encoding="utf-8") as f:
                f.write(f"TITLE: {title}\n")
                f.write(f"URL: {url}\n")
                f.write(f"---\n")
                f.write("(Content too short or not extractable)")
            fail += 1
            print(f"[{i}/{len(urls)}] SHORT: {title[:60]}")
    except Exception as e:
        with open(cache_file, "w", encoding="utf-8") as f:
            f.write(f"TITLE: {title}\n")
            f.write(f"URL: {url}\n")
            f.write(f"---\n")
            f.write(f"(Error: {str(e)[:100]})")
        fail += 1
        print(f"[{i}/{len(urls)}] FAIL: {title[:60]}: {e}")

print(f"\nDone: {success} fetched, {fail} failed, {skip} cached")
