"""
Fetch blog content from a URL: resolve final URL (after redirects),
extract title, publication date, and full article text.

Usage: python pipeline_fetch_blog.py <url>
Output: JSON to stdout with final_url, title, published_date, content.
"""
import json, re, os, sys
import urllib.request
import html as html_mod

BASE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(BASE, "blog_cache")
os.makedirs(CACHE_DIR, exist_ok=True)


def safe_filename(url):
    """Convert URL to a safe cache filename."""
    name = re.sub(r'https?://', '', url)
    name = re.sub(r'[^a-zA-Z0-9_]', '_', name)
    if len(name) > 120:
        name = name[:120]
    return name + ".txt"


def resolve_url(url):
    """Follow redirects and return the final URL."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        resp = urllib.request.urlopen(req, timeout=15)
        return resp.url
    except Exception:
        return url


def fetch_and_extract(url):
    """Fetch HTML from URL and extract article content."""
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    })
    resp = urllib.request.urlopen(req, timeout=20)
    final_url = resp.url
    raw_html = resp.read().decode("utf-8", errors="replace")

    # --- Extract publication date ---
    published_date = extract_published_date(raw_html, final_url)

    # --- Extract title ---
    title = extract_title(raw_html)

    # --- Extract article text ---
    content = extract_article_text(raw_html)

    return {
        "final_url": final_url,
        "title": title,
        "published_date": published_date,
        "content": content,
        "content_length": len(content),
    }


def extract_published_date(raw_html, url):
    """Extract publication date from page metadata or URL."""
    # Strategy 1: JSON-LD schema
    ld_match = re.search(
        r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>',
        raw_html, re.S
    )
    if ld_match:
        try:
            ld = json.loads(ld_match.group(1))
            date_str = ld.get("datePublished") or ld.get("dateCreated") or ""
            if date_str:
                m = re.match(r'(\d{4})-(\d{2})-(\d{2})', date_str)
                if m:
                    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        except Exception:
            pass

    # Strategy 2: meta tags
    for pattern in [
        r'<meta[^>]*property="article:published_time"[^>]*content="(\d{4}-\d{2}-\d{2})',
        r'<meta[^>]*name="publication[_-]?date"[^>]*content="(\d{4}-\d{2}-\d{2})',
        r'<time[^>]*datetime="(\d{4}-\d{2}-\d{2})',
    ]:
        m = re.search(pattern, raw_html, re.I)
        if m:
            return m.group(1)

    # Strategy 3: date in URL path
    m = re.search(r'/(\d{4})/(\d{2})/(\d{2})/', url)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    return ""


def extract_title(raw_html):
    """Extract page title."""
    # JSON-LD
    ld_match = re.search(
        r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>',
        raw_html, re.S
    )
    if ld_match:
        try:
            ld = json.loads(ld_match.group(1))
            t = ld.get("headline") or ld.get("name") or ""
            if t:
                return html_mod.unescape(t).strip()
        except Exception:
            pass

    # <title> tag
    m = re.search(r'<title[^>]*>(.*?)</title>', raw_html, re.S | re.I)
    if m:
        title = html_mod.unescape(re.sub(r'<[^>]+>', '', m.group(1))).strip()
        # Remove common suffixes
        for suffix in [" | Microsoft", " - Microsoft", " | Azure", " - TechCommunity"]:
            if title.endswith(suffix):
                title = title[:-len(suffix)].strip()
        return title

    # og:title
    m = re.search(r'<meta[^>]*property="og:title"[^>]*content="([^"]+)"', raw_html, re.I)
    if m:
        return html_mod.unescape(m.group(1)).strip()

    return ""


def extract_article_text(raw_html):
    """Extract clean article text content, preserving bold/italic formatting."""
    body_text = ""

    # Strategy 1: JSON-LD BlogPosting articleBody
    ld_match = re.search(
        r'<script[^>]*type="application/ld\+json"[^>]*>(.*?)</script>',
        raw_html, re.S
    )
    if ld_match:
        try:
            ld = json.loads(ld_match.group(1))
            body_text = ld.get("articleBody") or ld.get("description") or ""
        except Exception:
            pass

    # Strategy 2: HTML article/main content
    if len(body_text) < 200:
        body_text = _extract_from_html(raw_html)

    if not body_text:
        return ""

    # Clean up whitespace
    body_text = re.sub(r'\s+', ' ', body_text).strip()

    # Limit total length
    if len(body_text) > 8000:
        body_text = body_text[:8000]

    return body_text


def _extract_from_html(raw_html):
    """Extract article body from HTML, preserving bold/italic tags."""
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

    # Normalize bold/italic
    text = re.sub(r'<(strong|b)\b[^>]*>', '<b>', text, flags=re.I)
    text = re.sub(r'</(strong|b)\b[^>]*>', '</b>', text, flags=re.I)
    text = re.sub(r'<(em|i)\b[^>]*>', '<i>', text, flags=re.I)
    text = re.sub(r'</(em|i)\b[^>]*>', '</i>', text, flags=re.I)

    # Strip all other HTML tags
    text = re.sub(r'<(?!/?[bi]>)[^>]+>', ' ', text)
    text = html_mod.unescape(text)
    return text


def main():
    if len(sys.argv) < 2:
        print("Usage: python pipeline_fetch_blog.py <url>")
        sys.exit(1)

    url = sys.argv[1]

    # Check cache first
    cache_file = os.path.join(CACHE_DIR, safe_filename(url))
    cached_content = None
    if os.path.exists(cache_file):
        cached_content = open(cache_file, encoding="utf-8").read()

    try:
        result = fetch_and_extract(url)

        # Cache the result
        if result["content"] and len(result["content"]) > 50:
            with open(cache_file, "w", encoding="utf-8") as f:
                f.write(f"TITLE: {result['title']}\n")
                f.write(f"URL: {result['final_url']}\n")
                f.write(f"PUBLISHED: {result['published_date']}\n")
                f.write("---\n")
                f.write(result["content"][:5000])

        print(json.dumps(result, indent=2, ensure_ascii=False))

    except Exception as e:
        # If fetch fails but we have cache, use that
        if cached_content:
            lines = cached_content.split("\n")
            title = lines[0].replace("TITLE: ", "") if lines[0].startswith("TITLE: ") else ""
            cached_url = lines[1].replace("URL: ", "") if len(lines) > 1 and lines[1].startswith("URL: ") else url
            pub_date = lines[2].replace("PUBLISHED: ", "") if len(lines) > 2 and lines[2].startswith("PUBLISHED: ") else ""
            sep_idx = cached_content.find("---\n")
            content = cached_content[sep_idx+4:] if sep_idx >= 0 else cached_content

            result = {
                "final_url": cached_url,
                "title": title,
                "published_date": pub_date,
                "content": content,
                "content_length": len(content),
                "from_cache": True,
            }
            print(json.dumps(result, indent=2, ensure_ascii=False))
        else:
            error = {"error": str(e), "url": url, "final_url": url}
            print(json.dumps(error, indent=2, ensure_ascii=False))
            sys.exit(1)


if __name__ == "__main__":
    main()
