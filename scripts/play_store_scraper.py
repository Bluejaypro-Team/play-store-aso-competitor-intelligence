#!/usr/bin/env python3
"""
Play Store Metadata Scraper
Fetches public app metadata from Google Play Store web listings with
rate limiting, exponential backoff, and robust regex extraction.
"""

import sys
import time
import urllib.request
import urllib.error
import re
import html
import json
import argparse

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

LAST_REQUEST_TIME = 0.0
MIN_REQUEST_INTERVAL = 1.0  # 1 request per second

def _rate_limited_request(url: str, retries: int = 3) -> str:
    global LAST_REQUEST_TIME
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9,bn;q=0.8"
    }
    
    elapsed = time.monotonic() - LAST_REQUEST_TIME
    if elapsed < MIN_REQUEST_INTERVAL:
        time.sleep(MIN_REQUEST_INTERVAL - elapsed)
        
    req = urllib.request.Request(url, headers=headers)
    backoff = 1.0
    
    for attempt in range(retries):
        try:
            LAST_REQUEST_TIME = time.monotonic()
            with urllib.request.urlopen(req, timeout=12) as response:
                return response.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as e:
            if e.code == 429:
                sys.stderr.write(f"[WARNING] HTTP 429 Rate Limit hit. Backing off for {backoff * 2}s...\n")
                time.sleep(backoff * 2)
                backoff *= 2
            elif e.code >= 500:
                sys.stderr.write(f"[WARNING] HTTP {e.code} server error. Retry {attempt + 1}/{retries} in {backoff}s...\n")
                time.sleep(backoff)
                backoff *= 2
            else:
                sys.stderr.write(f"[ERROR] HTTP {e.code} for URL: {url}\n")
                return ""
        except Exception as ex:
            sys.stderr.write(f"[ERROR] Request failed: {ex}. Retrying...\n")
            time.sleep(backoff)
            backoff *= 2
            
    return ""

def scrape_play_store_app(package_id: str) -> dict:
    url = f"https://play.google.com/store/apps/details?id={package_id}&hl=en&gl=US"
    html_content = _rate_limited_request(url)
    
    if not html_content:
        return {
            "package_id": package_id,
            "url": url,
            "error": "Failed to retrieve page content or app does not exist.",
            "status": "NOT_FOUND"
        }
        
    title = ""
    h1_match = re.search(r'<h1[^>]*>(.*?)</h1>', html_content, re.IGNORECASE | re.DOTALL)
    if h1_match:
        clean_h1 = re.sub(r'<[^>]+>', '', h1_match.group(1)).strip()
        if clean_h1:
            title = html.unescape(clean_h1)
            
    if not title:
        title_tag = re.search(r'<title[^>]*>(.*?)</title>', html_content, re.IGNORECASE | re.DOTALL)
        if title_tag:
            clean_title = re.sub(r'<[^>]+>', '', title_tag.group(1))
            clean_title = clean_title.replace(" - Apps on Google Play", "").strip()
            title = html.unescape(clean_title)
            
    developer = ""
    dev_match = re.search(r'<div class="Vbf7gd">(.*?)</div>', html_content, re.DOTALL)
    if not dev_match:
        dev_match = re.search(r'href="/store/apps/dev[^"]*">(.*?)</a>', html_content, re.DOTALL)
    if not dev_match:
        dev_match = re.search(r'href="/store/apps/developer[^"]*">(.*?)</a>', html_content, re.DOTALL)
    if dev_match:
        developer = html.unescape(re.sub(r'<[^>]+>', '', dev_match.group(1))).strip()
        
    installs = ""
    installs_match = re.search(r'<div class="ClM7O">([0-9\+\,MK\s]+)</div>\s*<div class="g1rdde">Downloads</div>', html_content)
    if not installs_match:
        installs_match = re.search(r'([0-9\+\,MK]+)\s*(?:downloads|installs)', html_content, re.IGNORECASE)
    if installs_match:
        installs = installs_match.group(1).strip()
        
    rating = ""
    rating_match = re.search(r'<div class="TT9OT">([0-9\.]+)</div>\s*<div class="g1rdde">', html_content)
    if not rating_match:
        rating_match = re.search(r'aria-label="Rated\s*([0-9\.]+)\s*stars', html_content, re.IGNORECASE)
    if rating_match:
        rating = rating_match.group(1).strip()
        
    reviews_count = ""
    reviews_match = re.search(r'<div class="g1rdde">([0-9\+\,MK\s]+)\s*reviews</div>', html_content, re.IGNORECASE)
    if reviews_match:
        reviews_count = reviews_match.group(1).strip()
        
    description = ""
    desc_match = re.search(r'data-g-id="description"[^>]*><div[^>]*>(.*?)</div></div>', html_content, re.DOTALL)
    if not desc_match:
        desc_match = re.search(r'<div[^>]*itemprop="description"[^>]*>(.*?)</div>', html_content, re.DOTALL)
    if desc_match:
        raw_desc = desc_match.group(1)
        raw_desc = re.sub(r'<br\s*/?>', '\n', raw_desc, flags=re.IGNORECASE)
        raw_desc = re.sub(r'<[^>]+>', '', raw_desc)
        description = html.unescape(raw_desc).strip()
        
    icon_url = ""
    icon_match = re.search(r'<img[^>]*src="([^"]+)"[^>]*alt="Cover art"', html_content)
    if not icon_match:
        icon_match = re.search(r'<img[^>]*src="([^"]+)"[^>]*itemprop="image"', html_content)
    if icon_match:
        icon_url = icon_match.group(1)
        
    screenshots = []
    for sc in re.finditer(r'<img[^>]*src="([^"]+)"[^>]*alt="Screenshot image"', html_content):
        screenshots.append(sc.group(1))
        
    short_desc = ""
    if description:
        lines = [l.strip() for l in description.split("\n") if l.strip()]
        if lines:
            short_desc = lines[0][:80]
            
    return {
        "package_id": package_id,
        "url": url,
        "status": "SUCCESS",
        "title": title or package_id,
        "title_length": len(title or package_id),
        "developer": developer,
        "installs": installs or "100+",
        "rating": rating or "N/A",
        "reviews_count": reviews_count or "N/A",
        "short_description": short_desc,
        "short_desc_length": len(short_desc),
        "long_description": description,
        "long_desc_length": len(description),
        "icon_url": icon_url,
        "screenshots_count": len(screenshots)
    }

def main():
    parser = argparse.ArgumentParser(description="Scrape Google Play Store App Metadata")
    parser.add_argument("--package", required=True, help="Android package ID (e.g. com.probashi.taka)")
    parser.add_argument("--output", help="Optional output JSON file path")
    args = parser.parse_args()
    
    data = scrape_play_store_app(args.package)
    
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"[SUCCESS] Scraped metadata for '{args.package}' saved to: {args.output}")
    else:
        print(json.dumps(data, indent=2, ensure_ascii=False))

if __name__ == "__main__":
    main()
