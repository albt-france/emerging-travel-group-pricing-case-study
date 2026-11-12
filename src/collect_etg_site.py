"""Take a dated snapshot of Emerging Travel Group's public web pages.

Polite collection rules:
- robots.txt is checked for every URL before fetching (urllib.robotparser);
- one request at a time, at least 2 seconds between requests;
- only public marketing pages (no login, no search results, no prices).

Output: data/raw/etg_site/<YYYY-MM-DD>/<slug>.html and <slug>.txt (visible text),
plus a manifest.csv with URL, HTTP status and collection timestamp.
"""
from __future__ import annotations

import csv
import re
import time
from datetime import datetime, timezone
from html import unescape
from pathlib import Path
from urllib import robotparser
from urllib.parse import urlparse

import requests

ROOT = Path(__file__).resolve().parents[1]
OUT_BASE = ROOT / "data" / "raw" / "etg_site"
USER_AGENT = "Mozilla/5.0 (compatible; case-study-research/1.0)"
DELAY_SECONDS = 2.5

# Public pages listed in emergingtravel.com/sitemap-0.xml + the brands' home/about pages.
URLS = [
    "https://www.emergingtravel.com",
    "https://www.emergingtravel.com/about",
    "https://www.emergingtravel.com/brands",
    "https://www.emergingtravel.com/partners",
    "https://www.emergingtravel.com/travelertype",
    "https://www.emergingtravel.com/marketinghub",
    "https://www.ratehawk.com/",
    "https://www.zenhotels.com/about/",
]


def slugify(url: str) -> str:
    parsed = urlparse(url)
    path = parsed.path.strip("/").replace("/", "_") or "home"
    return f"{parsed.netloc.replace('www.', '').split('.')[0]}_{path}"


def visible_text(html: str) -> str:
    """Very small HTML-to-text: drop scripts/styles, strip tags, collapse spaces."""
    html = re.sub(r"(?is)<(script|style|noscript|svg).*?</\1>", " ", html)
    text = re.sub(r"(?s)<[^>]+>", "\n", html)
    text = unescape(text)
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def allowed(url: str, cache: dict[str, robotparser.RobotFileParser]) -> bool:
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    if base not in cache:
        # Fetch robots.txt with the same User-Agent (urllib's default agent can be refused,
        # which robotparser would wrongly read as "disallow everything").
        rp = robotparser.RobotFileParser()
        resp = requests.get(base + "/robots.txt", headers={"User-Agent": USER_AGENT}, timeout=30)
        rp.parse(resp.text.splitlines() if resp.status_code == 200 else [])
        cache[base] = rp
        time.sleep(DELAY_SECONDS)
    return cache[base].can_fetch(USER_AGENT, url)


def main() -> None:
    day = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    out_dir = OUT_BASE / day
    out_dir.mkdir(parents=True, exist_ok=True)
    robots: dict[str, robotparser.RobotFileParser] = {}
    rows = []
    for url in URLS:
        if not allowed(url, robots):
            print(f"skip (robots.txt): {url}")
            rows.append({"url": url, "status": "disallowed", "collected_utc": ""})
            continue
        resp = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
        slug = slugify(url)
        (out_dir / f"{slug}.html").write_text(resp.text, encoding="utf-8")
        (out_dir / f"{slug}.txt").write_text(visible_text(resp.text), encoding="utf-8")
        stamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
        rows.append({"url": url, "status": resp.status_code, "collected_utc": stamp})
        print(resp.status_code, url)
        time.sleep(DELAY_SECONDS)
    with open(out_dir / "manifest.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["url", "status", "collected_utc"])
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    main()
