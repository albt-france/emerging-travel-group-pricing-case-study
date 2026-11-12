"""Read archived copies of emergingtravel.com (Internet Archive Wayback Machine)
to see how the headline figures shown on the home page changed over time.

- One snapshot per year (closest to July), 6 requests, 3 s apart.
- Raw HTML saved to data/raw/etg_site/wayback/<timestamp>.html.
- The figures are then read by hand / regex and stored in data/public/etg_headline_history.csv
  (see extract step below, which only keeps numbers next to known labels).
"""
from __future__ import annotations

import re
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "etg_site" / "wayback"
UA = {"User-Agent": "Mozilla/5.0 (compatible; case-study-research/1.0)"}
TIMESTAMPS = ["20210723222605", "20220705064538", "20230705152153",
              "20240703023127", "20250702012436", "20260606170348"]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for ts in TIMESTAMPS:
        path = OUT / f"{ts}.html"
        if not path.exists():
            # "id_" returns the archived page without the Wayback toolbar.
            url = f"https://web.archive.org/web/{ts}id_/https://www.emergingtravel.com/"
            resp = requests.get(url, headers=UA, timeout=60)
            path.write_text(resp.text, encoding="utf-8")
            print(resp.status_code, ts, len(resp.text))
            time.sleep(3)
        html = path.read_text(encoding="utf-8")
        text = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
        text = re.sub(r"<[^>]+>", " ", text)
        text = re.sub(r"\s+", " ", text)
        hits = re.findall(r"(\d[\d,.]*\s?[KkMm]?\+?)\s+(travel agents|suppliers|source markets|"
                          r"accommodation|properties|hotels|clients|travelers|API partners|bookings)",
                          text)
        print(ts, hits[:20])


if __name__ == "__main__":
    main()
