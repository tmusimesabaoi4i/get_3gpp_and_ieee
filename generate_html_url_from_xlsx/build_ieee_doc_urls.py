import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import re
from urllib.parse import urlsplit, urlunsplit
from typing import List

def build_ieee_doc_urls_dirty(
        val: str,
        n_last: int
    ) -> List[str]:
    parts = urlsplit(val)
    q = parts.query
    m = re.match(r"^n=(\d+)(.*)$", q)
    if not m:
        raise ValueError(f"{emo.warn} 想定外のクエリ形式です（例: n=1XXX を想定）: " + q)

    suffix_raw = m.group(2)
    suffix = suffix_raw.lstrip("&")
    suffix_part = f"&{suffix}" if suffix else ""

    base = urlunsplit((parts.scheme, parts.netloc, parts.path, "", parts.fragment))

    urls = [f"{base}?n={i}{suffix_part}" for i in range(1, int(n_last) + 1)]
    return urls

if __name__ == "__main__":
    val = "https://mentor.ieee.org/802.11/documents?n=1&o=7d&is_group=00be&is_year=2024"
    n = 10
    nums = range(2, n + 1)  # [2..n]
    urls = build_ieee_doc_urls_dirty(val, max(nums))
    for u in urls:
        print(u)
