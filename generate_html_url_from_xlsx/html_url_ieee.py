import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import re
import pandas as pd
from pathlib import Path
from html import unescape
from typing import Tuple, List, Optional, Any, Union

from generate_html_url_from_xlsx.build_ieee_doc_urls import (
        build_ieee_doc_urls_dirty,
    )

from pure_download.download_util import (
        get_landing_and_session,
    )

from pure_download.download_html import (
        download_html_safely_msxml2,
    )

from folder_and_file.file_exists_in_folder import (
        file_exists_in_folder,
    )

from folder_and_file.delete_if_exists import (
        delete_if_exists,
    )

_MARKER_RE = re.compile( (
            r"</table></div>"
            r"<div[^>]*class=\"main_bottom\"[^>]*>"
            r"<div[^>]*class=\"tools\"[^>]*>"
            r"<div[^>]*class=\"task_menu\"[^>]*>"
            r"<a[^>]*href=\"(?:https?://mentor\.ieee\.org)?/802\.11/check-uploader\"[^>]*>"
            r"joingroup</a>"
        ),
        re.I,
    )

_DOCNUM_RE = re.compile(
    r"(?:https?://mentor\.ieee\.org)?/802\.11/documents\?n=(\d+)", re.I
    )

def _normalize_html_for_marker(
        s: str
    ) -> str:
    s = unescape(s)
    s = re.sub(r"<!--.*?-->", "", s, flags=re.S)
    s = re.sub(r"<script\b[^>]*>.*?</script>", "", s, flags=re.S | re.I)
    s = re.sub(r"<style\b[^>]*>.*?</style>", "", s, flags=re.S | re.I)
    s = s.replace("\u200b", "").replace("\ufeff", "").replace("\xa0", " ")
    s = re.sub(r"\s+", "", s).lower()
    return s

def find_pager_anchor_pos(
        html_path: Path,
    ) -> Tuple[int, int, str]:
    raw = html_path.read_text(encoding="utf-8", errors="ignore")
    compact_lower = _normalize_html_for_marker(raw)

    m = _MARKER_RE.search(compact_lower)
    if not m:
        return -1, -1, compact_lower

    return m.start(), m.end(), compact_lower

def _extract_ieee_document_page_numbers(
        folder_abs_path: str,
        filename: str,
    ) -> List[int]:
    p = Path(folder_abs_path) / filename
    if not p.is_absolute():
        raise ValueError(f"{emo.warn} folder_abs_path は絶対パスで指定してください。")
    if not p.exists():
        raise FileNotFoundError(f"{emo.warn} Excel ファイルが見つかりません: {p}")

    start, end, compact_lower = find_pager_anchor_pos(p)
    if start == -1:
        return []
    tail = compact_lower[end:]

    nums = [int(n) for n in _DOCNUM_RE.findall(tail)]
    return nums

def get_html_url_ieee(
        folder_abs_path: str,
        filename: str,
        sheet: Union[int, str] = 0,
        proxy: Optional[str] = None,
    ) -> Optional[Any]:
    excel_path = Path(folder_abs_path) / filename
    if not excel_path.is_absolute():
        raise ValueError(f"{emo.warn} folder_abs_path は絶対パスで指定してください。")
    if not excel_path.exists():
        raise FileNotFoundError(f"{emo.warn} Excel ファイルが見つかりません: {p}")

    suffix = excel_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        engine = "openpyxl"
    elif suffix == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"{emo.warn} 未対応の拡張子です: {suffix}（.xlsx / .xlsm / .xls）")
    df = pd.read_excel(
        excel_path,
        sheet_name=sheet,
        engine=engine,
        header=None,
        usecols=[1],
        skiprows=4,
        nrows=1,
        dtype=str
    )

    if df.empty:
        return None
    val = df.iat[0, 0]

    if pd.isna(val):
        return None 

    if not file_exists_in_folder(folder_abs_path, "tmp.html"):
        LANDING, sess = get_landing_and_session("IEEE")

        ext = download_html_safely_msxml2(
            val,
            folder_abs_path,
            "tmp",
            session=sess,
            referer=LANDING,
            proxy=proxy,
            connect_timeout=10,
            read_timeout=180,
            max_retries=5,
        )

    nums = _extract_ieee_document_page_numbers(folder_abs_path, "tmp.html")

    delete_if_exists(folder_abs_path, "tmp.html")

    if not nums:
        return val

    n_last = max(int(x) for x in nums)
    return build_ieee_doc_urls_dirty(val, n_last)

if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_ieee.xlsx"

    urls = get_html_url_ieee(folder_abs_path, filename)
    for u in urls:
        print(u)