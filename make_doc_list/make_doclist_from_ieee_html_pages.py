from __future__ import annotations
from pathlib import Path
from typing import Iterable, List, Tuple
import re
import pandas as pd

from bs4 import BeautifulSoup  # pip install beautifulsoup4

def _clean(s) -> str:
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ")
    s = re.sub(r"[\r\n\t]+", " ", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip()

def _read_local_html(p: Path) -> str:
    # UTF-8 を優先、ダメなら cp932 を試す
    try:
        return p.read_text(encoding="utf-8")
    except Exception:
        return p.read_text(encoding="cp932", errors="replace")

def _abspath_ieee(href: str, base: str = "https://mentor.ieee.org") -> str:
    href = _clean(href)
    if href.startswith("/"):
        return base + href
    return href

def _extract_dcn_from_tds(tds: list[str]) -> Tuple[str, str]:
    """
    IEEEの行の列順（例）:
      0: 掲載日(div.date_time), 1: 年(2024), 2: 文献番号(256), 3: 版(4),
      4: TG, 5: タイトル, 6: 著者, 7: アップロード日時(div.date_time), 8: Download
    """
    doc_no = _clean(tds[2]) if len(tds) > 2 else ""
    ver    = _clean(tds[3]) if len(tds) > 3 else ""
    # 0256 → 256 / 04 → 4
    if doc_no:
        doc_no = doc_no.lstrip("0") or "0"
    if ver:
        ver = ver.lstrip("0") or "0"
    return doc_no, ver

def _extract_dcn_from_href(href: str) -> Tuple[str, str]:
    # 例: /802.11/dcn/24/11-24-0256-04-00be-....docx → 256 / 4
    h = href.lower()
    m = re.search(r"11-\d{2}-(\d+)-(\d{2})-[0-9a-z]+", h, flags=re.I)
    if m:
        return (m.group(1).lstrip("0") or "0", m.group(2).lstrip("0") or "0")
    m2 = re.search(r"(\d+)-(\d{2})-[0-9a-z]+\.docx", h, flags=re.I)
    if m2:
        return (m2.group(1).lstrip("0") or "0", m2.group(2).lstrip("0") or "0")
    return "", ""

def make_doclist_from_ieee_html_pages(
    html_path: str | Path,
    html_file_name_ieee_list: Iterable[str],
    filepath_out: str | Path,
    filename_out: str,
) -> Path:
    """
    ローカルの IEEE Mentor 一覧 HTML 群から、DocList(IEEE) を作成し保存する。

    出力列: URL / 著者 / タイトル / 日付 / 文献番号 / バージョン
      - URL: 行末の Download リンクの href。先頭が '/' なら https://mentor.ieee.org を付与
      - 著者: 行の 7列目 (td.long)
      - タイトル: 行の 6列目 (td.long)
      - 日付: 行内の div.date_time の **最後**（アップロード日時）
      - 文献番号・バージョン: できれば 3列目・4列目を優先、無ければ href から抽出
    """
    html_dir = Path(html_path)
    out_dir = Path(filepath_out)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / filename_out

    rows: List[dict] = []

    for fname in html_file_name_ieee_list:
        fp = html_dir / fname
        if not fp.exists():
            # 見つからない場合はスキップ（堅牢性重視）
            continue

        html = _read_local_html(fp)
        soup = BeautifulSoup(html, "html.parser")

        # データ行
        tr_list = soup.select("tr.b_data_row") or soup.select("tr")

        for tr in tr_list:
            # Download リンク
            a = tr.select_one("td.list_actions a[href]") or \
                (tr.select("a[href]")[-1] if tr.select("a[href]") else None)
            if not a:
                continue
            href = a.get("href", "").strip()
            if not href:
                continue
            url_abs = _abspath_ieee(href)

            # 列テキスト
            tds = [_clean(td.get_text()) for td in tr.find_all("td")]
            if len(tds) < 6:
                continue

            # タイトル・著者
            title  = tds[5] if len(tds) > 5 else ""
            author = tds[6] if len(tds) > 6 else ""

            # 日付（Uploadedの方＝最後のdate_time）
            dt_divs = tr.select("div.date_time")
            date_str = _clean(dt_divs[-1].get_text()) if dt_divs else ""

            # 文献番号・版：列から取得、無ければ href から抽出
            doc_no, ver = _extract_dcn_from_tds(tds)
            if not doc_no or not ver:
                d2, v2 = _extract_dcn_from_href(href)
                doc_no = doc_no or d2
                ver    = ver or v2

            rows.append({
                "URL": url_abs,
                "著者": author,
                "タイトル": title,
                "日付": date_str,
                "文献番号": doc_no,
                "バージョン": ver,
            })

    df = pd.DataFrame(rows, columns=["URL", "著者", "タイトル", "日付", "文献番号", "バージョン"])

    # 最終トリムと空行除去
    for c in df.columns:
        df[c] = df[c].map(_clean)
    if not df.empty:
        df = df[~(df.eq("").all(axis=1))].reset_index(drop=True)

    df.to_excel(out_path, index=False)
    return out_path
