from __future__ import annotations
import re
from html import unescape
from pathlib import Path
from typing import Optional, Union
from urllib.parse import urljoin

# ---- <tr> ブロック抽出 ----
_TR_RE = re.compile(r"<tr\b[^>]*>[\s\S]*?</tr>", re.I)

# ---- input: class="downloadInput" かつ value が TDoc_List_Meeting で始まる ----
_INPUT_MARK_RE = re.compile(
    r'<input\b'
    r'(?=[^>]*\bclass\s*=\s*(?:"[^"]*\bdownloadInput\b[^"]*"|\'[^\']*\bdownloadInput\b[^\']*\'))'
    r'(?=[^>]*\bvalue\s*=\s*(?:"TDoc_List_Meeting[^"]*"|\'TDoc_List_Meeting[^\']*\'))'
    r'[^>]*>',
    re.I
)

# ---- <a class="file" href="...TDoc_List_Meeting....xlsx"> を優先 ----
_A_FILE_XLSX_RE = re.compile(
    r'<a\b'
    r'(?=[^>]*\bclass\s*=\s*(?:"[^"]*\bfile\b[^"]*"|\'[^\']*\bfile\b[^\']*\'))'
    r'(?=[^>]*\bhref\s*=\s*(?:"([^"]*TDoc_List_Meeting[^"]*\.xlsx[^"]*)"|\'([^\']*TDoc_List_Meeting[^\']*\.xlsx[^\']*)\'))'
    r'[^>]*>',
    re.I
)

# ---- フォールバック: <a href="...TDoc_List_Meeting....xlsx">（class 指定なし）----
_A_ANY_XLSX_RE = re.compile(
    r'<a\b(?=[^>]*\bhref\s*=\s*(?:"([^"]*TDoc_List_Meeting[^"]*\.xlsx[^"]*)"|\'([^\']*TDoc_List_Meeting[^\']*\.xlsx[^\']*)\'))[^>]*>',
    re.I
)

def _read_html_input(html_or_path: Union[str, bytes, Path], encoding: Optional[str] = None) -> str:
    """文字列/bytes/Path を受け取り、HTML文字列にして返す。"""
    # bytes -> decode
    if isinstance(html_or_path, (bytes, bytearray)):
        if encoding:
            return html_or_path.decode(encoding, errors="replace")
        # 自動判定（簡易）：UTF-8→CP932 の順で試す
        for enc in ("utf-8", "cp932"):
            try:
                return html_or_path.decode(enc)
            except Exception:
                pass
        return html_or_path.decode("utf-8", errors="replace")

    # Path or 既存ファイルパス文字列 -> ファイル読み
    if isinstance(html_or_path, Path) or (isinstance(html_or_path, str) and Path(html_or_path).exists()):
        p = Path(html_or_path)
        if encoding:
            return p.read_text(encoding=encoding, errors="replace")
        # 自動判定（簡易）：UTF-8→CP932
        for enc in ("utf-8", "cp932"):
            try:
                return p.read_text(encoding=enc)
            except Exception:
                pass
        return p.read_text(encoding="utf-8", errors="replace")

    # ここまでで Path でも bytes でもない -> 文字列として扱う
    s = str(html_or_path)
    return s

def _sanitize(html: str) -> str:
    """HTML を最小正規化（エンティティ解除、コメント/スクリプト/スタイル除去）"""
    s = unescape(html)
    s = re.sub(r"<!--.*?-->", "", s, flags=re.S)
    s = re.sub(r"<script\b[^>]*>.*?</script>", "", s, flags=re.I | re.S)
    s = re.sub(r"<style\b[^>]*>.*?</style>", "", s, flags=re.I | re.S)
    return s

def extract_meeting_excel_url(
    html_or_path: Union[str, bytes, Path],
    base_url: Optional[str] = None,
    encoding: Optional[str] = None,
) -> Optional[str]:
    """
    <tr> 行の中に
      <input class="downloadInput" value="TDoc_List_Meeting...">
    を含むものを見つけ、その同じ行の <a ... href="...xlsx"> の URL を返す。
    - html_or_path: HTML 文字列 / bytes / ファイルパス（str or Path）
    - base_url: 相対URLを絶対化したいときに指定（例: "https://www.3gpp.org"）
    - encoding: ファイル読込時の明示エンコーディング。未指定なら簡易自動判定。
    """
    raw = _read_html_input(html_or_path, encoding=encoding)
    s = _sanitize(raw)

    for tr in _TR_RE.findall(s):
        if not _INPUT_MARK_RE.search(tr):
            continue

        # 1) <a class="file" ... .xlsx> を優先
        m = _A_FILE_XLSX_RE.search(tr)
        if not m:
            # 2) フォールバック: class なしでも .xlsx を含む <a> を拾う
            m = _A_ANY_XLSX_RE.search(tr)
        if not m:
            continue

        href = (m.group(1) or m.group(2) or "").strip()
        if not href:
            continue

        # 絶対化
        if base_url and not re.match(r'^[a-z][a-z0-9+.\-]*://', href, flags=re.I):
            href = urljoin(base_url, href)
        return href

    return None
