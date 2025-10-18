import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path

from typing import Union, Optional
from typing import Iterable, Union, List, overload

import os
import re
import pandas as pd

from urllib.parse import urlparse, parse_qs



from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
)

def _sanitize_token(s: str) -> str:
    """
    ファイル名用に安全化（空白を削除、禁止文字を '_' に）。
    """
    s = (s or "").strip()
    s = re.sub(r'\.html?$', '', s, flags=re.I)           # 末尾の .html / .htm は除去
    s = re.sub(r'[\\/:*?"<>|\s]+', "_", s)               # 禁止/空白 -> _
    return s

def _change_drive(base: Path, drive: str) -> Path:
    d = (drive or "").strip().rstrip(":").upper()
    if not d or len(d) != 1 or not d.isalpha():
        raise ValueError(f"Invalid drive: {drive!r}")
    s = str(base)
    s2 = re.sub(r"^[A-Za-z]:", f"{d}:", s)
    return Path(s2)

def _sanitize_filename(s: str) -> str:
    s = (s or "").strip()
    return re.sub(r'[\\/:\*\?"<>\|]+', "_", s)

def cell(df, r: int, c: int) -> str:
    try:
        v = df.iat[r, c]
    except Exception:
        v = None
    return "" if (v is None or pd.isna(v)) else str(v).strip()

def get_downloads_path(drive: Optional[str] = None) -> Path:
    home = Path.home()
    names = ["Downloads", "downloads", "Download", "ダウンロード"]

    # 1) まずは現在のホーム配下で探索
    for name in names:
        p = (home / name)
        if p.exists() and p.is_dir():
            return p.resolve()

    # 2) 見つからなければ、指定ドライブに差し替えて再探索（Windows想定）
    if drive and (home.drive or os.name == "nt"):
        alt_home = _change_drive(home, drive)
        for name in names:
            p = (alt_home / name)
            if p.exists() and p.is_dir():
                return p.resolve()

    # 3) どれも無ければエラー
    tried = [str(home / n) for n in names]
    if drive and (home.drive or os.name == "nt"):
        tried += [str(_change_drive(home, drive) / n) for n in names]
    raise FileNotFoundError(f"{emo.warn} Downloads folder not found. Tried: " + ", ".join(tried))

def build_case_folder_from_excel(
    folder_abs_path: str, filename: str,
    sheet: Union[int, str] = 0,
    drive: Optional[str] = None,
    ) -> Path:

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
        dtype=str
    )

    date = _sanitize_filename(cell(df, 0, 1))     # B1
    case_id = _sanitize_filename(cell(df, 1, 1))  # B2

    if not date or not case_id:
        raise ValueError(f"{emo.warn} B1/B2 が空です。date='{date}' case_id='{case_id}'")

    downloads = get_downloads_path(drive=drive)

    downloads_filename = str(case_id)+'_'+str(date)

    create_subfolder_when_absent(downloads, downloads_filename)

    target = Path(downloads) / downloads_filename

    return target.resolve()

def _get_3gpp_html_name_single(url: str) -> str:
    if url is None:
        raise ValueError(f"{emo.warn} URL が None です。")
    url = str(url).strip()
    if not url:
        raise ValueError(f"{emo.warn} URL が空文字です。")

    parsed = urlparse(url)
    segments = [s for s in parsed.path.split('/') if s]  # 空を除外

    if not segments:
        raise ValueError(f"{emo.warn} URLにパスがありません: {url}")

    if segments[-1].lower() == "docs":
        if len(segments) < 2:
            raise ValueError(f"{emo.warn} Docs の前にシリーズ名が見つかりません: {url}")
        series = segments[-2]
    else:
        series = segments[-1]

    series = re.sub(r'[\\/:*?"<>|]+', "_", series).strip()
    series = re.sub(r'\.html?$', '', series, flags=re.I)

    if not series:
        raise ValueError(f"{emo.warn} シリーズ名を抽出できませんでした: {url}")

    series_s = _sanitize_token(series)
    return f"{series_s}.html"

@overload
def get_3gpp_html_name(urls: str) -> str: ...
@overload
def get_3gpp_html_name(urls: Iterable[str]) -> List[str]: ...
def get_3gpp_html_name(urls: Union[str, Iterable[str]]) -> Union[str, List[str]]:
    if isinstance(urls, str):
        return _get_3gpp_html_name_single(urls)

    results: List[str] = []
    for i, u in enumerate(urls):
        try:
            results.append(_get_3gpp_html_name_single(u))
        except Exception as e:
            raise ValueError(f"{emo.warn} {i} 番目のURLでエラー: {e}") from e
    return results

def _get_ieee_html_name_single(url: str) -> str:
    if url is None:
        raise ValueError(f"{emo.warn} URL が None です。")
    url = str(url).strip()
    if not url:
        raise ValueError(f"{emo.warn} URL が空文字です。")

    parsed = urlparse(url)
    segments = [s for s in parsed.path.split('/') if s]  # 空を除外

    if not segments:
        raise ValueError(f"{emo.warn} URLにパスがありません: {url}")

    q = parse_qs(parsed.query, keep_blank_values=True)
    q = {k.lower(): v for k, v in q.items()}  # キーを小文字化して頑健に

    def first(*keys: str) -> str | None:
        for k in keys:
            if k in q and len(q[k]) > 0:
                return q[k][0]
        return None

    year  = first("is_year", "year")
    group = first("is_group", "group")
    n     = first("n")

    if not year:
        raise ValueError(f"is_year/year が見つかりません: {url}")
    if not group:
        raise ValueError(f"is_group/group が見つかりません: {url}")
    if not n:
        raise ValueError(f"n が見つかりません: {url}")

    year_s  = _sanitize_token(year)
    group_s = _sanitize_token(group)
    n_s     = _sanitize_token(n)

    return f"{year_s}_{group_s}_{n_s}.html"

@overload
def get_ieee_html_name(urls: str) -> str: ...
@overload
def get_ieee_html_name(urls: Iterable[str]) -> List[str]: ...
def get_ieee_html_name(urls: Union[str, Iterable[str]]) -> Union[str, List[str]]:
    if isinstance(urls, str):
        return _get_ieee_html_name_single(urls)

    results: List[str] = []
    for i, u in enumerate(urls):
        try:
            results.append(_get_ieee_html_name_single(u))
        except Exception as e:
            raise ValueError(f"{emo.warn} {i} 番目のURLでエラー: {e}") from e
    return results