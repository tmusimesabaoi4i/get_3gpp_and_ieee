import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import os
from pathlib import Path
from urllib.parse import urlparse
from typing import Optional,List,Dict,Any
from pure_download.download_file import (
    download_file_safely_msxml2,
)

from pure_download.download_util import (
    # cookie_header_from_session,
    get_landing_and_session,
    # is_dir_like,
    # normalize_proxy_for_msxml2,
    # sanitize_filename,
    # to_double_backslash_literal,
)

from condition_row.condition_row_ieee import (
    condition_row_ieee,
)

from get_info_from_doclist.get_url import (
    get_url,
)

from get_info_from_doclist.get_name_ieee import (
    get_name_ieee,
)


from util import (
    build_case_folder_from_excel,
    )

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
    )
from folder_and_file.folder_exists_in_folder import (
    folder_exists_in_folder,
    )
from folder_and_file.file_exists_in_folder import (
    file_exists_in_folder,
    )
    
def fetch_ieee_docs(folder_abs_path: str, filename: str, proxy: Optional[str]) -> List[Dict[str, Any]]:
    """
    IEEE DocListに基づいてファイルをダウンロードするワークフロー。

    Args:
        proxy (str|None): 例 "http://proxy.example.com:8080"。None/空なら未使用。
        folder_abs_path (str): 入力Excelが置いてあるフォルダの絶対パス。
        filename (str): 入力Excelファイル名 (例: "input_ieee.xlsx")。

    Returns:
        List[Dict[str,Any]]: 各アイテムの処理結果（保存先、拡張子、スキップ有無など）。
    """
    # ===== ルート/下位フォルダの用意 =====
    base_path = build_case_folder_from_excel(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path, filename=filename, sheet=0, drive=None
    )
    base_path = Path(base_path)

    create_subfolder_when_absent(base_path, "EXCEL")  # type: ignore[name-defined]
    excel_dir = base_path / "EXCEL"

    create_subfolder_when_absent(base_path, "DOCS")   # type: ignore[name-defined]
    download_path = base_path / "DOCS"

    # ===== セッション取得 =====
    LANDING, sess = get_landing_and_session("ieee")  # type: ignore[name-defined]

    # ===== DocListからダウンロード対象抽出 =====
    rows = condition_row_ieee(folder_abs_path, filename)  # type: ignore[name-defined]
    download_urls = get_url(excel_dir, "ieee_doc_list.xlsx", rows)  # type: ignore[name-defined]
    download_filenames = get_name_ieee(excel_dir, "ieee_doc_list.xlsx", rows)  # type: ignore[name-defined]

    total = len(download_urls)
    if total == 0:
        print(f"{emo.warn} ダウンロード対象がありません。")  # type: ignore[name-defined]
        return []

    rep = 0
    results: List[Dict[str, Any]] = []

    # ===== ダウンロードループ =====
    for i, (download_url, download_filename) in enumerate(zip(download_urls, download_filenames), start=1):
        rep += 1

        try:
            row_key = str(rows[rep - 1])  # 行識別（文献番号など）
        except Exception:
            row_key = str(rep)

        # 文字列/名前を正規化
        download_url = str(download_url).strip()
        basename = Path(str(download_filename).strip()).name  # パス成分除去
        stem = f"{row_key}_{basename}"

        # URL末尾から拡張子を推定（存在チェック用）
        parsed = urlparse(download_url)
        pure_filename = os.path.basename(parsed.path) or "download.bin"
        ext_guess = os.path.splitext(pure_filename)[1] or ".bin"
        exist_name = stem + ext_guess

        # 既存チェック（拡張子推定ベース）
        if file_exists_in_folder(download_path, exist_name):  # type: ignore[name-defined]
            pct = round(rep / total * 100)
            print(f"[{pct}%] {emo.ok}  既存: {download_path / exist_name}")  # type: ignore[name-defined]
            results.append({
                "index": rep,
                "row": row_key,
                "url": download_url,
                "filename": basename,
                "download_path": str(download_path),
                "name": str(exist_name),
                "saved_path": str(download_path / exist_name),
                "ext": ext_guess,
                "skipped": True,
            })
            continue

        # 実ダウンロード（拡張子は関数側で最終決定）
        try:
            ext = download_file_safely_msxml2(  # type: ignore[name-defined]
                download_url,
                download_path,
                stem,  # 保存名の“本体”、拡張子は関数が決める
                session=sess,
                referer=LANDING,
                proxy=(proxy if proxy else None),
                connect_timeout=10,
                read_timeout=180,
                max_retries=5,
                use_curl_fallback=True,
            )
            pct = round(rep / total * 100)
            saved = download_path / stem
            print(f"[{pct}%] {emo.ok}  拡張子: {ext or '(不明)'}  保存先: {saved}")  # type: ignore[name-defined]
            results.append({
                "index": rep,
                "row": row_key,
                "url": download_url,
                "filename": basename,
                "download_path": str(download_path),
                "name": str(stem),
                "saved_path": str(download_path / stem),
                "ext": ext,
                "skipped": False,
            })
        except Exception as e:
            print(f"{emo.warn} エラー: {e}")  # type: ignore[name-defined]
            results.append({
                "index": rep,
                "row": row_key,
                "url": download_url,
                "filename": basename,
                "saved_path": None,
                "ext": None,
                "skipped": False,
                "error": str(e),
            })
            raise  # 上位で拾いたい前提

    return results