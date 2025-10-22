import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path
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

from condition_row.condition_row_3gpp import (
    condition_row_3gpp,
)

from get_info_from_doclist.get_url import (
    get_url,
)

from get_info_from_doclist.get_name_3gpp import (
    get_name_3gpp,
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

def fetch_3gpp_docs(folder_abs_path: str, filename: str, proxy: Optional[str]) -> List[Dict[str, Any]]:
    """
    3GPPのDocListに基づいてZIPをダウンロードするワークフローを実行する関数。

    Args:
        proxy (str|None): 例 "http://proxy.example.com:8080"。None/空なら未使用。
        folder_abs_path (str): 入力Excelが置いてあるフォルダの絶対パス。
        filename (str): 入力Excelファイル名 (例: "input_3gpp.xlsx")。

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
    LANDING, sess = get_landing_and_session("3gpp")  # type: ignore[name-defined]

    # ===== DocListからダウンロード対象抽出 =====
    rows = condition_row_3gpp(folder_abs_path, filename)  # type: ignore[name-defined]
    download_urls = get_url(excel_dir, "3gpp_doc_list.xlsx", rows)  # type: ignore[name-defined]
    download_filenames = get_name_3gpp(excel_dir, "3gpp_doc_list.xlsx", rows)  # type: ignore[name-defined]

    total = len(download_urls)
    if total == 0:
        print(f"{emo.warn} ダウンロード対象がありません。")  # type: ignore[name-defined]
        return []

    # ===== 進行用カウンタ =====
    rep = 0

    results: List[Dict[str, Any]] = []

    # ===== ダウンロードループ =====
    for i, (download_url, download_filename) in enumerate(zip(download_urls, download_filenames), start=1):
        rep += 1  # 進行カウント
        try:
            row_key = str(rows[rep - 1])  # 行識別用（例: 文献番号）
        except Exception:
            # rowsの長さ不一致などがあっても安全に進める
            row_key = str(rep)

        download_url = str(download_url).strip()
        basename = Path(str(download_filename).strip()).name  # パス成分除去
        stem = f"{row_key}_{basename}"
        zip_name = stem + ".zip"

        # 既存チェック
        if file_exists_in_folder(download_path, zip_name):  # type: ignore[name-defined]
            pct = round(rep / total * 100)
            print(f"[{pct}%] {emo.ok}  既存: {download_path / zip_name}")  # type: ignore[name-defined]
            results.append({
                "index": rep,
                "row": row_key,
                "url": download_url,
                "filename": basename,
                "download_path": str(download_path),
                "name": str(zip_name),
                "saved_path": str(download_path / zip_name),
                "ext": ".zip",
                "skipped": True,
            })
            continue

        # 実ダウンロード
        try:
            ext = download_file_safely_msxml2(  # type: ignore[name-defined]
                download_url,
                download_path,
                stem,  # 拡張子は関数側判定
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
                "name": str(zip_name),
                "saved_path": str(download_path / zip_name),
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
            # 上位で拾いたいケースが多い想定なので再送出
            raise

    return results