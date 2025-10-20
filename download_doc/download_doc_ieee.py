import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import os
from pathlib import Path
from urllib.parse import urlparse

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

folder_abs_path = r"C:\Users\yohei\Downloads"
filename = "input_ieee.xlsx"


base_path = build_case_folder_from_excel(  # type: ignore[name-defined]
    folder_abs_path=folder_abs_path, filename=filename, sheet=0, drive=None
)

base_path = Path(base_path)

create_subfolder_when_absent(base_path, "EXCEL")  # type: ignore[name-defined]

excel_dir = base_path / "EXCEL"

create_subfolder_when_absent(base_path, "DOCS")

download_path = base_path / "DOCS"

# ===== セッション取得 =====
LANDING, sess = get_landing_and_session("ieee")

rows = condition_row_ieee(folder_abs_path, filename)
download_urls = get_url(excel_dir, "ieee_doc_list.xlsx", rows)
download_filenames = get_name_ieee(excel_dir, "ieee_doc_list.xlsx", rows)

total = len(download_urls)

rep = 0  # 進行カウンタ（任意用途）
for i, (download_url, download_filename) in enumerate(zip(download_urls, download_filenames), start=1):
    rep += 1  # 進行カウント

    parsed = urlparse(download_url)
    pure_filename = os.path.basename(parsed.path) or "download.bin"
    file_extension = os.path.splitext(pure_filename)[1]

    if not file_exists_in_folder(download_path, str(rows[rep-1]) + "_" + download_filename + file_extension):
        download_url = str(download_url).strip()
        download_filename = Path(str(download_filename).strip()).name  # パス成分除去
        try:
            ext = download_file_safely_msxml2(
                download_url,
                download_path,
                str(rows[rep-1]) + "_" + download_filename,
                session=sess,
                referer=LANDING,
                # proxy="http://proxy.example.com:8080",
                connect_timeout=10,
                read_timeout=180,
                max_retries=5,
                use_curl_fallback=True,
            )
            print(f"[{round(rep/total*100)}%] {emo.ok}  拡張子: {ext or '(不明)'}  保存先: {download_path / Path(str(rows[rep-1]) + "_" + download_filename)}")
        except Exception as e:
            print(f"{emo.warn} エラー: {e}")
            raise