import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import re
import pandas as pd
from pathlib import Path
from html import unescape
from typing import Optional

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

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
)

from generate_html_url_from_xlsx.html_url_ieee import (
    get_html_url_ieee,
)

from util import (
    build_case_folder_from_excel,
    get_ieee_html_name,
)

def download_html_for_ieee(folder_abs_path: str, filename: str, drive: Optional[str] = None, proxy: Optional[str] = None) -> None:
    p = Path(folder_abs_path) / filename
    if not p.is_absolute():
        raise ValueError(f"{emo.warn} folder_abs_path は絶対パスで指定してください。")
    if not p.exists():
        raise FileNotFoundError(f"{emo.warn} Excel ファイルが見つかりません: {p}")

    suffix = p.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        engine = "openpyxl"
    elif suffix == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"{emo.warn} 未対応の拡張子です: {suffix}（.xlsx / .xlsm / .xls）")

    base_path = build_case_folder_from_excel(
        folder_abs_path = folder_abs_path,
        filename = filename,
        sheet = 0,
        drive = drive,
    )

    url_ieee_list = get_html_url_ieee(
        folder_abs_path = folder_abs_path,
        filename = filename,
    )

    html_file_name_ieee_list = get_ieee_html_name(url_ieee_list)

    create_subfolder_when_absent(base_path, "HTML")

    html_path = Path(base_path) / "HTML"

    LANDING, sess = get_landing_and_session("IEEE")
    
    for url_ieee, html_file_name_ieee in zip(url_ieee_list, html_file_name_ieee_list):
        if not file_exists_in_folder(html_path, html_file_name_ieee):
            ext = download_html_safely_msxml2(
                url_ieee,
                html_path,
                html_file_name_ieee,
                session=sess,
                referer=LANDING,
                proxy=proxy,
                connect_timeout=10,
                read_timeout=180,
                max_retries=5,
            )

if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_ieee.xlsx"

    download_html_for_ieee(folder_abs_path, filename)
