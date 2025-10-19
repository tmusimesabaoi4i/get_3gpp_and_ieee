import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path
from typing import Optional

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
    )

from folder_and_file.delete_if_exists import (
    delete_if_exists,
    )

from folder_and_file.file_exists_in_folder import (
    file_exists_in_folder,
    )

from make_doc_list.extract_meeting_excel_url import (
    extract_meeting_excel_url,
    )

from make_doc_list.make_doclist_from_ieee_html_pages import (
    make_doclist_from_ieee_html_pages
    )

from make_doc_list.clean_excel_links_and_dates import (
    clean_excel_links_and_dates
    )

from generate_html_url_from_xlsx.html_url_ieee import (
    get_html_url_ieee,
    )

from util import (
    build_case_folder_from_excel,
    get_ieee_html_name,
    )

def make_doc_list_ieee(
        folder_abs_path: str,
        filename: str,
        drive: Optional[str] = None,
    ) -> None:
    # 1) ケースフォルダ決定
    base_path = build_case_folder_from_excel(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path,
        filename=filename,
        sheet=0,
        drive=drive,
    )
    base_path = Path(base_path)

    # 2) IEEE 一覧ページURLと、保存済みHTMLのファイル名リスト
    url_ieee_list = get_html_url_ieee(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path,
        filename=filename,
    )
    html_file_name_ieee_list = get_ieee_html_name(url_ieee_list)  # type: ignore[name-defined]
    if not html_file_name_ieee_list:
        raise ValueError("IEEE のHTMLファイル名リストが空です。入力ExcelやURLの指定を確認してください。")

    # 3) サブフォルダ作成
    create_subfolder_when_absent(base_path, "HTML")   # type: ignore[name-defined]
    create_subfolder_when_absent(base_path, "EXCEL")  # type: ignore[name-defined]

    html_path  = base_path / "HTML"
    excel_path = base_path / "EXCEL"

    # 4) HTMLから DocList(IEEE) を作成
    out_filename = "ieee_doc_list.xlsx"
    if not file_exists_in_folder(excel_path, out_filename):  # type: ignore[name-defined]
        make_doclist_from_ieee_html_pages(  # type: ignore[name-defined]
            html_path,
            html_file_name_ieee_list,
            excel_path,
            out_filename,
        )

    # 5) DocList の A列ハイパーリンク除去＆D列日付 'YYYY/MM/DD' 正規化
    if file_exists_in_folder(excel_path, out_filename):  # type: ignore[name-defined]
        clean_excel_links_and_dates(  # type: ignore[name-defined]
            excel_path,
            out_filename,
        )

if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_ieee.xlsx"

    print("🚀 IEEE パイプライン開始...")
    make_doc_list_ieee(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path,
        filename=filename,
        # drive=args.drive,
        # proxy=args.proxy,
    )