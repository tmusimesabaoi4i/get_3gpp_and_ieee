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

from pure_download.download_util import (
    get_landing_and_session,
    )

from pure_download.download_file import (
    download_file_safely_msxml2,
    )

from make_doc_list.extract_meeting_excel_url import (
    extract_meeting_excel_url,
    )

from make_doc_list.make_doclist_from_3gpp_excel import (
    make_doclist_from_3gpp_excel
    )

from make_doc_list.clean_excel_links_and_dates import (
    clean_excel_links_and_dates
    )

from generate_html_url_from_xlsx.html_url_3gpp import (
    get_html_url_3gpp,
    )

from util import (
    build_case_folder_from_excel,
    get_3gpp_html_name,
    )

def make_doc_list_3gpp(
        folder_abs_path: str,
        filename: str,
        drive: Optional[str] = None,
        proxy: Optional[str] = None,
    ) -> None:
    # 1) ケースフォルダの決定
    base_path = build_case_folder_from_excel(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path,
        filename=filename,
        sheet=0,
        drive=drive,
    )
    base_path = Path(base_path)

    # 2) 3GPP の一覧HTMLのURL・ファイル名を取得
    url_3gpp = get_html_url_3gpp(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path,
        filename=filename,
    )
    html_file_name_3gpp = get_3gpp_html_name(url_3gpp)  # type: ignore[name-defined]

    # 3) HTML サブフォルダ作成 & 期待パスを決定
    create_subfolder_when_absent(base_path, "HTML")  # type: ignore[name-defined]
    html_path = base_path / "HTML"
    html_path_3gpp = html_path / html_file_name_3gpp

    # 4) HTML から会合 Excel の URL を取得
    excel_url = extract_meeting_excel_url(html_path_3gpp)  # type: ignore[name-defined]
    if not excel_url:
        raise ValueError(f"会合ExcelのURL取得に失敗: {html_path_3gpp}")

    # 5) EXCEL サブフォルダ作成
    create_subfolder_when_absent(base_path, "EXCEL")  # type: ignore[name-defined]
    excel_path = base_path / "EXCEL"

    # 6) エクセルのダウンロード（tmp.xlsx）
    LANDING, sess = get_landing_and_session("3gpp")  # type: ignore[name-defined]
    try:
        if not file_exists_in_folder(excel_path, "tmp.xlsx"):  # type: ignore[name-defined]
            _ = download_file_safely_msxml2(  # type: ignore[name-defined]
                excel_url,
                excel_path,
                "tmp.xlsx",
                session=sess,
                referer=LANDING,
                proxy=proxy,
                connect_timeout=10,
                read_timeout=180,
                max_retries=5,
                use_curl_fallback=True,
            )

        # 7) DocList 生成
        if not file_exists_in_folder(excel_path, "3gpp_doc_list.xlsx"):  # type: ignore[name-defined]
            make_doclist_from_3gpp_excel(  # type: ignore[name-defined]
                excel_path,
                "tmp.xlsx",
                excel_path,
                "3gpp_doc_list.xlsx",
            )

        # 8) DocList のリンク除去＆日付整形
        if file_exists_in_folder(excel_path, "3gpp_doc_list.xlsx"):  # type: ignore[name-defined]
            clean_excel_links_and_dates(  # type: ignore[name-defined]
                excel_path,
                "3gpp_doc_list.xlsx",
            )
    finally:
        # 9) 後片付け（tmp.xlsx 削除・存在すれば）
        delete_if_exists(excel_path, "tmp.xlsx")  # type: ignore[name-defined]

if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_3gpp.xlsx"

    print("🚀 3GPP パイプライン開始...")
    make_doc_list_3gpp(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path,
        filename=filename,
        # drive=args.drive,
        # proxy=args.proxy,
    )
