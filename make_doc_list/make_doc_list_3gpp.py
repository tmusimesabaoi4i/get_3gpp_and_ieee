import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
)

from generate_html_url_from_xlsx.html_url_ieee import (
    get_html_url_ieee,
)

from make_doc_list.extract_meeting_excel_url import (
    extract_meeting_excel_url,
)

from util import (
    build_case_folder_from_excel,
    get_3gpp_html_name,
)

from pure_download.download_util import (
    get_landing_and_session,
)

from pure_download.download_file import (
    download_file_safely_msxml2,
)

from folder_and_file.delete_if_exists import (
    delete_if_exists,
    )

from generate_html_url_from_xlsx.html_url_3gpp import (
    get_html_url_3gpp,
    )

from folder_and_file.file_exists_in_folder import (
    file_exists_in_folder,
    )

from make_doc_list.make_doclist_from_3gpp_excel import (
    make_doclist_from_3gpp_excel
)

from make_doc_list.clean_excel_links_and_dates import (
    clean_excel_links_and_dates
)

# <処理>

folder_abs_path = r"C:\Users\yohei\Downloads"
filename = "input_3gpp.xlsx"

base_path = build_case_folder_from_excel(
    folder_abs_path = folder_abs_path,
    filename = filename,
    sheet = 0,
    # drive = drive,
    )

url_3gpp = get_html_url_3gpp(
    folder_abs_path = folder_abs_path,
    filename = filename,
    )

html_file_name_3gpp = get_3gpp_html_name(url_3gpp)

create_subfolder_when_absent(base_path, "HTML")

html_path = Path(base_path) / "HTML"

html_path_3gpp = Path(html_path) / html_file_name_3gpp

# HTMLpathから、エクセルを探す
excel_url = extract_meeting_excel_url(html_path_3gpp)

excel_path = Path(base_path) / "EXCEL"

create_subfolder_when_absent(base_path, "EXCEL")

# エクセルダウンロード
LANDING, sess = get_landing_and_session("3gpp")

if not file_exists_in_folder(excel_path, "tmp.xlsx"):
    ext = download_file_safely_msxml2(
    excel_url,
    excel_path,
    "tmp.xlsx",
    session=sess,
    referer=LANDING,
    # proxy=proxy,
    connect_timeout=10,
    read_timeout=180,
    max_retries=5,
    use_curl_fallback=True,
    )


# エクセルから、データを取ってくる
if not file_exists_in_folder(excel_path, "3gpp_doc_list.xlsx"):
    make_doclist_from_3gpp_excel(excel_path,"tmp.xlsx",excel_path,"3gpp_doc_list.xlsx")


# DOCLIST作成 
if file_exists_in_folder(excel_path, "3gpp_doc_list.xlsx"):
    clean_excel_links_and_dates(excel_path, "3gpp_doc_list.xlsx")
    
delete_if_exists(excel_path,"tmp.xlsx")