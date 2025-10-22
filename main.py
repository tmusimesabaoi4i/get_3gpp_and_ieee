# main.py
from __future__ import annotations

import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from util import (
    get_proxy_from_cmd,
    get_downloads_path,
    build_case_folder_from_excel,
)

from download_html.html_for_3gpp import(
    download_html_for_3gpp
)

from download_html.html_for_ieee import(
    download_html_for_ieee
)

from make_doc_list.make_doc_list_3gpp import(
    make_doc_list_3gpp
)

from make_doc_list.make_doc_list_ieee import(
    make_doc_list_ieee
)

from download_doc.download_doc_3gpp import(
    fetch_3gpp_docs
)

from download_doc.download_doc_ieee import(
    fetch_ieee_docs
)

from download_doc.save_results_to_xlsx import(
    save_results_to_xlsx,
    write_res_zip_paths_to_xlsx,
    read_column_as_list
)

from about_zip.extract_zip_to_docs import(
    extract_zip_to_docs_from_fold
)

from combine.extract_paragraphs import(
    convert_office_to_html
)

# ========== 定数 ==========
VALID_DB = {"ieee", "3gpp"}
DRIVE_RE = re.compile(r"^[A-Za-z]$")

# ========== 設定/依存の型 ==========
@dataclass(frozen=True)
class Config:
    drive: str                 # 例: "C"
    db: str                    # "ieee" or "3gpp"
    excel_path: Path           # 絶対パス（存在・拡張子検証済）
    proxy_analysis: bool       # 既定 True
    verbose: int               # -v / -vv
    dry_run: bool              # --dry-run

# ========== ログ初期化 ==========
def setup_logging(verbosity: int) -> None:
    level = logging.WARNING if verbosity == 0 else (
        logging.INFO if verbosity == 1 else logging.DEBUG
    )
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

# ========== 引数型チェック ==========
def drive_type(s: str) -> str:
    s = (s or "").strip()
    if not DRIVE_RE.match(s):
        raise argparse.ArgumentTypeError(
            f"{emo.warn} ドライブレターは英字1文字で指定してください（例: {emo.folder} C / D / E）。"
        )
    return s.upper()

def db_type(s: str) -> str:
    s = (s or "").strip().lower()
    if s not in VALID_DB:
        raise argparse.ArgumentTypeError(
            f"{emo.warn} 解析対象DBは {emo.db} ieee または 3gpp を指定してください。"
        )
    return s

def excel_path_type(s: str) -> Path:
    p = Path(s)
    if not p.is_absolute():
        raise argparse.ArgumentTypeError(
            f"{emo.invalid} Excelパスは絶対パスで指定してください（例: {emo.excel} C:\\path\\to\\file.xlsx）。"
        )
    if not p.exists():
        raise argparse.ArgumentTypeError(
            f"{emo.ng} 指定したExcelファイルが存在しません：{p}"
        )
    if p.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise argparse.ArgumentTypeError(
            f"{emo.invalid} Excel拡張子ではありません（許可: .xlsx / .xlsm / .xls）。"
        )
    return p

# ========== CLI パーサ ==========
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="mytool",
        description=(
            f"{emo.info} 必須入力：\n"
            f"  • {emo.folder} ドライブレター（--drive）\n"
            f"  • {emo.db} 解析対象DB（--db = ieee | 3gpp）\n"
            f"  • {emo.excel} 入力Excelの絶対パス（--excel）\n\n"
            f"{emo.spark} オプション：\n"
            f"  • {emo.net} プロキシ分析（--proxy-analysis / --no-proxy-analysis, 既定=有効）\n"
            f"  • {emo.warn} ドライラン（--dry-run）\n"
            f"  • {emo.star} 冗長ログ（-v, -vv）"
        ),
        epilog=(
            f"{emo.thick}\n"
            f"例：\n"
            f"  python main.py --drive C --db ieee --excel C:\\path\\to\\file.xlsx\n"
            f"  python main.py --drive C --db 3gpp --excel C:\\work\\in.xlsx --no-proxy-analysis -vv\n"
        ),
        formatter_class=argparse.RawTextHelpFormatter,  # 改行そのまま表示
    )

    # === 必須3引数 ===
    p.add_argument(
        "--drive",
        required=True,
        type=drive_type,
        help=f"{emo.folder} ドライブレター（例: C / D / E）",
    )
    p.add_argument(
        "--db",
        required=True,
        type=db_type,
        choices=sorted(VALID_DB),
        help=f"{emo.db} 解析対象DB（ieee / 3gpp）",
    )
    p.add_argument(
        "--excel",
        required=True,
        type=excel_path_type,
        help=f"{emo.excel} 入力Excelの絶対パス（例: C:\\path\\to\\file.xlsx）",
    )

    # === プロキシ分析（既定ON）：--no-proxy-analysis で無効化（Python 3.9+） ===
    try:
        p.add_argument(
            "--proxy-analysis",
            action=argparse.BooleanOptionalAction,
            default=True,
            help=f"{emo.net} プロキシ分析を行う（既定: 有効）。--no-proxy-analysis で無効化。",
        )
    except Exception:
        # 旧Python互換（on/off）
        p.add_argument(
            "--proxy-analysis",
            choices=["on", "off"],
            default="on",
            help=f"{emo.net} プロキシ分析（on/off、既定 on）",
        )

    # === 既存オプション（雛形維持） ===
    p.add_argument(
        "-n", "--dry-run",
        action="store_true",
        help=f"{emo.warn} 実行せずに手順だけ確認",
    )
    p.add_argument(
        "-v", "--verbose",
        action="count",
        default=0,
        help=f"{emo.star} 冗長度を上げる（-v:INFO, -vv:DEBUG）",
    )

    return p

# ========== アプリ本体（副作用をここに閉じ込めない） ==========
def run(cfg: Config) -> int:
    log = logging.getLogger("mytool")
    log.debug("Config: %s", cfg)

    if cfg.dry_run:
        log.info("ドライラン: 実行はしません")
        return 0

    try:
        print(
            "\n%s\n%s DB=%s\n%s DRIVE=%s\n%s EXCEL=%s\n%s"
            % (emo.sep_full,  # 上部セパレータ
            emo.db, cfg.db,
            emo.folder, cfg.drive,
            emo.excel, cfg.excel_path,
            emo.dash_full)  # 下部セパレータ
        )

        log.info("%s プロキシ分析: %s",
                emo.net, "有効" if cfg.proxy_analysis else "無効")

        # === ここに本処理 ===
        # if cfg.db == "ieee": ...
        # elif cfg.db == "3gpp": ...
        # if cfg.proxy_analysis: do_proxy_check()
        # process_excel(cfg.excel_path)

        MY_DB = cfg.db

        MY_PROXY = get_proxy_from_cmd(prefer="http", allow_env_fallback=True)
        MY_DRIVE = cfg.drive
        MY_DOWNLOAD_PATH = get_downloads_path(MY_DRIVE)
        MY_INPUT_XLSX_PATH = build_case_folder_from_excel(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx")

        if MY_DB == "3gpp":
            print("3gpp")
            download_html_for_3gpp(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx",drive=MY_DRIVE,proxy=MY_PROXY)
            make_doc_list_3gpp(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx",drive=MY_DRIVE)
            res = fetch_3gpp_docs(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx",proxy=MY_PROXY)
            save_results_to_xlsx(res,MY_DOWNLOAD_PATH,"out_"+MY_DB)
            res_zip = extract_zip_to_docs_from_fold(MY_DOWNLOAD_PATH,"out_"+MY_DB)
            write_res_zip_paths_to_xlsx(res_zip,MY_DOWNLOAD_PATH,"out_file_"+MY_DB)
            l = read_column_as_list(MY_DOWNLOAD_PATH,"out_file_"+MY_DB,0)
            convert_office_to_html(l,MY_DOWNLOAD_PATH / "combine.html")


        if MY_DB == "ieee":
            print("ieee")
            download_html_for_ieee(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx",drive=MY_DRIVE,proxy=MY_PROXY)
            make_doc_list_ieee(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx",drive=MY_DRIVE)
            res = fetch_ieee_docs(MY_DOWNLOAD_PATH,"input_"+MY_DB+".xlsx",proxy=MY_PROXY)
            save_results_to_xlsx(res,MY_DOWNLOAD_PATH,"out_"+MY_DB)
            l = read_column_as_list(MY_DOWNLOAD_PATH,"out_"+MY_DB,6)
            convert_office_to_html(l,MY_DOWNLOAD_PATH / "combine.html")
            




        log.info("\n%s\n%s 処理が完了しました %s\n%s",
                emo.sep_full, emo.success, emo.finish, emo.dash_full)
        return 0

    except FileNotFoundError as e:
        log.error("%s ファイルが見つかりません: %s", emo.invalid, e)
        return 2
    except Exception:
        log.exception("%s 予期しないエラー", emo.fail)
        return 1

# ========== エントリポイント ==========
def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    ns = parser.parse_args(argv)

    # 旧Python向けの proxy_analysis 正規化
    proxy_flag = ns.proxy_analysis
    if isinstance(proxy_flag, str):
        proxy_flag = (proxy_flag.lower() == "on")

    cfg = Config(
        drive=ns.drive,
        db=ns.db,
        excel_path=ns.excel,
        proxy_analysis=bool(proxy_flag),
        verbose=ns.verbose,
        dry_run=ns.dry_run,
    )
    setup_logging(cfg.verbose)
    return run(cfg)

if __name__ == "__main__":
    sys.exit(main())