# main.py
from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

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
        raise argparse.ArgumentTypeError("ドライブレターは英字1文字（例: C / D / E）で指定してください。")
    return s.upper()

def db_type(s: str) -> str:
    s = (s or "").strip().lower()
    if s not in VALID_DB:
        raise argparse.ArgumentTypeError("解析対象DBは ieee または 3gpp を指定してください。")
    return s

def excel_path_type(s: str) -> Path:
    p = Path(s)
    if not p.is_absolute():
        raise argparse.ArgumentTypeError("Excelパスは絶対パスで指定してください。")
    if not p.exists():
        raise argparse.ArgumentTypeError("指定したExcelファイルが存在しません。")
    if p.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise argparse.ArgumentTypeError("Excel拡張子ではありません（.xlsx / .xlsm / .xls）。")
    return p

# ========== CLI パーサ ==========
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="mytool",
        description="必須: ドライブレター / 解析対象DB / 入力Excel絶対パス。オプション: プロキシ分析（既定ON）、ドライラン、冗長度。",
    )
    # 必須3引数
    p.add_argument("--drive", required=True, type=drive_type,
                   help="ドライブレター（例: C / D / E）")
    p.add_argument("--db", required=True, type=db_type, choices=sorted(VALID_DB),
                   help="解析対象DB（ieee / 3gpp）")
    p.add_argument("--excel", required=True, type=excel_path_type,
                   help=r"入力Excelの絶対パス（例: C:\path\to\file.xlsx）")

    # プロキシ分析（既定ON）：--no-proxy-analysis で無効化（Python 3.9+）
    try:
        p.add_argument("--proxy-analysis", action=argparse.BooleanOptionalAction,
                       default=True, help="プロキシ分析を行う（既定: 有効）。--no-proxy-analysis で無効化。")
    except Exception:
        # 旧Python互換（on/off）
        p.add_argument("--proxy-analysis", choices=["on", "off"], default="on",
                       help="プロキシ分析（on/off、既定 on）")

    # 既存オプション（雛形維持）
    p.add_argument("-n", "--dry-run", action="store_true", help="実行せずに手順だけ確認")
    p.add_argument("-v", "--verbose", action="count", default=0,
                   help="冗長度を上げる (-v:INFO, -vv:DEBUG)")
    return p

# ========== アプリ本体（副作用をここに閉じ込めない） ==========
def run(cfg: Config) -> int:
    log = logging.getLogger("mytool")
    log.debug("Config: %s", cfg)

    if cfg.dry_run:
        log.info("ドライラン: 実行はしません")
        return 0

    try:
        log.info("DB: %s / ドライブ: %s / Excel: %s", cfg.db, cfg.drive, cfg.excel_path)
        log.info("プロキシ分析: %s", "有効" if cfg.proxy_analysis else "無効")

        # === ここに本処理 ===
        # 例: if cfg.db == "ieee": ... / elif cfg.db == "3gpp": ...
        #     if cfg.proxy_analysis: do_proxy_check()
        #     process_excel(cfg.excel_path)

        print("処理が完了しました。")  # 必要に応じて結果出力
        return 0

    except FileNotFoundError as e:
        log.error("ファイルが見つかりません: %s", e)
        return 2
    except Exception:
        log.exception("予期しないエラー")
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
