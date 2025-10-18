import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path
from typing import Optional, Any
import pandas as pd

def get_html_url_3gpp(folder_abs_path: str, filename: str) -> Optional[Any]:
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
    df = pd.read_excel(
        p,
        sheet_name=0,
        engine=engine,
        header=None,
        usecols=[1],
        skiprows=5,
        nrows=1
    )

    if df.empty:
        return None
    val = df.iat[0, 0]

    return None if pd.isna(val) else val

if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_3gpp.xlsx"
    print(get_html_url_3gpp(folder_abs_path, filename))