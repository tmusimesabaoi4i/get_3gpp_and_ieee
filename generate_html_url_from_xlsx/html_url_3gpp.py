import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pathlib import Path
from typing import Optional, Any, Union
import pandas as pd

def get_html_url_3gpp(folder_abs_path: str, filename: str, sheet: Union[int, str] = 0) -> Optional[Any]:
    excel_path = Path(folder_abs_path) / filename
    if not excel_path.is_absolute():
        raise ValueError(f"{emo.warn} folder_abs_path は絶対パスで指定してください。")
    if not excel_path.exists():
        raise FileNotFoundError(f"{emo.warn} Excel ファイルが見つかりません: {p}")

    suffix = excel_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        engine = "openpyxl"
    elif suffix == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"{emo.warn} 未対応の拡張子です: {suffix}（.xlsx / .xlsm / .xls）")
    df = pd.read_excel(
        excel_path,
        sheet_name=sheet,
        engine=engine,
        header=None,
        usecols=[1],
        skiprows=5,
        nrows=1,
        dtype=str
    )

    if df.empty:
        return None
    val = df.iat[0, 0]

    return None if pd.isna(val) else val

if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_3gpp.xlsx"
    print(get_html_url_3gpp(folder_abs_path, filename))