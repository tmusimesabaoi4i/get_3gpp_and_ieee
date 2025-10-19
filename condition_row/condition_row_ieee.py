import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))

from pathlib import Path
import re
from typing import List, Optional, Sequence, Tuple
import pandas as pd

from folder_and_file.create_subfolder_when_absent import (
    create_subfolder_when_absent,
    )

from folder_and_file.file_exists_in_folder import (
    file_exists_in_folder,
    )

from util import (
    build_case_folder_from_excel,
    )

def load_condition(
    excel_path: str,
    sheet_name: str = "Sheet2",
    engine: Optional[str] = None,
    ) -> List[List[str]]:

    read_kwargs = dict(
        sheet_name=sheet_name,
        header=None,
        dtype=object,
        keep_default_na=False,
    )
    if engine is not None:
        read_kwargs["engine"] = engine

    df = pd.read_excel(excel_path, **read_kwargs)

    def _not_empty(x) -> bool:
        if x is None:
            return False
        return str(x).strip() != ""

    has_any = df.map(_not_empty).any(axis=1)
    if not has_any.any():
        return []

    last_row_idx = has_any[has_any].index.max()
    df = df.iloc[: last_row_idx + 1].copy()


    df = df.map(lambda x: "" if x is None else str(x).strip())

    keywords_per_row: List[List[str]] = []
    for _, row in df.iterrows():
        if not any(cell != "" for cell in row):
            continue
        kws = [cell for cell in row.tolist() if cell != ""]
        keywords_per_row.append(kws)

    return keywords_per_row

def load_base_lists_ieee(
    excel_path: str,
    sheet_name: str = "Sheet1",
    engine: Optional[str] = None,
    ) -> Tuple[List[str], List[str]]:

    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        engine=engine,
        header=None,
        usecols=[2],     # 3列目（= C列）
        dtype=str,
        keep_default_na=False,
    )


    df = df.map(lambda x: "" if x is None else str(x).strip())

    mask = (df.iloc[:, 0].ne(""))
    df = df[mask]

    base_title_list = df.iloc[:, 0].tolist()
    return base_title_list

def find_row_with_condition_ieee(
    keywords_per_row: List[List[str]],
    base_title_list: Sequence[Optional[str]],
    *,
    one_based_output: bool = False,
    ) -> List[List[int]]:
    """
    各 keywords_per_row[i] の全キーワード（フレーズ）を“単語完全一致”で含む
    タイトル行のインデックスだけを返す（アジェンダ縛りなし、全体から検索）。

    - 照合は 大小無視・空白圧縮・語境界(\b)で完全一致
      例: "condition HO" → タイトル中の同一フレーズにマッチ
    - キーワードが空の行は全タイトル一致
    - 返り値インデックスは 0始まり。one_based_output=True で 1始まりに変更可
    """
    # タイトルの正規化（先に一括で）
    title_norms = [_norm_space_casefold(t) for t in base_title_list]

    results: List[List[int]] = []

    for kw_list in keywords_per_row:
        # キーワードを正規化し、語境界で完全一致する正規表現に
        patterns = []
        for k in kw_list:
            nk = _norm_space_casefold(k)
            if not nk:
                continue
            # フレーズ全体を語境界で囲む。記号類はエスケープ
            patterns.append(re.compile(rf"\b{re.escape(nk)}\b"))

        matched: List[int] = []
        for j, tn in enumerate(title_norms):
            ok = all(p.search(tn) for p in patterns) if patterns else True
            if ok:
                matched.append(j + 1 if one_based_output else j)

        results.append(matched)

    return results

def _norm_space_casefold(s: Optional[str]) -> str:
    """両端トリム → 複数空白/改行を1つに圧縮 → casefold（大小無視）。"""
    if s is None:
        return ""
    return " ".join(str(s).split()).casefold()

def build_condition_row(
        result: Sequence[Sequence[int]],
    ) -> List[int]:
    condition_row = sorted({
        int(idx) for group in result for idx in group
    })
    return condition_row

# ##############################################
def condition_row_ieee(folder_abs_path: str, filename: str) -> List[int]:
    excel_in = Path(folder_abs_path) / filename
    if not excel_in.is_absolute():
        raise ValueError("folder_abs_path は絶対パスで指定してください。")
    if not excel_in.exists():
        raise FileNotFoundError(f"Excel ファイルが見つかりません: {excel_in}")

    suffix = excel_in.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        engine: Optional[str] = "openpyxl"
    elif suffix == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"未対応の拡張子です: {suffix}（.xlsx / .xlsm / .xls）")

    keywords_per_row = load_condition(   # type: ignore[name-defined]
        str(excel_in), engine=engine
    )

    base_path = build_case_folder_from_excel(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path, filename=filename, sheet=0, drive=None
    )

    base_path = Path(base_path)
    create_subfolder_when_absent(base_path, "EXCEL")  # type: ignore[name-defined]
    excel_dir = base_path / "EXCEL"

    doclist_path = excel_dir / "ieee_doc_list.xlsx"

    if file_exists_in_folder(excel_dir, "ieee_doc_list.xlsx"):
        base_title_list = load_base_lists_ieee(
            doclist_path, engine=engine
        )

        result_find = find_row_with_condition_ieee(  # type: ignore[name-defined]
            keywords_per_row,
            base_title_list,
            one_based_output=True,
        )

        condition_row = build_condition_row(result_find)  # type: ignore[name-defined]
        return condition_row

# ---------------- CLI エントリ ----------------
if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_ieee.xlsx"
    rows = condition_row_ieee(folder_abs_path, filename)
    # 結果を表示（例: 2,4,6）
    print(",".join(str(i) for i in rows))