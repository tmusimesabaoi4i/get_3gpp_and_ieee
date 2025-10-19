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

def load_condition_3gpp(
        excel_path: str,
        sheet_name: str = "Sheet2",
        engine: Optional[str] = None,
    ) -> Tuple[List[Optional[str]], List[List[str]]]:

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
        return [], []

    last_row_idx = has_any[has_any].index.max()
    df = df.iloc[: last_row_idx + 1].copy()

    def _norm(x) -> str:
        return "" if x is None else str(x).strip()

    agenda_items: List[Optional[str]] = []
    keywords_per_row: List[List[str]] = []

    for _, row in df.iterrows():
        if not any(_not_empty(v) for v in row):
            continue

        first = _norm(row.iloc[0]) if len(row) > 0 else ""
        agenda_items.append(first if first != "" else None)

        kws = [_norm(v) for v in row.iloc[1:].tolist()]
        kws = [k for k in kws if k != ""]
        keywords_per_row.append(kws)

    assert len(agenda_items) == len(keywords_per_row)
    return agenda_items, keywords_per_row

def load_base_lists(
    excel_path: str,
    sheet_name: str = "Sheet1",
    engine: Optional[str] = None,
    ) -> Tuple[List[str], List[str]]:

    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
        engine=engine,
        header=None,
        usecols=[2,5],     # 3列目（= C列）
        dtype=str,
        keep_default_na=False,
    )

    df = df.map(lambda x: "" if x is None else str(x).strip())

    mask = (df.iloc[:, 0].ne("")) & (df.iloc[:, 1].ne(""))
    df = df[mask]

    base_title_list = df.iloc[:, 0].tolist()
    base_ai_list    = df.iloc[:, 1].tolist()

    return base_title_list, base_ai_list

def _canon_agenda_strict(
        x: Optional[str]
    ) -> Optional[str]:
    if x is None:
        return None
    s = str(x).strip()
    if not s:
        return None
    s = s.replace("\u00a0", " ").replace("\u3000", " ")
    s = s.replace("\uff0e", ".")
    s = re.sub(r"\s*\.\s*", ".", s)
    s = re.sub(r"\s{2,}", " ", s)
    return s

def find_row_with_condition_3gpp(
        agenda_items: List[Optional[str]],
        keywords_per_row: List[List[str]],
        base_ai_list: Sequence[Optional[str]],
        base_title_list: Sequence[Optional[str]],
        *,
        one_based_output: bool = False,
    ) -> List[List[int]]:
    if len(base_ai_list) != len(base_title_list):
        raise ValueError("base_ai_list と base_title_list の長さが一致しません。")
    if len(agenda_items) != len(keywords_per_row):
        raise ValueError("agenda_items と keywords_per_row の長さが一致しません。")

    base_ai_norm  = [_canon_agenda_strict(v) for v in base_ai_list]
    title_norms   = [_norm_space_casefold(t) for t in base_title_list]

    ai_index: dict[str, List[int]] = {}
    for j, key in enumerate(base_ai_norm):
        if key is None:
            continue
        ai_index.setdefault(key, []).append(j)

    results: List[List[int]] = []

    for ag, kws in zip(agenda_items, keywords_per_row):
        key = _canon_agenda_strict(ag)
        candidates = ai_index.get(key, [])

        patterns = []
        for k in kws:
            nk = _norm_space_casefold(k)
            if not nk:
                continue
            pat = re.compile(rf"\b{re.escape(nk)}\b")
            patterns.append(pat)

        matched: List[int] = []
        for j in candidates:
            tn = title_norms[j]
            ok = all(p.search(tn) for p in patterns) if patterns else True
            if ok:
                matched.append(j + 1 if one_based_output else j)

        results.append(matched)

    return results

def _norm_space_casefold(
        s: Optional[str]
    ) -> str:
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
def condition_row_3gpp(folder_abs_path: str, filename: str) -> List[int]:
    excel_in = Path(folder_abs_path) / filename
    if not excel_in.is_absolute():
        raise ValueError("folder_abs_path は絶対パスで指定してください。")
    if not excel_in.exists():
        raise FileNotFoundError(f"Excel ファイルが見つかりません: {excel_in}")

    suffix = excel_in.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        engine = "openpyxl"
    elif suffix == ".xls":
        engine = "xlrd"
    else:
        raise ValueError(f"未対応の拡張子です: {suffix}（.xlsx / .xlsm / .xls）")

    agenda_items, keywords_per_row = load_condition_3gpp(  # type: ignore[name-defined]
        str(excel_in), engine=engine
    )

    base_path = build_case_folder_from_excel(  # type: ignore[name-defined]
        folder_abs_path=folder_abs_path, filename=filename, sheet=0, drive=None
    )

    base_path = Path(base_path)
    create_subfolder_when_absent(base_path, "EXCEL")  # type: ignore[name-defined]
    excel_dir = base_path / "EXCEL"

    doclist_path = excel_dir / "3gpp_doc_list.xlsx"

    if file_exists_in_folder(excel_dir, "3gpp_doc_list.xlsx"):
        base_title_list, base_ai_list = load_base_lists(
            doclist_path, engine=engine
        )

        result_find = find_row_with_condition_3gpp(  # type: ignore[name-defined]
            agenda_items,
            keywords_per_row,
            base_ai_list,
            base_title_list,
            one_based_output=True,
        )

        condition_row = build_condition_row(result_find)  # type: ignore[name-defined]
        return condition_row

# ---------------- CLI エントリ ----------------
if __name__ == "__main__":
    folder_abs_path = r"C:\Users\yohei\Downloads"
    filename = "input_3gpp.xlsx"
    rows = condition_row_3gpp(folder_abs_path, filename)
    # 結果を表示（例: 2,4,6）
    print(",".join(str(i) for i in rows))