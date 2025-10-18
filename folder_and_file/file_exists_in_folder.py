import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

from pure_download.download_util import (
    to_double_backslash_literal,
)

import re
from pathlib import Path, PureWindowsPath

def _to_path_from_any_windows_str(p: str) -> Path:
    """Windows 風文字列（r'...', 'C:\\x', PureWindowsPath('...') など）を実体 Path に"""
    s = str(p).strip()
    m = re.match(r'^\s*(?:Pure)?Windows?Path\s*\(\s*[rR]?[\'"](.+)[\'"]\s*\)\s*$', s)
    if m:
        s = m.group(1)
    s = s.replace("\\\\", "\\")  # 表示用の \\ を実体の \ に戻す
    return Path(PureWindowsPath(s))

def file_exists_in_folder(folder_abs_path: str, filename: str) -> bool:
    folder = _to_path_from_any_windows_str(folder_abs_path)
    if not folder.is_absolute() or not folder.is_dir():
        return False
    name_only = Path(_to_path_from_any_windows_str(filename).name)
    target = folder / name_only
    exists = target.is_file()
    if exists:
        print(f"{emo.file} {to_double_backslash_literal(str(folder))}に{to_double_backslash_literal(name_only.name)}が存在します")
    return exists

# ============== 実行部 ==============
if __name__ == "__main__":
    folder = r"C:\Users\yohei\Downloads"
    file1  = "mentor_dcn_list.html"
    file2  = r"sub\nope.txt"

    file_exists_in_folder(folder, file1)  # 存在すればメッセージを表示し、True を返す
    file_exists_in_folder(folder, file2)  # 何も表示せず、False を返す