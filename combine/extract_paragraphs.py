# -*- coding: utf-8 -*-
"""
Word/PPTX 段落を抽出して <li> 化、Heading でファイル名/スライド番号を示す。
要件:
- 開始前に WINWORD.EXE / POWERPNT.EXE を kill（cmdの taskkill を呼ぶ）
- 対象外拡張子はスキップ
- ファイルを開けない場合のみ print してスキップ（他は出力しない）
- CSS なし
- Windows + Microsoft Office + pywin32 前提
"""

from __future__ import annotations
import html
import subprocess
from pathlib import Path
from typing import Iterable, List, Dict

import pythoncom
import win32com.client as win32

WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf"}
PPT_EXTS  = {".ppt", ".pptx", ".pptm"}


# -------------------- 起動前に Office プロセスを kill --------------------
def kill_office_processes() -> None:
    """
    WINWORD.EXE と POWERPNT.EXE を強制終了。
    失敗しても例外は投げず、静かに無視（printもしない）。
    ※ 強制終了により未保存データは失われます。
    """
    # /F: 強制終了, /T: 子プロセス含む, /IM: 画像名指定
    for image_name in ("WINWORD.EXE", "POWERPNT.EXE"):
        try:
            subprocess.run(
                ["taskkill", "/IM", image_name, "/F", "/T"],
                capture_output=True,
                text=True,
                check=False
            )
        except Exception:
            pass


# -------------------- 共通ユーティリティ --------------------
def _clean_paragraph_text(s: str) -> str:
    if s is None:
        return ""
    s = s.replace("\r", "").replace("\x07", "")
    return s.strip()

def _ensure_paragraphs_list(items: Iterable[str]) -> List[str]:
    out: List[str] = []
    for t in items:
        t = _clean_paragraph_text(t)
        if t:
            out.append(t)
    return out


# -------------------- Word: 段落抽出 --------------------
def extract_paragraphs_from_word(path: Path, word_app=None) -> List[str]:
    created_here = False
    if word_app is None:
        word_app = win32.DispatchEx("Word.Application")
        word_app.Visible = False
        created_here = True

    doc = None
    paras: List[str] = []
    try:
        doc = word_app.Documents.Open(
            FileName=str(path),
            ReadOnly=True,
            AddToRecentFiles=False
        )
        for para in doc.Paragraphs:  # 1-based
            text = _clean_paragraph_text(str(para.Range.Text))
            if text:
                paras.append(text)
    finally:
        if doc is not None:
            doc.Close(False)
        if created_here and word_app is not None:
            word_app.Quit()
    return paras


# -------------------- PowerPoint: 段落抽出(スライド別) --------------------
def _iter_shape_paragraphs(shape) -> Iterable[str]:
    # グループ図形（中を再帰）
    try:
        gi = shape.GroupItems
        for i in range(1, gi.Count + 1):
            yield from _iter_shape_paragraphs(gi.Item(i))
        return
    except Exception:
        pass

    # 表（セル内テキスト）
    try:
        if getattr(shape, "HasTable", 0) == -1:
            table = shape.Table
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    cell_shape = table.Cell(r, c).Shape
                    if getattr(cell_shape, "TextFrame", None) is not None:
                        tf = cell_shape.TextFrame
                        if getattr(tf, "HasText", 0) == -1:
                            tr = tf.TextRange
                            count = tr.Paragraphs().Count
                            for i in range(1, count + 1):
                                yield _clean_paragraph_text(tr.Paragraphs(i).Text)
            return
    except Exception:
        pass

    # 通常の TextFrame
    try:
        if getattr(shape, "HasTextFrame", 0) == -1:
            tf = shape.TextFrame
            if getattr(tf, "HasText", 0) == -1:
                tr = tf.TextRange
                count = tr.Paragraphs().Count
                for i in range(1, count + 1):
                    yield _clean_paragraph_text(tr.Paragraphs(i).Text)
    except Exception:
        pass


def extract_paragraphs_from_ppt_grouped(path: Path, ppt_app=None) -> Dict[int, List[str]]:
    created_here = False
    if ppt_app is None:
        ppt_app = win32.DispatchEx("PowerPoint.Application")
        created_here = True

    pres = None
    grouped: Dict[int, List[str]] = {}
    try:
        pres = ppt_app.Presentations.Open(
            FileName=str(path),
            WithWindow=False,
            ReadOnly=True
        )
        for slide in pres.Slides:
            slide_no = slide.SlideIndex  # 1-based
            buf: List[str] = []
            for shape in slide.Shapes:
                for p in _iter_shape_paragraphs(shape):
                    if p:
                        buf.append(p)
            grouped[slide_no] = _ensure_paragraphs_list(buf)
    finally:
        if pres is not None:
            pres.Close()
        if created_here and ppt_app is not None:
            ppt_app.Quit()
    return grouped


# -------------------- HTML 生成 --------------------
def paragraphs_to_html(grouped: Dict[str, Dict]) -> str:
    """
    grouped:
      {
        "file_display": {
            "type": "word", "paras": [..]
            # or
            "type": "ppt",  "slides": { 1:[..], 2:[..], ... }
        },
        ...
      }
    """
    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append('<html lang="ja">')
    parts.append('<meta charset="UTF-8">')
    parts.append("<title>抽出テキスト</title>")
    parts.append("<body>")

    for file_disp, data in grouped.items():
        parts.append(f"<h2>{html.escape(file_disp)}</h2>")

        if data.get("type") == "ppt":
            slides: Dict[int, List[str]] = data.get("slides", {})
            for slide_no in sorted(slides.keys()):
                parts.append(f"<h3>Slide {slide_no}</h3>")
                paras = slides[slide_no]
                if paras:
                    parts.append("<ul>")
                    for p in paras:
                        safe = html.escape(p).replace("\n", "<br>")
                        parts.append(f"<li>{safe}</li>")
                    parts.append("</ul>")
        else:
            paras = data.get("paras", [])
            if paras:
                parts.append("<ul>")
                for p in paras:
                    safe = html.escape(p).replace("\n", "<br>")
                    parts.append(f"<li>{safe}</li>")
                parts.append("</ul>")

    parts.append("</body></html>")
    return "\n".join(parts)


# -------------------- メイン --------------------
def convert_office_to_html(
    dir_list: Iterable[str],
    file_list: Iterable[str],
    output_html_path: str | Path,
) -> Path:
    """
    dir_list と file_list を zip し、各ファイルを処理。
    - 開始前に Word/PowerPoint プロセスを kill
    - 対象外拡張子と存在しないファイルは静かにスキップ
    - 開けないファイルは print してスキップ
    - Word: <h2>ファイル名</h2> の直下に <li>
    - PowerPoint: <h2>ファイル名</h2> の下に <h3>Slide n</h3> + <li>
    """
    # ① Office プロセスを終了
    kill_office_processes()

    # ② COM 初期化
    pythoncom.CoInitialize()

    word_app = None
    ppt_app  = None
    grouped: Dict[str, Dict] = {}

    try:
        for d, f in zip(dir_list, file_list):
            src = Path(d) / f

            if not src.exists():
                # 存在しないファイルは静かにスキップ
                continue

            ext = src.suffix.lower()
            # 対象外拡張子はスキップ
            if ext not in WORD_EXTS and ext not in PPT_EXTS:
                continue

            file_disp = src.name  # 見出し表示

            try:
                if ext in WORD_EXTS:
                    if word_app is None:
                        word_app = win32.DispatchEx("Word.Application")
                        word_app.Visible = False
                    paras = extract_paragraphs_from_word(src, word_app)
                    grouped[file_disp] = {"type": "word", "paras": _ensure_paragraphs_list(paras)}

                elif ext in PPT_EXTS:
                    if ppt_app is None:
                        ppt_app = win32.DispatchEx("PowerPoint.Application")
                    slides = extract_paragraphs_from_ppt_grouped(src, ppt_app)
                    grouped[file_disp] = {"type": "ppt", "slides": slides}

            except Exception as e:
                # ★ ファイルが開けない等 → print してスキップ
                print(f"[SKIP] ファイルを開けませんでした: {src} / エラー: {e}")
                continue

        html_text = paragraphs_to_html(grouped)
        out_path = Path(output_html_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(html_text, encoding="utf-8")
        return out_path

    finally:
        try:
            if word_app is not None:
                word_app.Quit()
        except Exception:
            pass
        try:
            if ppt_app is not None:
                ppt_app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

# 直接実行テスト
if __name__ == "__main__":
    dirs = [r"C:\Users\yohei\Downloads\R2-2206906", r"C:\Users\yohei\Downloads\R2-2206905", r"C:\Users\yohei\Downloads",r"C:\Users\yohei\Downloads"]
    files = ["R2-2206906_C1-224008.docx", "R2-2206905_C1-223972.docx","11-25-1850-00-00bn-p-edca-on-npca-primary-channel.pptx","input_3gpp.xlsx"]
    output = r"C:\Users\yohei\Downloads\a.html"
    out = convert_office_to_html(dirs, files, output)
    print(f"✅ 出力: {out}")
