import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
from emoji.emoscript import emo

import os
from time import sleep
from urllib.parse import urlparse
from typing import Optional

from msxml2_util import (
    msxml2_all_headers_dict,
    msxml2_available,
    msxml2_request,
    msxml2_read_body_bytes,
    probe_remote_msxml2,
)

from download_util import (
    cookie_header_from_session,
    get_landing_and_session,
    is_dir_like,
    normalize_proxy_for_msxml2,
    sanitize_filename,
    to_double_backslash_literal,
)

def download_html_safely_msxml2(
        download_url: str,
        download_path: str,                 # ← ディレクトリ想定
        filename: str,                      # ← MUST（必須）。拡張子は .html に強制
        *,
        session=None,
        proxy: Optional[str] = None,
        connect_timeout: int = 10,
        read_timeout: int = 180,
        max_retries: int = 10,
        referer: Optional[str] = None,
        user_agent: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    ) -> str:

    if not msxml2_available():
        raise RuntimeError("MSXML2 ヘルパが未定義です（msxml2_request/msxml2_all_headers_dict/msxml2_read_body_bytes）。")
    if not download_url:
        raise ValueError("download_url が指定されていません。")
    if not filename or not isinstance(filename, str):
        raise ValueError("filename は必須です。")

    file_extension = ".htlm"

    # フォルダ指定への対応
    base = sanitize_filename(os.path.basename(filename))

    if not base.lower().endswith((file_extension)):
        base += file_extension
        
    # download_path\filename.html を作成
    final_path = os.path.join(download_path, base)
    os.makedirs(os.path.dirname(final_path) or ".", exist_ok=True)
    temp_path  = final_path + ".part"

    # 共通ヘッダ（圧縮無効化でバイト範囲のズレ防止）
    common_headers = {
        "User-Agent": user_agent,
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9,ja;q=0.8",
        "Connection": "close",
        "Accept-Encoding": "identity",
    }
    if referer:
        common_headers["Referer"] = referer

    # Cookie を session から付与
    cookie_hdr = cookie_header_from_session(session, download_url)
    if cookie_hdr:
        common_headers["Cookie"] = cookie_hdr

    # タイムアウト（ms）
    tms = (connect_timeout * 1000, connect_timeout * 1000, read_timeout * 1000, read_timeout * 1000)
    pxy = normalize_proxy_for_msxml2(proxy)

    # ---- 本体（MSXML2, 非2xxで .error.html 退避）----
    for attempt in range(1, max_retries + 1):
        try:
            print(f"{emo.info} [{attempt}/{max_retries} PROXY={pxy or 'NONE'}] GET {download_url} (HTML, MSXML2)")
            http = msxml2_request("GET", download_url, dict(common_headers), tms, pxy)
            status = int(http.status)

            if status in (418, 429):
                raise RuntimeError(f"{status} (temporary block)")
            if status < 200 or status >= 300:
                try:
                    with open(final_path + ".error.html", "w", encoding="utf-8", newline="") as ef:
                        ef.write(getattr(http, "responseText", "") or "")
                except Exception:
                    pass
                raise RuntimeError(f"HTTP {status}")

            html_text = http.responseText
            with open(temp_path, "w", encoding="utf-8", newline="") as f:
                f.write(html_text or "")
            os.replace(temp_path, final_path)
            print(f"{emo.ok} HTML 保存 → {final_path}")
            return file_extension

        except Exception as e:
            print(f"{emo.warn} 失敗 ({attempt}/{max_retries}) MSXML2: {e}")
            if attempt < max_retries:
                sleep(min(2 * attempt, 10))
                continue

    raise RuntimeError("HTMLのダウンロードに失敗しました。")

# ============== 実行部 ==============
if __name__ == "__main__":
    # download_url = "https://www.3gpp.org/ftp"
    download_url = "https://mentor.ieee.org/802.11"
    download_path = to_double_backslash_literal(r'C:\Users\yohei\Downloads')

    # --- (F) HTTPセッション準備（3GPP FTP / Cookie取得） ---
    LANDING, sess = get_landing_and_session("IEEE")

    try:
        ext = download_html_safely_msxml2(
            download_url,
            download_path,              # フォルダ or フルパスどちらでもOK
            "ieee",
            session=sess,               # ★ Cookie を自動で付与
            referer=LANDING,            # ★ Referer も付与
            # proxy="http://proxy.example.com:8080",  # 必要なら
            connect_timeout=10,
            read_timeout=180,
            max_retries=5,
            )
            
        print(f"{emo.info} 拡張子: {ext or '(不明)'}")
    except Exception as e:
        print(f"{emo.warn} エラー: {e}")
        raise
