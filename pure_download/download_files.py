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

def current_partial_size(
        path: str
    ) -> int:
    try:
        return os.path.getsize(path)
    except Exception:
        return 0

def truncate_file(
        path: str,
        size: int
    ) -> None:
    with open(path, "r+b") as f:
        f.truncate(size)

def download_file_safely_msxml2(
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
        use_curl_fallback: bool = True,
    ) -> str:

    if not msxml2_available():
        raise RuntimeError("MSXML2 ヘルパが未定義です（msxml2_request/msxml2_all_headers_dict/msxml2_read_body_bytes）。")
    if not download_url:
        raise ValueError("download_url が指定されていません。")

    parsed = urlparse(download_url)
    pure_filename = os.path.basename(parsed.path) or "download.bin"
    file_extension = os.path.splitext(pure_filename)[1]

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

    # 1) HEAD プローブ
    total_size, accept_ranges, if_range_token = probe_remote_msxml2(download_url, common_headers, tms, pxy)

    # 2) .part 整合チェック
    part_size0 = current_partial_size(temp_path)
    if total_size is not None and total_size >= 0:
        if part_size0 == total_size:
            os.replace(temp_path, final_path)
            print(f"✅ 既に全量取得済み → {final_path}")
            return file_extension
        elif part_size0 > total_size:
            print(f"⚠️ 部分ファイル超過: {part_size0} > {total_size} → 切り詰め")
            try:
                truncate_file(temp_path, total_size)
                part_size0 = total_size
            except Exception as te:
                print(f"⚠️ 切り詰め失敗: {te} → 全量取り直し")
                try: os.remove(temp_path)
                except Exception: pass
                part_size0 = 0
    else:
        if not accept_ranges:
            part_size0 = 0

    # 3) 本体ダウンロード（MSXML2）
    for attempt in range(1, max_retries + 1):
        part_size = current_partial_size(temp_path)
        try:
            headers = dict(common_headers)
            if part_size > 0 and accept_ranges:
                headers["Range"] = f"bytes={part_size}-"
                if if_range_token:
                    headers["If-Range"] = if_range_token

            print(f"[{attempt}/{max_retries} PROXY={pxy or 'NONE'}] GET {download_url} (resume {part_size}, MSXML2)")

            http = msxml2_request("GET", download_url, headers, tms, pxy)
            status = int(http.status)

            if status == 416:
                print("⚠️ 416 受信 → 再プローブして整合性回復を試行")
                total_size, accept_ranges, if_range_token = probe_remote_msxml2(download_url, common_headers, tms, pxy)
                ps = current_partial_size(temp_path)
                if total_size is not None:
                    if ps == total_size:
                        os.replace(temp_path, final_path)
                        print(f"✅ 416 だったが既に全量取得済み → {final_path}")
                        return file_extension
                    if ps > total_size:
                        print("⚠️ 416: 部分ファイル超過 → 切り詰めて再試行")
                        truncate_file(temp_path, total_size)
                raise RuntimeError("Retry after 416")

            if part_size > 0 and "Range" in headers and status == 200:
                part_size = 0

            if status in (418, 429):
                raise RuntimeError(f"{status} (temporary block)")
            if status < 200 or status >= 300:
                raise RuntimeError(f"HTTP {status}")

            data = msxml2_read_body_bytes(http)
            mode = "ab" if part_size > 0 and status == 206 else "wb"
            with open(temp_path, mode) as f:
                f.write(data)

            os.replace(temp_path, final_path)
            print(f"✅ 成功（MSXML2）→ {final_path}")
            return file_extension

        except Exception as e:
            print(f"⚠️ 失敗 ({attempt}/{max_retries}) MSXML2: {e}")
            if attempt < max_retries:
                sleep(min(2 * attempt, 10))
                continue

    raise RuntimeError("ダウンロードに失敗しました。")

# ============== 実行部 ==============
if __name__ == "__main__":
    download_url = "https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_105bis/Docs/R2-1903010.zip"
    # download_url = "https://mentor.ieee.org/802.11/dcn/25/11-25-1818-00-0PAR-par-review-sc-mtg-agenda-and-comment-slides-2025-november-bangkok.pptx"
    download_path = to_double_backslash_literal(r'C:\Users\yohei\Downloads')

    # --- (F) HTTPセッション準備（3GPP FTP / Cookie取得） ---
    LANDING, sess = get_landing_and_session("IEEE")

    try:
        ext = download_file_safely_msxml2(
            download_url,
            download_path,              # フォルダ or フルパスどちらでもOK
            "3gpp",
            session=sess,               # ★ Cookie を自動で付与
            referer=LANDING,            # ★ Referer も付与
            # proxy="http://proxy.example.com:8080",  # 必要なら
            connect_timeout=10,
            read_timeout=180,
            max_retries=5,
            use_curl_fallback=True,
            )
            
        print(f"📝 拡張子: {ext or '(不明)'}")
    except Exception as e:
        print(f"✗ エラー: {e}")
        raise