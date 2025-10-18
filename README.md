# get_3gpp_and_ieee

# How To CLONE

```
git clone https://github.com/tmusimesabaoi4i/get_3gpp_and_ieee.git
```

# How To PUSH

```
git add ./
```

and

```
git commit --allow-empty -m "　"
```

and

```
git push
```

# 入力エクセル

## 3gpp
<!-- 3GPP｜① 日本語ラベル -->
<table style="width:100%;table-layout:fixed;border-collapse:collapse;">
  <colgroup>
    <col style="width:8%">
    <col style="width:22%">
    <col style="width:70%">
  </colgroup>
  <thead>
    <tr>
      <th style="border:1px solid #ccc;padding:6px 8px;text-align:right;">行</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">A列（ラベル）</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">B列</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">1</td><td style="border:1px solid #ccc;padding:6px 8px;">日付</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">20251018</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">2</td><td style="border:1px solid #ccc;padding:6px 8px;">案件番号</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">特願20XX-XXXXXX</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">3</td><td style="border:1px solid #ccc;padding:6px 8px;">TSG</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">tsg_ran</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">4</td><td style="border:1px solid #ccc;padding:6px 8px;">作業部会(WG)</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">WG2_RL2</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">5</td><td style="border:1px solid #ccc;padding:6px 8px;">会期/会合(シリーズ)</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">TSGR2_105bis</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">6</td><td style="border:1px solid #ccc;padding:6px 8px;">文書ディレクトリURL</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_105bis/Docs</td></tr>
  </tbody>
</table>

<!-- 3GPP｜② 変数名（snake_case） -->
<table style="width:100%;table-layout:fixed;border-collapse:collapse;margin-top:10px;">
  <colgroup>
    <col style="width:8%">
    <col style="width:22%">
    <col style="width:70%">
  </colgroup>
  <thead>
    <tr>
      <th style="border:1px solid #ccc;padding:6px 8px;text-align:right;">行</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">A列（変数名）</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">B列</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">1</td><td style="border:1px solid #ccc;padding:6px 8px;">date</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">20251018</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">2</td><td style="border:1px solid #ccc;padding:6px 8px;">case_id</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">特願20XX-XXXXXX</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">3</td><td style="border:1px solid #ccc;padding:6px 8px;">tsg</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">tsg_ran</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">4</td><td style="border:1px solid #ccc;padding:6px 8px;">wg</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">WG2_RL2</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">5</td><td style="border:1px solid #ccc;padding:6px 8px;">series</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">TSGR2_105bis</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">6</td><td style="border:1px solid #ccc;padding:6px 8px;">docs_url</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_105bis/Docs</td></tr>
  </tbody>
</table>

<!-- IEEE｜① 日本語ラベル -->
<table style="width:100%;table-layout:fixed;border-collapse:collapse;margin-top:18px;">
  <colgroup>
    <col style="width:8%">
    <col style="width:22%">
    <col style="width:70%">
  </colgroup>
  <thead>
    <tr>
      <th style="border:1px solid #ccc;padding:6px 8px;text-align:right;">行</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">A列（ラベル）</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">B列</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">1</td><td style="border:1px solid #ccc;padding:6px 8px;">日付</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">20251018</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">2</td><td style="border:1px solid #ccc;padding:6px 8px;">案件番号</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">特願20XX-XXXXXX</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">3</td><td style="border:1px solid #ccc;padding:6px 8px;">タスクグループ</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">be</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">4</td><td style="border:1px solid #ccc;padding:6px 8px;">年</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">2024</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">5</td><td style="border:1px solid #ccc;padding:6px 8px;">文書一覧URL</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">https://mentor.ieee.org/802.11/documents?n=1&amp;o=7d&amp;is_group=00be&amp;is_year=2024</td></tr>
  </tbody>
</table>

<!-- IEEE｜② 変数名（snake_case） -->
<table style="width:100%;table-layout:fixed;border-collapse:collapse;margin-top:10px;">
  <colgroup>
    <col style="width:8%">
    <col style="width:22%">
    <col style="width:70%">
  </colgroup>
  <thead>
    <tr>
      <th style="border:1px solid #ccc;padding:6px 8px;text-align:right;">行</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">A列（変数名）</th>
      <th style="border:1px solid #ccc;padding:6px 8px;">B列</th>
    </tr>
  </thead>
  <tbody>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">1</td><td style="border:1px solid #ccc;padding:6px 8px;">date</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">20251018</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">2</td><td style="border:1px solid #ccc;padding:6px 8px;">case_id</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">特願20XX-XXXXXX</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">3</td><td style="border:1px solid #ccc;padding:6px 8px;">task_group</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">be</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">4</td><td style="border:1px solid #ccc;padding:6px 8px;">year</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">2024</td></tr>
    <tr><td style="border:1px solid #ccc;padding:6px 8px;text-align:right;">5</td><td style="border:1px solid #ccc;padding:6px 8px;">docs_url</td><td style="border:1px solid #ccc;padding:6px 8px;word-break:break-all;overflow-wrap:anywhere;">https://mentor.ieee.org/802.11/documents?n=1&amp;o=7d&amp;is_group=00be&amp;is_year=2024</td></tr>
  </tbody>
</table>



# How To USE

このツールは **Python スクリプトとして実行**しても、**単一 .exe** にビルドして実行しても使えます。  
ここでは両方のやり方と、.exe の作り方（PyInstaller）をまとめます。

---

## 1) Python で実行

### ヘルプ
```bash
python main.py --help
```

### 例：プロキシ分析は既定で有効
```bash
python main.py --drive C --db ieee --excel C:\Users\yohei\Downloads\ieee_input.xlsx
```

### 例：ログ冗長度（INFO）
```bash
python main.py --drive C --db ieee --excel C:\Users\yohei\Downloads\ieee_input.xlsx -v
```

### 例：ログ冗長度（DEBUG）
```bash
python main.py --drive C --db ieee --excel C:\Users\yohei\Downloads\ieee_input.xlsx -vv
```

### 例：プロキシ分析を無効化（Python 3.9+）
```bash
python main.py --drive C --db 3gpp --excel C:\work\in.xlsx --no-proxy-analysis
```

### 例：旧Python（3.8 など）
```bash
python main.py --drive C --db 3gpp --excel C:\work\in.xlsx --proxy-analysis off
```

> **必須引数**：`--drive`（ドライブレター）, `--db`（`ieee`/`3gpp`）, `--excel`（Excel の**絶対パス**）  
> **オプション**：`--no-proxy-analysis` または `--proxy-analysis off`（既定は有効）  
> **ログ**：`-v`（INFO）, `-vv`（DEBUG） / `--help` で全一覧表示

---

## 2) .exe で実行（配布・配信用）

ビルド済みの `mytool.exe` がある前提での実行例です。  
コマンドライン・オプションは **Python 実行時と同じ**です。

### ヘルプ
```bash
mytool.exe --help
```

### 例：プロキシ分析は既定で有効
```bash
mytool.exe --drive C --db ieee --excel C:\Users\yohei\Downloads\ieee_input.xlsx
```

### 例：ログ冗長度
```bash
mytool.exe --drive C --db ieee --excel C:\Users\yohei\Downloads\ieee_input.xlsx -vv
```

### 例：プロキシ分析を無効化
```bash
mytool.exe --drive C --db 3gpp --excel C:\work\in.xlsx --no-proxy-analysis
```

#### ダブルクリック運用したい場合（バッチ例）
```bat
@echo off
"%~dp0mytool.exe" --drive C --db ieee --excel C:\Users\yohei\Downloads\ieee_input.xlsx -vv --log-file "%~dp0mytool.log"
pause
```

> コンソール非表示ビルド（`-w`）だと標準出力が見えません。ログは `--log-file` を併用してください。

---

## 3) .exe の作り方（PyInstaller）

### セットアップ
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -U pip
pip install pyinstaller
# （必要に応じて）requests, pywin32 など依存を追加
pip install requests pywin32
```

### 単一ファイル .exe をビルド
```bash
pyinstaller -F --paths . -n mytool main.py
```
- 出力: `dist\mytool.exe`
- `--paths .` はプロジェクト直下のパッケージ（例：`emoji/`, `pure_download/`）を見せるための検索パス追加
- `pywin32`（`win32com` を使う場合）は自動同梱されます。環境によっては保険として以下を足すと安定します：
  ```bash
  pyinstaller -F --paths . -n mytool --hidden-import win32timezone main.py
  ```

### コンソールを出さない（GUI/バックグラウンド）
```bash
pyinstaller -F -w --paths . -n mytool main.py
```
- コンソールが出ない代わりに、**ログ出力は `--log-file` を必ず指定**してください。

---

## トラブルシューティング

- **「モジュールが見つからない」**  
  → `--paths .` を付けて再ビルド。インポートは**絶対インポート**（例：`from emoji.emoscript import emo`）に。

- **pywin32 周りのエラー**  
  → `pip install pywin32` 済みか確認。`--hidden-import win32timezone` を付けて再ビルド。

- **HTTPS 証明書エラー**  
  → `pip install certifi` を確認。プロキシ環境なら OS/社内ルート証明書設定も要確認。

---

## 仕様の要点（再掲）

- **必須**：`--drive` / `--db`（`ieee` or `3gpp`）/ `--excel`（絶対パスの Excel）  
- **プロキシ分析**：既定 **有効**。無効化は `--no-proxy-analysis`（旧Pythonは `--proxy-analysis off`）  
- **ログ**：`-v`（INFO）, `-vv`（DEBUG）、必要なら `--log-file` でファイル保存
