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
<table>
  <thead>
    <tr>
      <th align="right">行</th>
      <th>A列（ラベル/変数名）</th>
      <th>B列</th>
    </tr>
  </thead>
  <tbody>
    <!-- 3GPP｜① 日本語ラベル -->
    <tr><th colspan="3" align="left">3GPP｜① 日本語ラベル</th></tr>
    <tr><td align="right">1</td><td>日付</td><td>20251018</td></tr>
    <tr><td align="right">2</td><td>案件番号</td><td>特願20XX-XXXXXX</td></tr>
    <tr><td align="right">3</td><td>TSG</td><td>tsg_ran</td></tr>
    <tr><td align="right">4</td><td>作業部会(WG)</td><td>WG2_RL2</td></tr>
    <tr><td align="right">5</td><td>会期/会合(シリーズ)</td><td>TSGR2_105bis</td></tr>
    <tr><td align="right">6</td><td>文書ディレクトリURL</td><td><a href="https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_105bis/Docs">Docs</a></td></tr>

    <!-- 3GPP｜② 変数名（snake_case） -->
    <tr><th colspan="3" align="left">3GPP｜② 変数名（snake_case）</th></tr>
    <tr><td align="right">1</td><td>date</td><td>20251018</td></tr>
    <tr><td align="right">2</td><td>case_id</td><td>特願20XX-XXXXXX</td></tr>
    <tr><td align="right">3</td><td>tsg</td><td>tsg_ran</td></tr>
    <tr><td align="right">4</td><td>wg</td><td>WG2_RL2</td></tr>
    <tr><td align="right">5</td><td>series</td><td>TSGR2_105bis</td></tr>
    <tr><td align="right">6</td><td>docs_url</td><td><a href="https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_105bis/Docs">Docs</a></td></tr>

    <!-- IEEE｜① 日本語ラベル -->
    <tr><th colspan="3" align="left">IEEE｜① 日本語ラベル</th></tr>
    <tr><td align="right">1</td><td>日付</td><td>20251018</td></tr>
    <tr><td align="right">2</td><td>案件番号</td><td>特願20XX-XXXXXX</td></tr>
    <tr><td align="right">3</td><td>タスクグループ</td><td>be</td></tr>
    <tr><td align="right">4</td><td>年</td><td>2024</td></tr>
    <tr><td align="right">5</td><td>文書一覧URL</td>
        <td><a href="https://mentor.ieee.org/802.11/documents?n=1&amp;o=7d&amp;is_group=00be&amp;is_year=2024">Documents</a></td></tr>

    <!-- IEEE｜② 変数名（snake_case） -->
    <tr><th colspan="3" align="left">IEEE｜② 変数名（snake_case）</th></tr>
    <tr><td align="right">1</td><td>date</td><td>20251018</td></tr>
    <tr><td align="right">2</td><td>case_id</td><td>特願20XX-XXXXXX</td></tr>
    <tr><td align="right">3</td><td>task_group</td><td>be</td></tr>
    <tr><td align="right">4</td><td>year</td><td>2024</td></tr>
    <tr><td align="right">5</td><td>docs_url</td>
        <td><a href="https://mentor.ieee.org/802.11/documents?n=1&amp;o=7d&amp;is_group=00be&amp;is_year=2024">Documents</a></td></tr>
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
