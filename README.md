# SI Attribution Checklist Word Exporter

新規化合物データの帰属チェックリストを、Supporting Information 用の Word 文章へ整形するための Python スクリプトとテンプレートです。

Chem-Station 記事:
<https://www.chem-station.com/blog/2025/09/macro2.html>

## Contents

- `ExceltoWord2.py`: Excel ファイルから `*_SI.docx` を生成する Python スクリプト
- `templates/帰属チェックリスト_ver17.xlsx`: 公開用にメタデータを整理したテンプレート
- `docs/automator-quick-action.md`: macOS で右クリック実行するための設定手順
- `docs/windows-exe.md`: Windows 用 `.exe` の使い方メモ
- `reference/*.pdf`: 既存マニュアルの参照用 PDF

## Requirements

- Python 3.10 以降を推奨
- Microsoft Word は不要
- 必要パッケージ:

```bash
pip install -r requirements.txt
```

## Quick Start

### 1. 仮想環境を作成する

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2. テンプレート Excel を編集する

- `templates/帰属チェックリスト_ver17.xlsx` をコピーして使用してください
- Python スクリプトは `Python` シートを読みます
- 基本形式は `A列 = データタイトル`, `B列 = データ内容` です

### 3. 実行する

macOS / Linux:

```bash
python3 ExceltoWord2.py /path/to/帰属チェックリスト.xlsx
```

Windows:

```powershell
python .\ExceltoWord2.py .\帰属チェックリスト.xlsx
```

成功すると、入力 Excel と同じ場所に `ファイル名_SI.docx` が保存されます。

## What The Script Does

このスクリプトは次の処理を行います。

1. `Python` シートのデータを読み込む
2. タイトル太字、化学式の下付き、核種番号の上付き、`m/z` や `J` の体裁を自動整形する
3. Word 形式の段落として出力する
4. `1H NMR`, `13C NMR`, `HRMS` を参照し、H 数・C 数の整合性チェックを末尾へ付加する

## macOS で右クリック実行したい場合

Automator / Quick Action を使う方法を [`docs/automator-quick-action.md`](docs/automator-quick-action.md) にまとめています。

## Windows で `.exe` を使いたい場合

このリポジトリには 484 MB の Windows 配布フォルダは含めていません。操作方法のみ [`docs/windows-exe.md`](docs/windows-exe.md) と `reference/ExceltoWord_manual.pdf` に残しています。

## Notes

- テンプレート Excel は公開前に内部の絶対パスと作成者メタデータを整理しています
- 生成物の `*_SI.docx` や一時ファイルは Git 管理しない想定です
- ライセンスは未設定です。必要に応じて追加してください
