# SI Attribution Checklist Word Exporter

新規化合物データの帰属チェックリストを、Supporting Information 用の Word 文章へ整形するための Python スクリプトとテンプレートです。

## Windows Users

Windows で `.exe` を使いたい場合は、ソースコードではなく Release からダウンロードしてください。

- Latest Release: [Download from Releases](https://github.com/machub31/si-checklist-word-exporter/releases/latest)
- Direct ZIP: [exe app for windows.zip](https://github.com/machub31/si-checklist-word-exporter/releases/latest/download/exe%20app%20for%20windows.zip)

使い方:
1. `exe app for windows.zip` をダウンロードする
2. zip を展開して `ExceltoWord` フォルダを取り出す
3. `ExceltoWord.exe` に `.xlsx` ファイルをドラッグ&ドロップする
4. 入力 Excel と同じ場所に `*_SI.docx` が作成される

## Mac Users

Mac では Finder の右クリックから使えるように、Automator の Quick Action を設定してください。

- Setup guide: [docs/automator-quick-action.md](docs/automator-quick-action.md)
- Reference PDF: [reference/Automator_Quick_Action_Setup_Mac.pdf](reference/Automator_Quick_Action_Setup_Mac.pdf)

使い方:
1. `docs/automator-quick-action.md` を見ながら Quick Action を1回設定する
2. 帰属チェック済みの `.xlsx` を右クリックする
3. `Quick Action` から保存した名前を選ぶ
4. 入力 Excel と同じ場所に `*_SI.docx` が作成される

Chem-Station 記事:
<https://www.chem-station.com/blog/2025/09/macro2.html>

## Contents

- `ExceltoWord2.py`: Excel ファイルから `*_SI.docx` を生成する Python スクリプト
- `templates/CheckList_rev_18.xlsx`: 公開用にメタデータを整理したテンプレート
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

### 1. テンプレート Excel を編集する

- `templates/CheckList_rev_18.xlsx` をコピーして使用してください
- Python スクリプトは `Python` シートを読みます
- 基本形式は `A列 = データタイトル`, `B列 = データ内容` です

### 2. 実行する

- Windows 利用者は Releases の `exe app for windows.zip` を使ってください
- Mac 利用者は Automator の Quick Action を設定して右クリック実行してください

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

利用者は Releases から `exe app for windows.zip` をダウンロードしてください。

- Release page: [Latest Release](https://github.com/machub31/si-checklist-word-exporter/releases/latest)
- Direct download: [exe app for windows.zip](https://github.com/machub31/si-checklist-word-exporter/releases/latest/download/exe%20app%20for%20windows.zip)

配布者向けのビルド手順も含めた説明は [`docs/windows-exe.md`](docs/windows-exe.md) を参照してください。

## Developer Notes

コマンドラインから直接使いたい場合は、依存関係を入れた上で次のように実行できます。

```bash
python3 ExceltoWord2.py /path/to/checklist.xlsx
```

## Notes

- テンプレート Excel は公開前に内部の絶対パスと作成者メタデータを整理しています
- 生成物の `*_SI.docx` や一時ファイルは Git 管理しない想定です
- ライセンスは未設定です。必要に応じて追加してください
