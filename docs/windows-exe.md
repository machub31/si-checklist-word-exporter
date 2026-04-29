# Windows `.exe` 版メモ

## 利用者向け

通常の利用者は GitHub Releases から zip をダウンロードしてください。

- Release page: [Latest Release](https://github.com/machub31/si-checklist-word-exporter/releases/latest)
- Direct download: [exe app for windows.zip](https://github.com/machub31/si-checklist-word-exporter/releases/latest/download/exe%20app%20for%20windows.zip)

### 使い方

1. `exe app for windows.zip` をダウンロードする
2. zip を展開して `ExceltoWord` フォルダを取り出す
3. `ExceltoWord.exe` に、文章化したい帰属チェックリスト `.xlsx` をドラッグ&ドロップする
4. 入力 Excel と同じ場所に `*_SI.docx` が出力される

### 注意

- `.exe` 単体では動作しません。展開後の `ExceltoWord` フォルダ一式が必要です
- 64-bit Windows 向けです
- macOS では使えません

## 配布者向け

### GitHub での作り方

このリポジトリには GitHub Actions の workflow を入れてあります。

1. `main` に push する
2. GitHub の `Actions` タブで `Build Windows EXE` を開く
3. 成功した run から artifact `exe app for windows` をダウンロードする
4. Release に `exe app for windows.zip` を添付する

### 配布時の注意

- リポジトリ本体には巨大な配布フォルダは commit しません
- 配布は Release 添付を基本にしてください
- artifact は配布者の中間成果物、Release は利用者向けの公開窓口です

詳細は `reference/ExceltoWord_manual.pdf` を参照してください。
