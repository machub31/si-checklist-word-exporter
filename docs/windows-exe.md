# Windows `.exe` 版メモ

## GitHub での作り方

このリポジトリには GitHub Actions の workflow を入れてあります。

1. `main` に push する
2. GitHub の `Actions` タブで `Build Windows EXE` を開く
3. 成功した run から artifact `exe app for windows` をダウンロードする
4. 展開すると `exe app for windows.zip` 相当の中身として `ExceltoWord/` フォルダが得られる

## 使い方

1. `ExceltoWord` フォルダ一式をローカル PC にコピーする
2. `ExceltoWord.exe` に、文章化したい帰属チェックリスト `.xlsx` をドラッグ&ドロップする
3. 入力 Excel と同じ場所に `*_SI.docx` が出力される

## 注意

- `.exe` 単体では動作しません。配布フォルダごと必要です
- 32-bit Windows や macOS では使えません
- リポジトリ本体には巨大な配布フォルダは commit しません
- 配布は artifact または GitHub Release 添付で扱うのが前提です

詳細は `reference/ExceltoWord_manual.pdf` を参照してください。
