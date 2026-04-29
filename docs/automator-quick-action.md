# macOS Automator Quick Action

`ExceltoWord2.py` を Finder の右クリックから実行する設定例です。

## 手順

1. `Automator.app` を開く
2. `New Document` を選ぶ
3. `Quick Action` を選ぶ
4. 検索バーに `shell` と入力し、`Run Shell Script` を追加する
5. 次のスクリプトを貼り付ける

```bash
for f in "$@"
do
    /path/to/python3 /path/to/ExceltoWord2.py "$f"
done
```

6. `Command + S` で保存する
   例: `Run ExceltoWord`
7. 帰属チェック済みの `.xlsx` を右クリックし、`Quick Action` から保存した名前を選ぶ

## Path の確認方法

### `python3` の path

Terminal で次を実行します。

```bash
which python3
```

環境によっては `python` のみ使える場合があります。

### `ExceltoWord2.py` の path

1. Finder で `ExceltoWord2.py` を右クリック
2. `Get Info` を開く
3. `Where` を確認する
4. 必要なら `Option` を押しながら右クリックしてパスをコピーする

## 補足

- Excel ファイルは閉じた状態で実行してください
- 仮想環境を使う場合は、その環境内の `python3` を指定してください
- 既存 PDF は `reference/Automator_Quick_Action_Setup_Mac.pdf` にあります
