# PDF 奇数・偶数ページ分割ツール

複数の PDF を 1 ページずつの PDF に分割し、奇数ページを `odd/`、偶数ページを `even/` に振り分けます。
出力ファイル名は `元ファイル名_p001.pdf` のようにページ番号付きです。

## ブラウザ版(インストール不要)

1. `split_odd_even.html` をダウンロードしてダブルクリックで開く(Chrome / Edge / Safari など)。
2. PDF をまとめて選択(またはドラッグ&ドロップ)。
3. 「分割して ZIP をダウンロード」を押すと `split_odd_even.zip` が保存される。
   ZIP を展開すると `odd/` と `even/` フォルダが入っている。

処理はすべてブラウザ内で完結し、PDF は外部に送信されません。
(pdf-lib / JSZip を cdnjs から読み込むためインターネット接続が必要です。)

## Python 版

```bash
pip install pypdf
python3 split_odd_even.py 入力1.pdf 入力2.pdf ... -o 出力先
# 例: python3 split_odd_even.py *.pdf -o out
```

出力先に `odd/` と `even/` が作られます。
