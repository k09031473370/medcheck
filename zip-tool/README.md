# かんたん圧縮ツール

フォルダを **パスワード付き ZIP（AES-256 暗号化）** に圧縮・展開するシンプルなツールです。

## 準備

```
pip install pyzipper
```

（Python 3.8 以上。Windows / Mac の公式 Python なら GUI 部品 tkinter は同梱済み）

## 使い方

### 画面で使う

```
python zip_tool.py
```

- **圧縮タブ**: フォルダを選ぶ → パスワードを2回入力 → 「圧縮する」
- **展開タブ**: ZIP を選ぶ → パスワード入力 → 「展開する」

パスワードを空欄にすると暗号化なしの普通の ZIP になります。

### コマンドで使う

```
python zip_tool.py compress 送るフォルダ 送るフォルダ.zip
python zip_tool.py extract  送るフォルダ.zip 展開先フォルダ
```

パスワードは実行後に聞かれます（`-p パスワード` でも指定可）。

## exe にする（Windows）

```
pip install pyinstaller
pyinstaller --onefile --noconsole --name かんたん圧縮 zip_tool.py
```

`dist/かんたん圧縮.exe` ができます。Python のない PC でもダブルクリックで動きます。

## 注意

- AES 暗号化 ZIP は **Windows 標準の「すべて展開」では開けません**。
  相手側はこのツール、または 7-Zip / Lhaplus / CubeICE などで開いてください。
- パスワードを忘れると中身は取り出せません。
