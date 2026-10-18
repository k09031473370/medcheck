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

## 自己解凍形式（相手はソフト不要）

exe 版で「自己解凍形式」にチェックを入れて圧縮すると、`.exe` ファイルができます。
受け取った人は **ダブルクリック → パスワード入力** だけで、exe と同じ場所に展開され、フォルダが開きます。

- 相手が **Windows** の場合のみ使えます（Mac / スマホは不可）
- ファイルサイズが約 10MB 増えます
- **メール添付では exe はほぼ確実にブロックされます**（Gmail・Outlook など。ZIP に入れても不可）。
  ギガファイル便・Google ドライブ・OneDrive などの共有リンクで渡してください
- 署名のない exe なので、初回に「Windows によって PC が保護されました」と出ることがあります。
  「詳細情報」→「実行」で開けます。相手にあらかじめ伝えておくと安心です

コマンドの場合: `かんたん圧縮.exe compress フォルダ 送る.exe --sfx`

## 注意

- AES 暗号化 ZIP は **Windows 標準の「すべて展開」では開けません**。
  相手側はこのツール、または 7-Zip / Lhaplus / CubeICE などで開いてください。
- パスワードを忘れると中身は取り出せません。
