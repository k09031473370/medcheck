# 健診ナビ11 検査結果一括取込ツール (form_import)

自院Excelフォーム(1行目ヘッダ・105列・1人1行)をCSV保存したものを読み込み、
健診ナビ(KENSHIN navi 11)のDBへ直接 UPDATE で検査結果を書き込むツールです。

- DB: SQL Server `KNSV\SQLEXPRESS` / `K166_SIBAURAUSER`
- 接続情報: `\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt` を自動で読む
  (解釈できない場合は `-ConnectionString "..."` で直接指定)
- 人の特定: `T_KANJA_G` の `KEN_YMD`(受診日 'YYYY/MM/DD') + `KEN_NO`(受付番号) → `PK_SEQ`
- 書込: `UPDATE T_KENSA SET KEKKA=... WHERE PK_SEQ=... AND KOMOKU_CD=...`
  (枠は予約時に作成済みの前提。INSERTはしない=枠が無い項目は「枠なし」エラーになるだけで安全)
- 所見項目: `KEKKA`=所見文 / `KEKKA_CD`=結果CD / `HANTEI_KIGO`=判定 を
  所見マスタ `T_SYOKEN2` と突合して書込
- **書込前に対象者の T_KENSA 全行を `backup\` にCSV保存**(自動)
- BMI・判定は書込後に健診ナビの「自動判定」で計算してください

## ファイル構成

```
C:\kenshin-navi\
  form_import.ps1        … 本体 (コマンドライン)
  form_import_gui.ps1    … GUI版 (画面から操作。中身は本体を呼び出すだけ)
  ecg_import.ps1         … 心電図結果CSV取込 (検査日,ID,判定記号,所見コード1-5形式)
  db_tool.ps1            … DB調査・掃除ツール
  取込マニュアル.md / .docx … スタッフ向け運用マニュアル
  form\
    mapping.csv          … 列番号 ⇔ KOMOKU_CD 対応表 (要・現地で追記)
    value_map.csv        … 値の変換表 (尿定性・聴力など)
    ecg_items.csv        … 心電図所見1〜5の KOMOKU_CD (要・現地で追記)
    ecg_code_map.csv     … 心電計の所見コード → KEKKA_CD 変換表 (要・現地で追記)
  backup\                … 書込前バックアップ (自動生成)
```

### 心電図取込 (ecg_import.ps1) の初期設定

1. `ecg_import.ps1 -Csv <ecg.csv> -ListCodes` … CSV内の装置コード一覧と変換表の充足を確認 (DB不要)
2. `form_import.ps1 -DumpSyoken ZK011` … 健診ナビ側の心電図所見マスタ一覧を表示
3. 1と2を突き合わせて `form\ecg_code_map.csv` の KEKKA_CD を記入
   (NONE行=所見なし・判定Aの人に書く「異常なし」のKEKKA_CD)
4. `form_import.ps1 -Only <受付番号> -KenYmd <受診日> -DumpItems` で心電図所見1〜5の
   KOMOKU_CD を確認し `form\ecg_items.csv` に記入
5. 判定も取り込む場合は、心電図判定の KOMOKU_CD を `ecg_items.csv` の HANTEI 行に記入
   (装置の判定記号 A/B/C12 等を KEKKA と HANTEI_KIGO にそのまま書込。KEKKA_CD は触らない。
   初回は実機で1件手入力→ -DumpItems で判定の格納形式を確認しておくと確実)

## GUI版 (form_import_gui.ps1)

コマンドを打たずに画面から操作できます。**Excelファイル(.xlsx)を直接指定可能**
(PCのExcelを使って自動でCSVに変換するため、手動のCSV保存が不要)。

起動:
```powershell
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import_gui.ps1
```
デスクトップにショートカットを作る場合の「リンク先」:
```
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\kenshin-navi\form_import_gui.ps1
```

画面のボタンは番号順に使います:
1. **列確認** … CSVの列番号一覧 (mapping.csv の Col 記入用。DB接続なし)
2. **枠一覧(DB)** … 対象者の T_KENSA (KOMOKU_CD 記入用)
3. **プレビュー** … 書込内容の確認 (DBは読むだけ)
4. **書込実行** … 確認ダイアログの後にDB書込

GUI版・コマンド版のどちらを使っても処理は同一です(GUIは form_import.ps1 を呼ぶだけ)。

## 事務PCへの配置

このフォルダ(`kenshin-navi-tool`)の中身をそのまま `C:\kenshin-navi\` にコピーしてください。
CSVはUTF-8(BOM付き)です。**mapping.csv をExcelで編集すると問診の KOMOKU_CD `084001` の
先頭ゼロが消えることがあるため、メモ帳やVSCodeでの編集を推奨**します。

## 初回セットアップ手順 (対応表の穴埋め)

mapping.csv は「列番号(Col)」と「KOMOKU_CD」の一部が未記入です(環境依存のため)。
以下の手順で15分程度で埋められます。

### 1. フォームをCSV保存

Excelでフォームを開き「名前を付けて保存」→「CSV (コンマ区切り) (*.csv)」。
(文字コードは既定のANSI=Shift-JISでOK。「CSV UTF-8」で保存した場合は実行時に `-CsvEncoding UTF8` を付ける)

### 2. CSVの列番号を確認 (-Inspect / DB接続なし)

```powershell
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -Csv C:\...\form.csv -Inspect
```

全105列の「列番号・ヘッダ名・1人目の値」が一覧表示されます。
これを見ながら mapping.csv の `Col` 列を埋めます
(受付番号 KENNO 行は必須。身長〜問診22、および 診察50-54 / 自覚症状57-59 /
胸部X線96-100 / 心電図101-105 の並びが合っているかも確認)。

### 3. KOMOKU_CD を確認 (-DumpItems)

```powershell
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -Only 4001 -KenYmd 2026/07/02 -DumpItems
```

対象者の T_KENSA の全枠(KOMOKU_CD と現在値)が表示され、全列ダンプCSVも `backup\` に保存されます。
これを見ながら mapping.csv の `KOMOKU_CD` 列を埋めます(問診 084001〜084022 は記入済み)。
どのコードがどの項目か分からない場合は、健診ナビの画面で1項目だけ手入力(例: 身長に 999)
→ もう一度 -DumpItems を実行すると、値が入った行=その項目のコードだと分かります。

### 4. 所見マスタの確認 (-DumpSyoken)

フォームの所見文がマスタの表記と完全一致しないと取り込めないため、事前に確認します。

```powershell
# 診察:SHIN / 眼底:GANTEI / 心電図:ZK011 / 腹部エコー:ZK041
# 胸部X線:ZK020(部位)+ZK021(所見) / 胃部X線:ZK030(部位)+ZK031(所見)
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -DumpSyoken SHIN
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -DumpSyoken ZK021
```

## 実行手順 (受付番号4001でテスト)

```powershell
# 1) プレビュー (書込なし)
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 `
  -Csv "C:\Users\User\Desktop\form.csv" -Only 4001 -KenYmd 2026/07/02

# 2) 内容を確認して問題なければ書込
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 `
  -Csv "C:\Users\User\Desktop\form.csv" -Only 4001 -KenYmd 2026/07/02 -Commit
```

- プレビューの「状態」列: `OK`=書込可能 / `列未設定`・`項目CD未設定`=対応表が未記入(スキップ) /
  `枠なし`=T_KENSAに行が無い / `変換不可`=値を変換できない / `所見未登録`=T_SYOKEN2に一致なし
- エラー行(枠なし・変換不可・所見未登録)があると書込は中止されます。
  修正できない項目を飛ばしてOK行だけ書き込むには `-Force` を付けます。
- 書込は1人分をトランザクションで実行(途中失敗時は全ロールバック)。
- `-Only` を付けなければCSV全行(全員)を処理します。

書込後は健診ナビで対象者を開いて **「自動判定」を実行**(BMI・各判定の計算)し、
画面で値を確認 → 東振協向け出力のテストへ。

## 胸部X線(2段階所見)についての注意

胸部X線・胃部X線は「部位(ZK020/ZK030)+所見(ZK021/ZK031)」の2段階です。
本ツールは現状、`KEKKA`=「部位文+空白+所見文」、`KEKKA_CD`・`HANTEI_KIGO`=所見側の値
として書き込みます。**健診ナビ実機での格納形式と一致するか、初回は必ず確認してください**:

1. 健診ナビの画面でテスト対象者に胸部X線所見を1件手入力
2. `-DumpItems` でその行の `KEKKA` / `KEKKA_CD` / `HANTEI_KIGO` の入り方を確認
3. 形式が違う場合(例: KEKKA_CDが部位CD+所見CDの連結 等)は form_import.ps1 の
   `Build-Plan` 内 SHOKEN2 の箇所を修正

診察(SHIN)・心電図(ZK011)は単一段階なのでそのまま動く想定です。

## 元に戻したいとき

書込のたびに `backup\T_KENSA_<PK_SEQ>_<日時>.csv` に書込前の全列が保存されます。
戻す場合はこのCSVの KEKKA / KEKKA_CD / HANTEI_KIGO を見ながら UPDATE で復元してください
(健診ナビの画面から手修正でも可)。

## テストで作った不要データの掃除 (db_tool.ps1)

アプリの「検体検査結果取込」を試した際に作った **検査センター「院内」** などを削除する手順です。
健診ナビのテーブル名は環境ごとに異なるため、まず検索で場所を特定します。

```powershell
# 1) 「院内」がどのテーブル・列に入っているか探す (DB全体を検索)
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -FindText 院内

# 2) 見つかったテーブルの中身を確認 (例: テーブルが M_CENTER、列が CENTER_NM だった場合)
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -Table M_CENTER

# 3) 削除プレビュー (まだ消えない)
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -Table M_CENTER -Where "CENTER_NM = N'院内'" -Delete

# 4) 実削除 (削除前に backup\ へ自動バックアップ)
powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -Table M_CENTER -Where "CENTER_NM = N'院内'" -Delete -Commit
```

注意:
- **可能ならまず健診ナビのマスタ保守画面から削除**を試してください(画面から消せればそれが一番安全)。
- **「江東微研(城南)」「江東微研(城西)」は本物のセンターなので絶対に消さない**こと。
  自分で作った「院内」など、心当たりのある行だけを対象にしてください。
- 削除前に、そのセンターのコード(例: 3)が依頼データ等から参照されていないかも
  `-FindText` やテーブル確認で見ておくと安全です(テスト取込データが残っていれば先にそちらを削除)。
- DB以外の痕跡: 取込テストで作ったフォルダ/ファイル (`C:\KTB\kekkadata` のテストファイル、
  `\\KNSV\KenshinNavi\log\検査結果取込\` のログ) は普通にエクスプローラーで削除できます。
- 迷ったら削除せず、`-FindText` の結果を貼って相談してください。

## トラブルシューティング

- **接続エラー**: `SQLServerConnect.txt` の形式が解釈できない場合は
  `-ConnectionString "Data Source=KNSV\SQLEXPRESS;Initial Catalog=K166_SIBAURAUSER;User ID=xxx;Password=xxx"`
  を直接指定(前回UPDATE成功時に使った接続文字列と同じものでOK)。
- **文字化け**: CSVを「CSV UTF-8」で保存した場合は `-CsvEncoding UTF8` を付ける。
- **受診者が見つからない**: 受診日の書式は 'YYYY/MM/DD'。`-KenYmd 2026/07/02` の指定と
  T_KANJA_G の KEN_YMD が一致しているか確認。
- 実行は健診ナビで対象者の画面を閉じた状態で行ってください(排他・表示の食い違い防止)。
