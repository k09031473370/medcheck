<#
  「社員番号」がDBのどの欄に入るのかを突き止める (check11.ps1)
  check11.bat をダブルクリックすると実行され、結果 r_check11.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  分かっていること:
    帳票の差込み文字 **社員番号 は健診ナビが用意している
    (退避\701_新規 Microsoft Excel Worksheet.xlsx が差込み文字の一覧になっている)。
    健診ナビの「個人マスタ」画面にも 社員番号 と 部署 の入力欄がある。
    残るは、それがDBのどの欄に入るのか。

  やりかた:
    健診ナビの個人マスタで、テスト用の人1人の「社員番号」に
    目印の値 (既定 99999) を入れて登録してから、これを実行する。
    DB全体の文字列の欄からその値を探して、入った場所を特定する。

  ※ 前の版は動的SQLを使っていて db_tool.ps1 に弾かれた。
     db_tool.ps1 にある -FindText (全テーブル文字列検索) を使う形に直した。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check11.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }

Write-Host ''
Write-Host '健診ナビの個人マスタで「社員番号」に入れた値を教えてください。'
$val = Read-Host '探す値 [既定 99999]'
if ([string]::IsNullOrWhiteSpace($val)) { $val = '99999' }
$val = $val.Trim()

Write-Host ''
Write-Host 'DB全体を探します。1〜2分かかります…'

"=== 社員番号の入り先さがし $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W "探す値: $val"

if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

W ''
W '--- 1. その値が入っている欄 ---'
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -FindText $val -MaxRows 5 *>&1 |
    Out-File $out -Append -Encoding Default

W ''
W '※ 受付番号など、たまたま同じ数字が入っているだけの欄も出ます。'
W '※ 個人マスタ(T_KOJIN1 / T_KOJIN2)に出てくる欄が、社員番号の入り先です。'

# 部署の入り先も一緒に確かめる (個人マスタの社員番号のすぐ上にある欄)
W ''
W '--- 2. 部署まわり (T_BUSYO と、個人マスタの部署らしい欄) ---'
$sql = @"
SELECT TOP 20 j.KOJIN_ID AS 管理ID, j.KANJI_SIMEI AS 氏名,
       j.KOJIN_NO AS 個人番号, j.KARUTE_NO AS カルテNo, j.TECHO_NO AS 手帳No,
       j2.SYUSSEKI_NO AS 出席番号, j2.DIVISION AS DIVISION,
       j2.GYOSYU AS 業種, j2.CLASS AS CLASS, j2.GAKUNEN AS 学年
FROM T_KOJIN1 j LEFT JOIN T_KOJIN2 j2 ON j2.KOJIN_ID = j.KOJIN_ID
WHERE j.KANA_SIMEI LIKE N'ﾃｽﾄ%' OR j.KANA_SIMEI LIKE N'テスト%' OR j.KANJI_SIMEI LIKE N'テスト%'
ORDER BY j.KOJIN_ID DESC
"@
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows 25 *>&1 |
    Out-File $out -Append -Encoding Default

W ''
W '--- 3. 個人マスタ T_KOJIN1 の列一覧 (社員番号がこちらにある場合の確認用) ---'
$sql2 = @"
SELECT COLUMN_NAME AS COL, DATA_TYPE AS TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'T_KOJIN1'
ORDER BY ORDINAL_POSITION
"@
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql2 -MaxRows 80 *>&1 |
    Out-File $out -Append -Encoding Default

W ''
W '=== 完了。この内容をチャットに貼り付けてください ==='
notepad $out
