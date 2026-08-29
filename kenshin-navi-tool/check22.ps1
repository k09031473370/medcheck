<#
  自動判定エラーの原因さがし その3: 判定基準マスタの中身を見る (check22.ps1)
  check22.bat をダブルクリックすると実行され、結果 r_check22.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check21 で分かったこと:
    コース YSA1(定期健診城西) は実データ97件すべて結果値ゼロ。
    = このコースで「自動判定」が実行されたことは過去に一度もない。
    → YSA1 用の判定基準マスタに、形の崩れた行・空の行が隠れている疑いが濃厚。
      (自動判定は基準の文字列を Substring で切り出すので、短い・空だと今回のエラーになる)

  ここでやること:
    T_KIJUN1 / T_KIJUN2 / T_HANTEIS / M_METAKIJUN_R6 の構造と中身を見て、
    どの行が Substring で切れない形をしているかの手がかりを集める。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check22.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 判定基準マスタ調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 基準マスタ4つの列構成 ---' @"
SELECT TABLE_NAME AS テーブル, ORDINAL_POSITION AS 順, COLUMN_NAME AS 列,
       DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME IN ('T_KIJUN1','T_KIJUN2','T_HANTEIS','M_METAKIJUN_R6')
ORDER BY TABLE_NAME, ORDINAL_POSITION
"@ 200

Q '--- 2. T_KIJUN1 の中身 (先頭30行) ---' @"
SELECT TOP 30 * FROM T_KIJUN1
"@ 40

Q '--- 3. T_KIJUN2 の中身 (先頭30行) ---' @"
SELECT TOP 30 * FROM T_KIJUN2
"@ 40

Q '--- 4. T_HANTEIS の中身 (先頭30行) ---' @"
SELECT TOP 30 * FROM T_HANTEIS
"@ 40

Q '--- 5. 行数の見当 ---' @"
SELECT 'T_KIJUN1' AS テーブル, COUNT(*) AS 行数 FROM T_KIJUN1
UNION ALL SELECT 'T_KIJUN2', COUNT(*) FROM T_KIJUN2
UNION ALL SELECT 'T_HANTEIS', COUNT(*) FROM T_HANTEIS
UNION ALL SELECT 'M_METAKIJUN_R6', COUNT(*) FROM M_METAKIJUN_R6
"@ 10

# コース列を持っていれば YSA1 の行だけも見る (列が無いテーブルではスキップされる)
Q '--- 6. T_KIJUN1 に COURSE_CD 列があれば YSA1 の行 ---' @"
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_NAME='T_KIJUN1' AND COLUMN_NAME='COURSE_CD')
    EXEC('SELECT TOP 60 * FROM T_KIJUN1 WHERE LTRIM(RTRIM(COURSE_CD)) = ''YSA1''')
ELSE
    SELECT 'T_KIJUN1 に COURSE_CD 列は無い' AS 結果
"@ 80

Q '--- 7. T_KIJUN2 に COURSE_CD 列があれば YSA1 の行 ---' @"
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_NAME='T_KIJUN2' AND COLUMN_NAME='COURSE_CD')
    EXEC('SELECT TOP 60 * FROM T_KIJUN2 WHERE LTRIM(RTRIM(COURSE_CD)) = ''YSA1''')
ELSE
    SELECT 'T_KIJUN2 に COURSE_CD 列は無い' AS 結果
"@ 80

# 城西学園の団体・コース設定側に判定グループらしき列が無いかも見ておく
Q '--- 8. コースマスタ側の YSA1 の設定 (判定グループなどの持ち方) ---' @"
SELECT TOP 10 * FROM T_COURSE1 WHERE LTRIM(RTRIM(COURSE_CD)) = 'YSA1'
"@ 20

W ''
W '=== 読み方 ==='
W '  これは手がかり集めの回です。1 で基準の文字列がどの列に入っているかが分かり、'
W '  2〜4 の中身で「基準値の文字列がどういう形式か」(例: 000130085 のような固定桁) が見えます。'
W '  固定桁の列に、桁の足りない行・空の行があれば、それが自動判定を落としている行です。'
W '  6・7 で YSA1 専用の基準行が出れば、他コースの行と見比べて崩れを探します。'
W '  r_check22.txt をそのまま送ってください。次で犯人の行まで絞り込みます。'

notepad $out
