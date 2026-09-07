<#
  判定記号(A/B/C/D…)の使われ方を調べる (check36.ps1)
  check36.bat をダブルクリックすると実行され、結果 r_check36.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  東振協のファイルには判定(A/B/C/D)の列があります。
  それを健診ナビの判定欄にそのまま書いてよいかを確かめます。
  健診ナビ側が C3/C6/C12 のような書き方なら、C は使えないので変換が要ります。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check36.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 判定記号の調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★健診ナビで実際に使われている判定記号 (多い順) ---' @"
SELECT TOP 40 '[' + LTRIM(RTRIM(k.HANTEI_KIGO)) + ']' AS 判定記号, COUNT(*) AS 件数
FROM T_KENSA k
WHERE LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
GROUP BY LTRIM(RTRIM(k.HANTEI_KIGO))
ORDER BY COUNT(*) DESC
"@ 45

Q '--- 2. ★所見の項目だけで見た判定記号 (胸部/胃部/心電図/診察) ---' @"
SELECT LEFT(LTRIM(RTRIM(k.KOMOKU_CD)), 6) AS 項目,
       '[' + LTRIM(RTRIM(k.HANTEI_KIGO)) + ']' AS 判定記号,
       COUNT(*) AS 件数
FROM T_KENSA k
WHERE LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
  AND (   LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '077010%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '077300%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '067112%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '017101%')
GROUP BY LEFT(LTRIM(RTRIM(k.KOMOKU_CD)), 6), LTRIM(RTRIM(k.HANTEI_KIGO))
ORDER BY 1, 3 DESC
"@ 60

Q '--- 3. 判定記号のマスタらしき表を探す ---' @"
SELECT c.TABLE_NAME AS テーブル, c.COLUMN_NAME AS 列
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.COLUMN_NAME LIKE '%HANTEI%'
ORDER BY c.TABLE_NAME, c.COLUMN_NAME
"@ 80

Q '--- 4. ★8/21 自動判定が回っているか (1人あたり判定が付いた項目の数) ---' @"
SELECT x.判定数 AS 判定が付いた項目数, COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k
           WHERE k.PK_SEQ = s.PK_SEQ
             AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '') AS 判定数
  FROM T_KENSIN s
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
GROUP BY x.判定数
ORDER BY x.判定数
"@ 40

Q '--- 5. くらべる: 院内の日 (8/20) は1人あたりいくつ判定が付いているか ---' @"
SELECT x.判定数 AS 判定が付いた項目数, COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k
           WHERE k.PK_SEQ = s.PK_SEQ
             AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '') AS 判定数
  FROM T_KENSIN s
  WHERE s.D_KENSIN = '2026/08/20' AND s.F_TORIKESI = 0
) x
GROUP BY x.判定数
ORDER BY x.判定数
"@ 40

W ''
W '=== 読み方 ==='
W '  1・2 に [C] が出ていれば、東振協の C をそのまま書いて大丈夫です。'
W '     [C3] [C6] [C12] しか無くて [C] が無いなら、C をどれに読み替えるかを'
W '     決めないと書けません (これは医療上の判断なので、こちらでは決められません)。'
W ''
W '  4 で「判定が付いた項目数」が 5〜10 のあたりに88人が固まっていれば、'
W '     8/21 はまだ自動判定が回っていないということです。'
W '     5 の院内の日 (30〜60項目くらいのはず) とくらべてください。'

notepad $out
