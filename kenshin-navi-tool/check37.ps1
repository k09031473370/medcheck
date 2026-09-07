<#
  自動判定の仕組みを調べる (check37.ps1)
  check37.bat をダブルクリックすると実行され、結果 r_check37.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  健診ナビの自動判定が1人ずつしかできないため、
  同じ判定をこちらで計算して一括で入れられないかを検討します。

  そのためには、まず健診ナビが何を見て判定しているかを知る必要があります。
  基準値マスタ (T_KIJUN1 / T_KIJUN2) がその答えのはずです。
#>
[CmdletBinding()]
param(
    [string]$Ymd  = '2026/08/21',   # 判定がまだの日
    [string]$Ymd2 = '2026/08/20'    # 判定が済んでいる日 (くらべる用)
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check37.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 自動判定の仕組みの調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★基準値マスタ T_KIJUN1 の列 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KIJUN1'
ORDER BY ORDINAL_POSITION
"@ 60

Q '--- 2. ★基準値マスタ T_KIJUN2 の列 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KIJUN2'
ORDER BY ORDINAL_POSITION
"@ 60

Q '--- 3. ★中身を見る: 血色素量(069005) の基準値 ---' @"
SELECT TOP 30 * FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '069005'
"@ 35

Q '--- 4. ★中身を見る: ALT/GPT(068637) の基準値 ---' @"
SELECT TOP 30 * FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '068637'
"@ 35

Q '--- 5. T_KIJUN1 の中身 (先頭20行) ---' @"
SELECT TOP 20 * FROM T_KIJUN1 ORDER BY KOMOKU_CD
"@ 25

Q '--- 6. 基準値が登録されている項目数 ---' @"
SELECT 'T_KIJUN1' AS 表, COUNT(DISTINCT LTRIM(RTRIM(KOMOKU_CD))) AS 項目数, COUNT(*) AS 行数 FROM T_KIJUN1
UNION ALL
SELECT 'T_KIJUN2', COUNT(DISTINCT LTRIM(RTRIM(KOMOKU_CD))), COUNT(*) FROM T_KIJUN2
"@ 10

Q '--- 7. 判定記号のマスタ T_HANTEIS ---' @"
SELECT TOP 40 * FROM T_HANTEIS
"@ 45

Q '--- 8. ★自動判定を実行した日付が残っているか (D_JIDOHANTEI) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       COUNT(*) AS 人数,
       SUM(CASE WHEN s.D_JIDOHANTEI IS NULL THEN 1 ELSE 0 END) AS 判定未実行,
       MAX(CONVERT(varchar(20), s.D_JIDOHANTEI, 120)) AS 最後に実行した日時
FROM T_KENSIN s
WHERE s.D_KENSIN IN ('$Ymd', '$Ymd2') AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN
ORDER BY 1
"@ 10

Q '--- 9. ★答え合わせ用: 判定済みの日(8/20)の実際の値と判定 ---' @"
SELECT TOP 60 LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       k.KEKKA AS 結果, '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.HANTEI_CD),''))) + ']' AS 判定CD
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd2' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('069005','068637','001253','001254','069207','069206')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), k.KEKKA, k.HANTEI_KIGO, k.HANTEI_CD
ORDER BY 1, 3
"@ 65

W ''
W '=== 読み方 ==='
W '  1〜5 で、健診ナビが「値がいくつなら判定は何か」をどう持っているかが分かります。'
W '     性別・年齢で基準が変わる作りなら、その列も出てきます。'
W ''
W '  8 で 8/21 が「判定未実行 88人」なら、やはり回っていません。'
W ''
W '  9 は答え合わせに使います。'
W '     こちらで計算した判定が、この表と全部一致するなら'
W '     一括で入れるツールを作っても安全だと言えます。'
W '     一致しないものが1つでもあれば作りません。'

notepad $out
