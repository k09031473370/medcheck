<#
  総合判定 (B/C/D/E/F/G) ごとに、どの項目で引っかかっているかを調べる (check55.ps1)
  check55.bat をダブルクリックすると実行され、結果 r_check55.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  総合判定には「各判定のうち最も重いもの」が入る。
  つまり総合判定が D の人は、どこかの項目が D になっている。それを判定別に集計する。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check55.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 総合判定ごとの原因項目 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 総合判定(編集) 11000 を持つ人
$SG = @"
(SELECT s.PK_SEQ, LTRIM(RTRIM(ISNULL(t.HANTEI_KIGO,''))) AS SG
 FROM T_KENSIN s
 JOIN T_KENSA t ON t.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(t.KOMOKU_CD)) = '11000'
 WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
   AND LTRIM(RTRIM(ISNULL(t.HANTEI_KIGO,''))) <> '') x
"@

Q '--- 1. 総合判定の分布 ---' @"
SELECT x.SG AS 総合判定, COUNT(*) AS 人数
FROM $SG
GROUP BY x.SG ORDER BY x.SG
"@ 20

Q '--- 2. ★総合判定ごとに、その判定を持つ「くくり判定」(600番台) ---' @"
SELECT x.SG AS 総合判定, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       MAX(m.MEISYO1) AS くくり判定, COUNT(*) AS 人数
FROM $SG
JOIN T_KENSA k ON k.PK_SEQ = x.PK_SEQ
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = x.SG
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600%'
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
GROUP BY x.SG, LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY x.SG, 4 DESC
"@ 120

Q '--- 3. ★★総合判定ごとに、その判定を持つ「個別の検査項目」 ---' @"
SELECT x.SG AS 総合判定, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       MAX(m.MEISYO1) AS 検査項目, COUNT(*) AS 人数,
       MIN(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 値の例1,
       MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 値の例2
FROM $SG
JOIN T_KENSA k ON k.PK_SEQ = x.PK_SEQ
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = x.SG
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '600%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%'
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
GROUP BY x.SG, LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY x.SG, 4 DESC
"@ 200

Q '--- 4. ★G(要治療)の人の明細 ---' @"
SELECT g.KANJI_SIMEI AS 氏名, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 検査項目,
       '[' + LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 30) + ']' AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定
FROM $SG
JOIN T_KENSIN s      ON s.PK_SEQ   = x.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = x.PK_SEQ
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('', 'A', 'B')
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE x.SG = 'G'
ORDER BY g.KANA_SIMEI, k.KOMOKU_CD
"@ 60

Q '--- 5. 参考: 全員(88人)の項目別 判定分布 (A/B以外がある項目だけ) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 検査項目,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'A' THEN 1 ELSE 0 END) AS A,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'B' THEN 1 ELSE 0 END) AS B,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'C' THEN 1 ELSE 0 END) AS C,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'D' THEN 1 ELSE 0 END) AS D,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'E' THEN 1 ELSE 0 END) AS E,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'F' THEN 1 ELSE 0 END) AS F,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('G','J') THEN 1 ELSE 0 END) AS GJ
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('','A','B') THEN 1 ELSE 0 END) > 0
ORDER BY 8 DESC, 7 DESC, 6 DESC, 1
"@ 120

W ''
W '=== 読み方 ==='
W '  2 で「総合判定Dの人は、どの分野がDなのか」が分かります。'
W '  3 が本体です。判定ごとに、どの検査値・所見が原因かが人数つきで出ます。'
W '  5 は全員分の一覧です。「視力は 88人中 何人がF か」のような見方ができます。'
W '     判定を見直すかどうかを決めるときの材料になります。'
notepad $out
