<#
  石屋に J (治療中) が 1 件も無いのはなぜか (check107.ps1)
  2026 年に くくり 血圧 600150 が J の人 (他団体) と石屋の人で、問診 服薬歴の入り方・J の付き方を比べる。
  check107.bat → r_check107.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check107.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋に J が無い理由 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$JBP = "EXISTS (SELECT 1 FROM T_KENSA j WHERE j.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD)) = '600150' AND LTRIM(RTRIM(ISNULL(j.HANTEI_KIGO,''))) = 'J')"

Q '--- 1. ★血圧が J の人 (2026 年・他団体) の 受診日 × コース × 人数 (最近のもの) ---' @"
SELECT TOP 30 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, COUNT(*) AS 人数
FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND $JBP
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), LTRIM(RTRIM(ISNULL(s.COURSE_CD,'')))
ORDER BY 1 DESC
"@ 30

Q '--- 2. ★服薬歴(血圧) 084001 の 結果 / 結果CD の入り方: 血圧 J の人 (他団体) と 石屋 ---' @"
SELECT 'J の人' AS 誰, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS 結果, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS 結果CD, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS 判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND $JBP AND LTRIM(RTRIM(k.KOMOKU_CD)) = '084001'
GROUP BY LTRIM(RTRIM(ISNULL(k.KEKKA,''))), LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))), LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))
UNION ALL
SELECT '石屋', LTRIM(RTRIM(ISNULL(k.KEKKA,''))), LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))), LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))), COUNT(*)
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = '084001'
GROUP BY LTRIM(RTRIM(ISNULL(k.KEKKA,''))), LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))), LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))
ORDER BY 1, 5 DESC
"@ 30

Q '--- 3. ★血圧が J の人 (他団体) で、J が付いている項目 (くくりだけか、測定値の項目にも付くか) ---' @"
SELECT TOP 30 LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, COUNT(*) AS J件数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND $JBP AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'J'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)) ORDER BY 3 DESC
"@ 30

Q '--- 4. 見本: 最近の 血圧 J の人 1 人の 問診 (084001〜084003) と 血圧の項目・くくり ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 結果, MAX(LTRIM(RTRIM(ISNULL(k.KEKKA_CD,'')))) AS 結果CD,
       MAX(LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))) AS 判定, MAX(CONVERT(varchar(10), s.D_KENSIN, 111)) AS 受診日
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.PK_SEQ = (SELECT TOP 1 s2.PK_SEQ FROM T_KENSIN s2 WHERE s2.F_TORIKESI = 0 AND s2.D_KENSIN >= '2026/08/01'
                  AND EXISTS (SELECT 1 FROM T_KENSA j WHERE j.PK_SEQ = s2.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD)) = '600150' AND LTRIM(RTRIM(ISNULL(j.HANTEI_KIGO,''))) = 'J')
                  ORDER BY s2.D_KENSIN DESC)
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) IN ('084001','084002','084003','600150','600340','600320','11000') OR m.MEISYO1 LIKE '%血圧%' OR m.MEISYO1 LIKE '%服薬%')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)) ORDER BY 1
"@ 40

Q '--- 5. 見本: 石屋 4691 の同じ項目 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 結果, MAX(LTRIM(RTRIM(ISNULL(k.KEKKA_CD,'')))) AS 結果CD,
       MAX(LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))) AS 判定
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND s.UKE_NO_KENSA = 4691
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) IN ('084001','084002','084003','600150','600340','600320','11000') OR m.MEISYO1 LIKE '%血圧%' OR m.MEISYO1 LIKE '%服薬%')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)) ORDER BY 1
"@ 40

Q '--- 6. 判定基準の表の列 (J をどこで決めているかの手がかり) ---' @"
SELECT c.TABLE_NAME AS 表, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型
FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME IN ('T_KIJUN1','T_KIJUN2') ORDER BY c.TABLE_NAME, c.ORDINAL_POSITION
"@ 80

W ''
W '=== 読み方 ==='
W '  2: J の人と石屋で、服薬歴の 結果/結果CD の入り方 (例 はい/1 と はい/空) が違えば、それが原因の候補。'
W '  3: J がくくり (600150) にしか付いていなければ、画面で手入力しているか、くくりの判定基準で付けている。'
W '  4・5: J の人と 4691 で、服薬歴や血圧の項目の持ち方を見くらべる。'
notepad $out
