<#
  総合判定が F(要精検) になった原因を調べる (check54.ps1)
  check54.bat をダブルクリックすると実行され、結果 r_check54.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  総合判定には「各判定のうち最も重いもの」が入る。
  F の人が、どの項目で F を持っているのかを、集計と人ごとの一覧で出す。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check54.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 総合判定 F の原因 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 報告書に印字されるのは 11000 (総合判定(編集))
$FPEOPLE = @"
SELECT s.PK_SEQ FROM T_KENSIN s
JOIN T_KENSA t ON t.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(t.KOMOKU_CD)) = '11000'
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(t.HANTEI_KIGO,''))) = 'F'
"@

Q '--- 1. 総合判定(編集)の分布 ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 総合判定,
       MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 文言, COUNT(*) AS 人数
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND LTRIM(RTRIM(k.KOMOKU_CD)) = '11000'
GROUP BY k.HANTEI_KIGO ORDER BY 3 DESC
"@ 20

Q '--- 2. ★F の人が F を持っている「くくり判定」(600番台) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS くくり判定, COUNT(*) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND s.PK_SEQ IN ($FPEOPLE)
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600%'
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'F'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 3 DESC
"@ 40

Q '--- 3. ★★F の人が F を持っている「個別の検査項目」 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 検査項目, COUNT(*) AS 人数,
       MIN(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 値の例1,
       MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 値の例2
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND s.PK_SEQ IN ($FPEOPLE)
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'F'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '600%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 3 DESC
"@ 60

Q '--- 4. ★人ごとの一覧 (誰が どの項目で F か) ---' @"
SELECT g.KANJI_SIMEI AS 氏名, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 検査項目,
       '[' + LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 30) + ']' AS 結果
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND s.PK_SEQ IN ($FPEOPLE)
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'F'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '600%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%'
ORDER BY g.KANA_SIMEI, k.KOMOKU_CD
"@ 200

Q '--- 5. 参考: F の人で、F を持つ個別項目が1つも無い人 (くくり判定だけ F の人) ---' @"
SELECT g.KANJI_SIMEI AS 氏名, s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND s.PK_SEQ IN ($FPEOPLE)
  AND NOT EXISTS (
    SELECT 1 FROM T_KENSA k2 WHERE k2.PK_SEQ = s.PK_SEQ
      AND LTRIM(RTRIM(ISNULL(k2.HANTEI_KIGO,''))) = 'F'
      AND LTRIM(RTRIM(k2.KOMOKU_CD)) NOT LIKE '600%'
      AND LTRIM(RTRIM(k2.KOMOKU_CD)) NOT LIKE '610%'
      AND LTRIM(RTRIM(k2.KOMOKU_CD)) NOT LIKE '1000%'
      AND LTRIM(RTRIM(k2.KOMOKU_CD)) NOT LIKE '1100%')
ORDER BY g.KANA_SIMEI
"@ 40

W ''
W '=== 読み方 ==='
W '  2 が「どの分野で要精検になったか」(血液一般・脂質代謝・消化器 など)。'
W '  3 が本体です。「どの検査値・所見が F だったか」が人数つきで分かります。'
W '  4 は人ごとの明細です。誰にどう説明するかを決める材料になります。'
W '  5 に名前が出たら、個別項目には F が無いのにくくり判定だけ F の人です。'
W '     所見(胸部・胃部・心電図)の F がくくり判定に出ている場合などが該当します。'
notepad $out
