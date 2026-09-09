<#
  「血圧で要治療(G)」になっている理由を調べる (check50.ps1)
  check50.bat をダブルクリックすると実行され、結果 r_check50.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  秋葉達也さんの結果報告書が
    ・総合判定 G 要治療
    ・医師指示事項に【血圧】重症の高血圧を認めます。治療の必要があります。
  なのに、血圧の値は 122/71 で正常。
  どの項目が G を持っていて、その文章がどこから来ているかを見る。

  秋葉さんは以前テスト用に使った人なので、そのときの残りが疑わしい。
#>
[CmdletBinding()]
param(
    [string]$Ymd  = '2026/08/21',
    [string]$Name = '秋葉達也',     # 空白なしで
    [string]$Cmp  = '小野千絵'      # くらべる用 (血圧が正常でAの人)
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check50.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
# 氏名の全角/半角スペースを無視して照合する
$byName = "REPLACE(REPLACE(g.KANJI_SIMEI, N'　', ''), ' ', '') LIKE N'%$Name%'"
$byCmp  = "REPLACE(REPLACE(g.KANJI_SIMEI, N'　', ''), ' ', '') LIKE N'%$Cmp%'"

"=== 血圧 G の原因調査 ($Name / $Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★この人の受診の行 (T_KENSIN)。テストのときの残りが無いか ---" @"
SELECT s.PK_SEQ, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.F_TORIKESI AS 取消,
       '[' + LTRIM(RTRIM(ISNULL(s.SEQ1,''))) + ']' AS SEQ1, s.KOJIN_ID,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS 検査行数
FROM T_KENSIN s
JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE $byName
ORDER BY s.D_KENSIN DESC
"@ 20

Q "--- 2. ★★$Ymd のこの人で、判定が A/B 以外の項目すべて ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 60) + ']' AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HL,''))) + ']' AS HL
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
JOIN T_KOJIN1 g      ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND $byName
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('', 'A', 'B')
ORDER BY k.KOMOKU_CD
"@ 60

Q "--- 3. ★血圧まわりの項目 (値・判定・HL) ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']' AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(k.HL,''))) + ']' AS HL,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
JOIN T_KOJIN1 g      ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND $byName
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) IN ('001253','001254','001255','001256','600150','084001')
       OR m.MEISYO1 LIKE N'%血圧%')
ORDER BY k.KOMOKU_CD
"@ 30

Q "--- 4. ★★総合判定・医師指示事項の枠 (10000/11000番台) の中身と判定 ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 70) + ']' AS 結果
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
JOIN T_KOJIN1 g      ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND $byName
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1100%')
ORDER BY k.KOMOKU_CD
"@ 40

Q "--- 5. ★「重症の高血圧」という文章が $Ymd の誰に入っているか ---" @"
SELECT g.KANJI_SIMEI AS 氏名, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + ISNULL(bp1.KEKKA,'') + '/' + ISNULL(bp2.KEKKA,'') + ']' AS 血圧1回目
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
JOIN T_KOJIN1 g      ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KENSA bp1 ON bp1.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(bp1.KOMOKU_CD)) = '001253'
LEFT JOIN T_KENSA bp2 ON bp2.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(bp2.KOMOKU_CD)) = '001254'
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND k.KEKKA LIKE N'%重症の高血圧%'
ORDER BY g.KANA_SIMEI
"@ 30

Q "--- 6. その文章のマスタ (T_SYOKEN2)。文章自体に判定記号が付いているか ---" @"
SELECT LTRIM(RTRIM(y.SYOKEN_CD)) AS 所見リスト, LTRIM(RTRIM(y.KEKKA_CD)) AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(y.HANTEI_KIGO,''))) + ']' AS 判定,
       LEFT(y.SYOKEN, 70) AS 文章
FROM T_SYOKEN2 y
WHERE y.SYOKEN LIKE N'%重症の高血圧%'
"@ 20

Q "--- 7. くらべる用: $Cmp の総合判定・医師指示事項・血圧判定 ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 60) + ']' AS 結果
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
JOIN T_KOJIN1 g      ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND $byCmp
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1100%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) IN ('001253','001254','600150','084001'))
ORDER BY k.KOMOKU_CD
"@ 40

Q "--- 8. $Ymd で総合判定が G の人の一覧 (他にもいないか) ---" @"
SELECT g.KANJI_SIMEI AS 氏名, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 総合判定
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) = '10000'
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('G','J','F','E')
ORDER BY k.HANTEI_KIGO, g.KANA_SIMEI
"@ 40

W ''
W '=== 読み方 ==='
W '  1 に 8/21 以外の受診の行 (テストのときのもの) が残っていないか。'
W '  2 が本体です。G を持っている項目がどれかが分かります。'
W '     血圧の値そのもの (001253) が G なのか、'
W '     医師指示事項の枠 (11001〜) に入っている文章が G を持っているのか。'
W '  4 と 6 で、「重症の高血圧」の文章に判定記号 G が付いているなら、'
W '     その文章が枠に残っているだけで総合判定が G に引きずられている、ということです。'
W '  5 で同じ文章が他の人にも入っていないかを見ます。'
W '  8 で G の人が他にもいないかを見ます。'

notepad $out
