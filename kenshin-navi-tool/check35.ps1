<#
  心電図の判定が出ない理由を調べる (check35.ps1)
  check35.bat をダブルクリックすると実行され、結果 r_check35.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  調べること
    1. 心電図の所見マスタ(ZK011)に、そもそも判定記号が入っているか
    2. 8/21 の心電図の行 (067112A/B/C) に何が入っているか
    3. 心電図の判定を入れる専用の項目が別にあるか
    4. 判定がちゃんと出ている過去の日と、8/21 で何が違うか
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'    # 調べたい受診日
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check35.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 心電図の判定の調査 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★心電図の所見マスタ ZK011 に判定記号が入っているか ---' @"
SELECT LTRIM(RTRIM(KEKKA_CD)) AS 結果CD,
       SYOKEN AS 所見,
       '[' + LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) + ']' AS 判定記号
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK011'
ORDER BY LEN(LTRIM(RTRIM(KEKKA_CD))), LTRIM(RTRIM(KEKKA_CD))
"@ 120

Q '--- 2. ★8/21 の心電図の行に何が入っているか (上位40人) ---' @"
SELECT g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'       AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']'    AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定
FROM T_KENSA k
JOIN T_KENSIN s   ON s.PK_SEQ  = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '067112%'
ORDER BY g.KANA_SIMEI, k.KOMOKU_CD
"@ 40

Q '--- 3. ★8/21 の心電図: 判定が入っている人 / 入っていない人の数 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,'')))       NOT IN ('','#') THEN 1 ELSE 0 END) AS 結果あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''           THEN 1 ELSE 0 END) AS 判定あり,
       COUNT(*) AS 枠の数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '067112%'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1
"@ 20

Q '--- 4. 心電図まわりの項目マスタ (判定専用の項目が別にあるか) ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(m.SYOKEN_CD,''))) + ']' AS 所見リスト
FROM T_KOMOKU m
WHERE LTRIM(RTRIM(m.KOMOKU_CD)) LIKE '0671%'
   OR m.MEISYO1 LIKE N'%心電図%'
ORDER BY 1
"@ 80

Q '--- 5. ★判定が出ている日を探す (心電図の判定が入っている受診日 上位20) ---' @"
SELECT TOP 20 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       COUNT(*) AS 心電図の行数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '067112%'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY s.D_KENSIN
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) > 0
ORDER BY s.D_KENSIN DESC
"@ 25

Q '--- 6. その「判定が出ている日」の中身を10行だけ見る ---' @"
SELECT TOP 10 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'       AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']'    AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '067112%'
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
ORDER BY s.D_KENSIN DESC, k.PK_SEQ
"@ 15

Q '--- 7. くらべる: 同じ8/21で判定が出ている項目 (心電図以外) ---' @"
SELECT TOP 25 LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) > 0
ORDER BY 4 DESC
"@ 30

W ''
W '=== 読み方 ==='
W '  1 の「判定記号」が全部 [] なら、それが原因です。'
W '     健診ナビの所見マスタ(ZK011)に判定記号が登録されていないので、'
W '     所見を入れても判定の出しようがありません。'
W '     → 健診ナビのマスタ保守で ZK011 の各所見に判定記号を入れてもらう話になります。'
W ''
W '  2 で「結果」に所見が入っていて「判定」だけ [] なら、'
W '     取込は成功していて、判定だけが空という状態です。'
W ''
W '  5 に日付が並べば、過去には判定が出ていたということです。'
W '     6 でその日の中身を見て、8/21 と何が違うかを比べます。'
W '     5 が空なら、この健診ナビでは心電図に判定を出す運用自体が無い可能性が高いです。'
W ''
W '  7 に項目が並べば、自動判定そのものは動いています。'
W '     (心電図だけ出ない、ということになります)'

notepad $out
