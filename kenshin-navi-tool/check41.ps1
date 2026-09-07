<#
  本当の判定漏れを見つける (check41.ps1)
  check41.bat をダブルクリックすると実行され、結果 r_check41.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  「結果はあるのに判定が空」の項目には2種類あります。
    ・全員が空          → もともと判定を出さない項目 (問診・フィルムNo など)。正常
    ・一部の人だけ空    → ★本当の漏れ。これを探します

  尿蛋白が 87人中84人にしか判定が付いていない、という状態を追うために作りました。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check41.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 本当の判定漏れの調査 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★★一部の人だけ判定が空の項目 (= 本当の漏れ) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 結果あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) =  '' THEN 1 ELSE 0 END) AS 判定なし
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) > 0
   AND SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) =  '' THEN 1 ELSE 0 END) > 0
ORDER BY 5 DESC, 1
"@ 50

Q '--- 2. ★★その人たちの名前と値 (上の項目すべて) ---' @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'    AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = ''
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN (
      SELECT LTRIM(RTRIM(k2.KOMOKU_CD))
      FROM T_KENSA k2 JOIN T_KENSIN s2 ON s2.PK_SEQ = k2.PK_SEQ
      WHERE s2.D_KENSIN = '$Ymd' AND s2.F_TORIKESI = 0
        AND LTRIM(RTRIM(ISNULL(k2.KEKKA,''))) NOT IN ('','#')
      GROUP BY LTRIM(RTRIM(k2.KOMOKU_CD))
      HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k2.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) > 0
         AND SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k2.HANTEI_KIGO,''))) =  '' THEN 1 ELSE 0 END) > 0)
ORDER BY k.KOMOKU_CD, g.KANA_SIMEI
"@ 80

Q '--- 3. くらべる: 尿蛋白(069206) の値ごとの判定の付き方 ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'    AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       COUNT(*) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069206'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY k.KEKKA, k.KEKKA_CD, k.HANTEI_KIGO
ORDER BY 4 DESC
"@ 20

Q '--- 4. 尿蛋白の基準値マスタ (どう判定する設定か) ---' @"
SELECT * FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '069206'
"@ 30

Q '--- 5. 尿糖(069207) の基準値マスタ (くらべる用。こちらは全員に付いている) ---' @"
SELECT * FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '069207'
"@ 30

W ''
W '=== 読み方 ==='
W '  1 に出るのが「一部の人だけ判定が空」の項目です。'
W '     ここに出たものだけが、本当に直すべき漏れです。'
W '     問診やフィルムNoのように全員が空の項目は、ここには出ません。'
W ''
W '  2 でその人の名前と値が分かります。'
W ''
W '  3 で、どの値のときに判定が付いていないかが分かります。'
W '     特定の値 ((+) など) だけ付いていないなら、基準値マスタにその値の'
W '     設定が無い、ということです。'
W ''
W '  4 と 5 を見くらべれば、尿蛋白のマスタに何が足りないかが見えます。'

notepad $out
