<#
  医師指示事項がDBのどの欄かを突き止める (check44.ps1)
  check44.bat をダブルクリックすると実行され、結果 r_check44.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  結果報告書の「医師指示事項」に、判定に応じた文章を入れたい。
  そのためには、その欄がどの項目コードなのかを確定させる必要があります。
  推測で書き込むと、医師の指示欄に誤った文章が入ります。

  結果報告書で中身が見えている人 (既定は山本 稔さん) の全項目を出し、
  画面の文章と突き合わせて特定します。
#>
[CmdletBinding()]
param(
    [string]$Ymd   = '2026/08/21',
    [string]$PkSeq = '2005092'      # 山本 稔さん (結果報告書を開いていた人)
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check44.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 医師指示事項の欄をさがす $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★★この人 (PK_SEQ=$PkSeq) の、文章が入っている項目すべて ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       k.KEKKA AS 中身
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.PK_SEQ = $PkSeq
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND LEN(LTRIM(RTRIM(k.KEKKA))) >= 10
ORDER BY 1
"@ 60

Q "--- 2. ★★この人の 総合判定まわり (1000x / 1100x / 11000) ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'       AS 中身,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.PK_SEQ = $PkSeq
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1100%')
ORDER BY 1
"@ 40

Q '--- 3. ★★お手本の文章が、どの項目コードに入っているか (過去データ全部) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, COUNT(*) AS 件数,
       MAX(LEFT(k.KEKKA, 40)) AS 文章の例
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.KEKKA LIKE N'%胸部X線検査にて異常所見を認めます%'
   OR k.KEKKA LIKE N'%胸部Ｘ線検査にて異常所見を認めます%'
   OR k.KEKKA LIKE N'%心電図異常があります%'
   OR k.KEKKA LIKE N'%心電図異常については%'
   OR k.KEKKA LIKE N'%胃部レントゲン検査にて異常所見を認めます%'
   OR k.KEKKA LIKE N'%胃部内視鏡検査にて異常所見を認めます%'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 3 DESC
"@ 30

Q '--- 4. ★お手本の文章の実物 (どう登録されているか。番号との対応を見る) ---' @"
SELECT DISTINCT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       k.KEKKA AS 文章
FROM T_KENSA k
WHERE k.KEKKA LIKE N'%胸部X線検査にて異常所見を認めます%'
   OR k.KEKKA LIKE N'%胸部Ｘ線検査にて異常所見を認めます%'
   OR k.KEKKA LIKE N'%心電図異常があります%'
   OR k.KEKKA LIKE N'%心電図異常については%'
   OR k.KEKKA LIKE N'%胃部レントゲン検査にて異常所見を認めます%'
   OR k.KEKKA LIKE N'%胃部内視鏡検査にて異常所見を認めます%'
ORDER BY 1, 2
"@ 40

Q '--- 5. ★コメントのマスタらしき表を探す (308 や 58 の番号が引けるか) ---' @"
SELECT c.TABLE_NAME AS テーブル, COUNT(*) AS 列数
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME LIKE '%COMMENT%' OR c.TABLE_NAME LIKE '%SIJI%'
   OR c.TABLE_NAME LIKE '%SOUGOU%'  OR c.TABLE_NAME LIKE '%SOGO%'
   OR c.TABLE_NAME LIKE '%BUNSYO%'  OR c.TABLE_NAME LIKE '%MONKU%'
GROUP BY c.TABLE_NAME
ORDER BY 1
"@ 40

Q "--- 6. ★$Ymd の人ごとの総合判定の使用状況 (何枠まで埋まっているか) ---" @"
SELECT x.使用枠数 AS 埋まっている枠の数, COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ
           AND LTRIM(RTRIM(c.KOMOKU_CD)) LIKE '1100[1-9]'
           AND LTRIM(RTRIM(ISNULL(c.KEKKA,''))) NOT IN ('','#')) AS 使用枠数
  FROM T_KENSIN s
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
GROUP BY x.使用枠数
ORDER BY x.使用枠数
"@ 20

Q "--- 7. ★$Ymd に用意されている総合判定の枠 (11001〜) ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 枠の数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#') THEN 1 ELSE 0 END) AS 中身あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1000%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '1100%')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1
"@ 40

W ''
W '=== 読み方 ==='
W '  1 と 2 を、結果報告書の画面と見くらべてください。'
W '     医師指示事項に出ている文章と同じものが入っている項目コードが答えです。'
W '     山本 稔さんなら「【視力】低下が見られます。眼科にて視力矯正…」です。'
W ''
W '  3 で、お手本の6つの文章が実際にどの項目に入っているかが分かります。'
W '     ここが 11001〜11005 に集まっていれば、そこが医師指示事項です。'
W ''
W '  4 は文章の実物です。308番などの番号が結果CDに入っているかを見ます。'
W '     入っていれば、番号を指定するだけで文章を呼び出せます。'
W ''
W '  6 と 7 で、1人あたり何枠まで使えるかが分かります。'
W '     追加する文章は「空いている枠」に入れます。既にある文章は消しません。'

notepad $out
