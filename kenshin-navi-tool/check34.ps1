<#
  取込の結果を数える (check34.ps1)
  check34.bat をダブルクリックすると実行され、結果 r_check34.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  何人に・何件のデータが入ったかを数えます。
  ファイルにあった88人と合っているかを確かめるためのものです。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check34.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 取込結果の確認 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★人数: 値が入っている人 / 入っていない人 ---' @"
SELECT CASE WHEN x.件数 > 0 THEN N'値あり' ELSE N'値なし' END AS 区分,
       COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k
           WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('', '#')) AS 件数
  FROM T_KENSIN s
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
GROUP BY CASE WHEN x.件数 > 0 THEN N'値あり' ELSE N'値なし' END
"@ 10

Q '--- 2. ★合計件数 ---' @"
SELECT COUNT(*) AS 値の総数,
       COUNT(DISTINCT k.PK_SEQ) AS 人数,
       COUNT(DISTINCT k.KOMOKU_CD) AS 項目数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('', '#')
"@ 10

Q '--- 3. 項目ごとの人数 (多い順) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('', '#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY COUNT(*) DESC, 1
"@ 200

Q '--- 4. 値が1件も入っていない人 (未受診のはず) ---' @"
SELECT g.KANJI_SIMEI AS 漢字氏名, g.KANA_SIMEI AS カナ氏名, s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND NOT EXISTS (SELECT 1 FROM T_KENSA k
                   WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('', '#'))
ORDER BY g.KANA_SIMEI
"@ 40

Q '--- 5. 1人あたりの件数の散らばり (極端に少ない人がいないか) ---' @"
SELECT x.件数 AS 件数, COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k
           WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('', '#')) AS 件数
  FROM T_KENSIN s
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
WHERE x.件数 > 0
GROUP BY x.件数
ORDER BY x.件数
"@ 60

W ''
W '=== 読み方 ==='
W '  1 で「値あり 88 / 値なし 10」なら狙いどおりです。'
W '     ファイルにあったのが88人、予約だけで結果が来ていない人が10人。'
W '  2 の値の総数が 5,700 件前後なら、取込は全部通っています。'
W '  3 の「判定あり」は自動判定が付けた記号の数です。'
W '     自動判定をまだ実行していなければ、取込時に付くものだけ(所見・問診)になります。'
W '  4 に並ぶのが未受診の人です。10人のはずです。'
W '  5 で極端に件数の少ない人がいれば、その人だけ取りこぼしている可能性があります。'

notepad $out
