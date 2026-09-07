<#
  自動判定の漏れを調べる (check40.ps1)
  check40.bat をダブルクリックすると実行され、結果 r_check40.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  1人ずつ手で自動判定を回したので、押し忘れた人がいないかを確かめます。
  あわせて、結果は入っているのに判定だけ空の項目も洗い出します。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check40.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 自動判定の漏れの調査 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★自動判定を回した日時が残っているか (人数) ---' @"
SELECT CASE WHEN LTRIM(RTRIM(ISNULL(s.D_JIDOHANTEI,''))) = '' THEN N'回していない' ELSE N'回した' END AS 区分,
       COUNT(*) AS 人数,
       MIN(s.D_JIDOHANTEI) AS 最初, MAX(s.D_JIDOHANTEI) AS 最後
FROM T_KENSIN s
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
GROUP BY CASE WHEN LTRIM(RTRIM(ISNULL(s.D_JIDOHANTEI,''))) = '' THEN N'回していない' ELSE N'回した' END
"@ 10

Q '--- 2. ★★自動判定を回していない人 (結果はあるのに判定日が空) ---' @"
SELECT x.氏名, x.カナ, x.結果あり, x.判定あり, x.PK_SEQ
FROM (
  SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ, s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
           AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#'))               AS 結果あり,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
           AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '')                   AS 判定あり,
         LTRIM(RTRIM(ISNULL(s.D_JIDOHANTEI,'')))                               AS 判定日
  FROM T_KENSIN s
  LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
WHERE x.結果あり > 0 AND x.判定日 = ''
ORDER BY x.カナ
"@ 100

Q '--- 3. ★1人あたりの判定の数 (少ない人がいないか) ---' @"
SELECT x.判定あり AS 判定が付いた項目数, COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
           AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '') AS 判定あり
  FROM T_KENSIN s
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
GROUP BY x.判定あり
ORDER BY x.判定あり
"@ 60

Q '--- 4. ★判定の数が少ない人 (下から15人) ---' @"
SELECT TOP 15 x.氏名, x.カナ, x.結果あり, x.判定あり, x.判定日, x.PK_SEQ
FROM (
  SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ, s.PK_SEQ,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
           AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#'))     AS 結果あり,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
           AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '')         AS 判定あり,
         LTRIM(RTRIM(ISNULL(s.D_JIDOHANTEI,'')))                     AS 判定日
  FROM T_KENSIN s
  LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
  WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
WHERE x.結果あり > 0
ORDER BY x.判定あり, x.カナ
"@ 20

Q '--- 5. ★★結果はあるのに判定が空の項目 (多い順) ---' @"
SELECT TOP 40 LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 結果あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = '' THEN 1 ELSE 0 END) AS 判定なし
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = '' THEN 1 ELSE 0 END) > 0
ORDER BY 4 DESC, 1
"@ 45

Q '--- 6. ★尿検査 (069206 蛋白 / 069207 糖 / 069211 潜血) の入り方 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 枠の数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#') THEN 1 ELSE 0 END) AS 結果あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''     THEN 1 ELSE 0 END) AS 判定あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('069206','069207','069211')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1
"@ 10

Q '--- 7. ★尿検査の結果が入っていない人 ---' @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ, s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
               AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#'))
  AND NOT EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
                   AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('069206','069207','069211')
                   AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#'))
ORDER BY g.KANA_SIMEI
"@ 40

Q '--- 8. 総合判定 (自動判定が最後に出すもの) が入っている人数 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 枠の数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#') THEN 1 ELSE 0 END) AS 結果あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''     THEN 1 ELSE 0 END) AS 判定あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) = '11000' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600%')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 4 DESC, 1
"@ 45

W ''
W '=== 読み方 ==='
W '  1 で「回していない」が 0 なら、押し忘れはありません。'
W '     10人は結果が来ていない人なので、そこは回っていなくて当然です。'
W ''
W '  2 に名前が並べば、その人が押し忘れです。その人だけ健診ナビで開いて回してください。'
W ''
W '  3 で、ほとんどの人が同じくらいの数(50〜70)に固まっていれば正常です。'
W '     ひとりだけ 6〜7 のような人がいたら、その人は回っていません。'
W ''
W '  5 は「結果はあるが判定が空」の項目です。ここに出るのは次のどれかです。'
W '     ・そもそも判定を出さない項目 (フィルムNo・問診など) → 正常'
W '     ・基準値が登録されていない項目 → 健診ナビのマスタの問題'
W '     どの項目が並ぶかを見て、直すべきものがあるか判断します。'
W ''
W '  7 に出るのは尿の結果が無い人です。'
W '     東振協のファイルでも空だったのは 阿久津 有加 さん1人だけでした。'
W '     それ以外の名前が出たら、取込で落ちています。'

notepad $out
