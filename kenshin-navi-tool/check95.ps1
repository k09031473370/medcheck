<#
  採血時間 (001282 食後時間/空腹) が血糖・中性脂肪の判定に効いているか見る (check95.ps1)
  check95.bat をダブルクリックすると実行され、結果 r_check95.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  背景 (2026/10/16): 石屋 9/9・9/10 の 141 人は 001282 が空のまま (123 列版の対応表が ＊空腹 列を取り込んでいない)。
  9/3 の 58 人は入っている。入れるかどうかを決めるために、判定への影響を見る。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check95.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 採血時間 001282 と血糖・中性脂肪の判定 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"

Q '--- 1. 001282 (食後時間/空腹) の選択肢 (T_KOMOKU → T_SYOKEN2) ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, LTRIM(RTRIM(m.SYOKEN_CD)) AS リスト,
       y.KEKKA_CD AS 結果CD, y.SYOKEN AS 選択肢, y.HANTEI_KIGO AS 判定
FROM T_KOMOKU m LEFT JOIN T_SYOKEN2 y ON LTRIM(RTRIM(y.SYOKEN_CD)) = LTRIM(RTRIM(m.SYOKEN_CD))
WHERE LTRIM(RTRIM(m.KOMOKU_CD)) = '001282'
ORDER BY y.KEKKA_CD
"@ 30

Q '--- 2. 血糖・中性脂肪に関係する項目 (名前に 血糖 / 中性脂肪 / ﾄﾘｸﾞﾘ を含むもの) ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       (SELECT COUNT(*) FROM T_KIJUN1 j WHERE j.KIJUN_CD = '1' AND LTRIM(RTRIM(j.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))) AS 基準値1の行数,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
         WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))) AS 石屋の枠,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
         WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD)) AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS 石屋の値あり
FROM T_KOMOKU m
WHERE m.MEISYO1 LIKE N'%血糖%' OR m.MEISYO1 LIKE N'%中性脂肪%' OR m.MEISYO1 LIKE N'%ﾄﾘｸﾞﾘ%' OR m.MEISYO1 LIKE N'%TG%'
   OR LTRIM(RTRIM(m.KOMOKU_CD)) IN ('068681','068682','H30-0011','R06-0002','068702','001282')
ORDER BY 1
"@ 60

Q '--- 3. ★血糖・中性脂肪の基準値1 (空腹時と随時で分かれているか) ---' @"
SELECT LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       CASE j.SEIBETU WHEN '0' THEN '共通' WHEN '1' THEN '男' WHEN '2' THEN '女' ELSE j.SEIBETU END AS 性別,
       j.HANTEI_KIGO AS 判定, ISNULL(j.SAISYO,'') + ' 〜 ' + ISNULL(j.SAIDAI,'') AS 範囲, ISNULL(j.HL,'') AS HL, j.RENBAN AS 連番
FROM T_KIJUN1 j JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE j.KIJUN_CD = '1'
  AND (m.MEISYO1 LIKE N'%血糖%' OR m.MEISYO1 LIKE N'%中性脂肪%' OR LTRIM(RTRIM(j.KOMOKU_CD)) IN ('068682','H30-0011','R06-0002','068702'))
ORDER BY 1, 3, 7
"@ 120

Q '--- 4. ★石屋: 受診日ごと 001282 の入り方 × 血糖 068682 の判定 (自動判定済みの人で、001282 の有無で判定の付き方が違うか) ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日,
       CASE WHEN p.空腹 = '' THEN '(001282 空)' ELSE p.空腹 END AS 採血時間,
       COUNT(*) AS 人数,
       SUM(CASE WHEN p.血糖値 <> '' THEN 1 ELSE 0 END) AS 血糖値あり,
       SUM(CASE WHEN p.血糖判定 <> '' THEN 1 ELSE 0 END) AS 血糖判定あり,
       SUM(CASE WHEN p.血糖判定 = 'A' THEN 1 ELSE 0 END) AS A, SUM(CASE WHEN p.血糖判定 = 'B' THEN 1 ELSE 0 END) AS B,
       SUM(CASE WHEN p.血糖判定 = 'C' THEN 1 ELSE 0 END) AS C, SUM(CASE WHEN p.血糖判定 = 'D' THEN 1 ELSE 0 END) AS D,
       SUM(CASE WHEN p.血糖判定 IN ('E','F','G','J') THEN 1 ELSE 0 END) AS EFGJ,
       SUM(CASE WHEN p.TG判定 <> '' THEN 1 ELSE 0 END) AS 中性脂肪判定あり,
       SUM(CASE WHEN p.総合 <> '' THEN 1 ELSE 0 END) AS 総合判定あり
FROM (
  SELECT s.PK_SEQ, s.D_KENSIN,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.KEKKA,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '001282'),'') AS 空腹,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.KEKKA,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068682'),'') AS 血糖値,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068682'),'') AS 血糖判定,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = 'R06-0002'),'') AS TG判定,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = 'MT00A'),'') AS 総合
  FROM T_KENSIN s WHERE $ISHIYA
) p
GROUP BY p.D_KENSIN, p.空腹
ORDER BY 1, 2
"@ 30

Q '--- 5. 参考: 血糖 068682 の値の範囲ごとに付いている判定 (石屋 3 日分、自動判定済みの人) ---' @"
SELECT CASE WHEN p.空腹 = '' THEN '(001282 空)' ELSE p.空腹 END AS 採血時間,
       CASE WHEN p.v < 100 THEN '1: 〜99' WHEN p.v < 110 THEN '2: 100〜109' WHEN p.v < 126 THEN '3: 110〜125' WHEN p.v < 140 THEN '4: 126〜139' ELSE '5: 140〜' END AS 血糖の範囲,
       p.血糖判定 AS 判定, COUNT(*) AS 人数
FROM (
  SELECT s.PK_SEQ,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.KEKKA,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '001282'),'') AS 空腹,
    TRY_CONVERT(float, (SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.KEKKA,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068682')) AS v,
    ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068682'),'') AS 血糖判定
  FROM T_KENSIN s WHERE $ISHIYA
) p
WHERE p.v IS NOT NULL
GROUP BY p.空腹, CASE WHEN p.v < 100 THEN '1: 〜99' WHEN p.v < 110 THEN '2: 100〜109' WHEN p.v < 126 THEN '3: 110〜125' WHEN p.v < 140 THEN '4: 126〜139' ELSE '5: 140〜' END, p.血糖判定
ORDER BY 1, 2, 3
"@ 60

W ''
W '=== 読み方 ==='
W '  1: 001282 の選択肢。9/9・9/10 の 123 列ファイルの「＊空腹=1」がどの選択肢に当たるかを決めるのに使う。'
W '  3: 血糖 (068682 未判別) の基準値が 1 組しか無ければ、採血時間は判定に効いていない。空腹時用と随時用の 2 組があれば効いている。'
W '  4: 自動判定済みの人 (総合判定あり) について、001282 が空の人でも血糖の判定が付いていれば、少なくとも「判定が付かない」ことはない。'
W '     付いていなければ、採血時間が無いと血糖が判定されない = 9/9・9/10 に 001282 を入れる必要がある。'
W '  5: 同じ血糖値の範囲で、001282 ありとなしで判定の文字が違えば、採血時間で基準が切り替わっている。'
notepad $out
