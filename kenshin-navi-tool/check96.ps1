<#
  石屋: 便潜血を 1 本しか出していない 10 人の結果と判定を見る (check96.ps1)
  check96.bat をダブルクリックすると実行され、結果 r_check96.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check96.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 便潜血 1 本だけの 10 人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★10 人の 便潜血 (OPJ010 セット / 069245 1回目 / 069246 2回目) の値と判定 ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       ISNULL(k.KEKKA,'') AS 結果, ISNULL(k.KEKKA_CD,'') AS 結果CD, ISNULL(k.HANTEI_KIGO,'') AS 判定, ISNULL(k.F_HANTEI,'') AS F_HANTEI
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA IN (4655,4666,4673,4675,4699,4707,4708,4710,4735,4769)
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('OPJ010','069245','069246')
ORDER BY s.UKE_NO_KENSA, k.KOMOKU_CD
"@ 60

Q '--- 2. 参考: 2 本とも出した人の 便潜血 の判定の付き方 (セット OPJ010 / 1回目 / 2回目 ごとの判定の分布) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, ISNULL(k.KEKKA,'') AS 結果, ISNULL(k.HANTEI_KIGO,'') AS 判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('OPJ010','069245','069246')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), ISNULL(k.KEKKA,''), ISNULL(k.HANTEI_KIGO,'')
ORDER BY 1, 2, 3
"@ 40

Q '--- 3. 10 人の 総合判定 (MT00A) と 自動判定が済んでいるか (BMI があるか) ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       ISNULL((SELECT TOP 1 ISNULL(k.HANTEI_KIGO,'') FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = 'MT00A'),'') AS 総合判定,
       ISNULL((SELECT TOP 1 ISNULL(k.KEKKA,'') FROM T_KENSA k JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)) WHERE k.PK_SEQ = s.PK_SEQ AND m.MEISYO1 LIKE N'%BMI%'),'') AS BMI
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA IN (4655,4666,4673,4675,4699,4707,4708,4710,4735,4769)
ORDER BY s.UKE_NO_KENSA
"@ 20

Q '--- 4. ★健診ナビにある 便潜血 の項目 (1 日法・1 本法 のような別項目があるか) と 石屋の枠数 ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, LTRIM(RTRIM(ISNULL(m.SYOKEN_CD,''))) AS 選択肢リスト,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
         WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
           AND LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))) AS 石屋の枠,
       (SELECT COUNT(*) FROM T_KENSA k WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))) AS 全体の枠
FROM T_KOMOKU m
WHERE m.MEISYO1 LIKE N'%便%' OR LTRIM(RTRIM(m.KOMOKU_CD)) LIKE 'OPJ%' OR LTRIM(RTRIM(m.KOMOKU_CD)) IN ('069245','069246')
ORDER BY 1
"@ 40

Q '--- 5. ★便潜血 2回目 (069246) の選択肢 (「未提出」「未実施」のような選択肢があるか) ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, y.KEKKA_CD AS 結果CD, y.SYOKEN AS 選択肢, ISNULL(y.HANTEI_KIGO,'') AS 判定
FROM T_KOMOKU m JOIN T_SYOKEN2 y ON LTRIM(RTRIM(y.SYOKEN_CD)) = LTRIM(RTRIM(m.SYOKEN_CD))
WHERE LTRIM(RTRIM(m.KOMOKU_CD)) IN ('069245','069246')
ORDER BY 1, y.KEKKA_CD
"@ 40

Q '--- 6. 参考: 過去 (今年) に 便潜血 2回目 だけ空のまま 自動判定済みの人が他にいるか (セットの判定がどうなっているか) ---' @"
SELECT TOP 20 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号,
       ISNULL(k1.KEKKA,'') AS 一回目, ISNULL(k1.HANTEI_KIGO,'') AS 判定1, ISNULL(k2.KEKKA,'') AS 二回目, ISNULL(k2.HANTEI_KIGO,'') AS 判定2,
       ISNULL(ks.KEKKA,'') AS セット結果, ISNULL(ks.HANTEI_KIGO,'') AS セット判定,
       ISNULL((SELECT TOP 1 ISNULL(x.HANTEI_KIGO,'') FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = 'MT00A'),'') AS 総合判定
FROM T_KENSIN s
JOIN T_KENSA k1 ON k1.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k1.KOMOKU_CD)) = '069245'
JOIN T_KENSA k2 ON k2.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k2.KOMOKU_CD)) = '069246'
LEFT JOIN T_KENSA ks ON ks.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ks.KOMOKU_CD)) = 'OPJ010'
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND s.D_KENSIN < '2026/09/01'
  AND LTRIM(RTRIM(ISNULL(k1.KEKKA,''))) <> '' AND LTRIM(RTRIM(ISNULL(k2.KEKKA,''))) = ''
ORDER BY s.D_KENSIN DESC
"@ 25

W ''
W '=== 読み方 ==='
W '  4: 「便潜血1日法」のような別の項目があれば、1 本として入れ直す選択肢がある (ただしコースの項目を変えることになる)。'
W '  5: 2回目 の選択肢に「未提出」などがあれば、2回目 にそれを入れて 2 日法のまま 1 本で判定させられる。'
W '  6: 過去に同じ状態 (1回目だけ) の人が自動判定でどう判定されたかの実例。'
W '  1: 4673 の 1回目が (+) で、セット OPJ010 か 1回目 の判定に F (要精検) が付いていれば問題なし。判定が空なら自動判定がまだ。'
W '  2: 2 本出した人でどの行に判定が付くか (セットに付くのか、1回目・2回目それぞれに付くのか) の確認用。'
W '  3: 総合判定が空の人は自動判定がまだ走っていない。'
notepad $out
