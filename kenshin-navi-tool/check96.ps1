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

W ''
W '=== 読み方 ==='
W '  1: 4673 の 1回目が (+) で、セット OPJ010 か 1回目 の判定に F (要精検) が付いていれば問題なし。判定が空なら自動判定がまだ。'
W '  2: 2 本出した人でどの行に判定が付くか (セットに付くのか、1回目・2回目それぞれに付くのか) の確認用。'
W '  3: 総合判定が空の人は自動判定がまだ走っていない。'
notepad $out
