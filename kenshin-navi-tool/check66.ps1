<#
  受付番号の入り方を見る (check66.ps1)
  check66.bat をダブルクリックすると実行され、結果 r_check66.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  受付番号は2か所にある
    ① T_KENSIN.UKE_NO_KENSA  … 予約に振っておく番号 (未受付でも入る)
    ② T_KANJA_G.KEN_YMD + KEN_NO … 当日受付した実績
  結果取込は ①②どちらでも人を特定できる (form_import.ps1 の Resolve-PkSeq)。

  ここで見たいこと
    ・普段どちらに番号が入っているか
    ・番号の振り方 (日ごとに1から / コース別 / 時間順 など)
    ・マイクロン(0000000475)の予約が取り込まれているか、番号が空か
#>
[CmdletBinding()]
param(
    [string]$Dantai = '0000000475',    # マイクロンメモリジャパン
    [string]$Ymd    = '2026/08/21'     # 比べる用の過去の健診日 (城西)
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check66.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受付番号の入り方 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KENSIN の受付番号まわりの列 ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KENSIN')
  AND (c.name LIKE '%UKE%' OR c.name LIKE '%KEN_NO%' OR c.name LIKE '%NO%')
ORDER BY c.column_id
"@ 60

Q '--- 2. T_KANJA_G の列 ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KANJA_G') ORDER BY c.column_id
"@ 60

Q '--- 3. ★直近の健診日ごとに、受付番号が入っているか ---' @"
SELECT TOP 25 s.D_KENSIN AS 受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) <> '' THEN 1 ELSE 0 END) AS 予約番号あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.JUSIN_KEN_NO,''))) <> '' THEN 1 ELSE 0 END) AS 受診券番号あり,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) THEN 1 ELSE 0 END) AS 受付済み
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01'
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN DESC
"@ 40

Q "--- 4. ★番号の振り方 ($Ymd の実物。先頭30人) ---" @"
SELECT TOP 30 s.D_KENSIN AS 受診日, LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       '[' + LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) + ']' AS 予約番号,
       '[' + LTRIM(RTRIM(ISNULL(s.JUSIN_KEN_NO,''))) + ']' AS 受診券番号,
       (SELECT TOP 1 '[' + LTRIM(RTRIM(ISNULL(q.KEN_NO,''))) + ']' FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付番号,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
ORDER BY s.UKE_NO_KENSA, s.PK_SEQ
"@ 40

Q '--- 5. ★受付番号が入っている日の例 (振り方を見る) ---' @"
SELECT TOP 40 s.D_KENSIN AS 受診日, LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD,
       LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) AS 予約番号,
       (SELECT TOP 1 LTRIM(RTRIM(ISNULL(q.KEN_NO,''))) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付番号,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) <> ''
  AND s.D_KENSIN >= '2026/06/01'
ORDER BY s.D_KENSIN DESC, s.UKE_NO_KENSA
"@ 50

Q "--- 6. ★マイクロン ($Dantai) の予約が入っているか ---" @"
SELECT s.D_KENSIN AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD, COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) <> '' THEN 1 ELSE 0 END) AS 予約番号あり
FROM T_KENSIN s
WHERE s.DANTAI_CD1 = '$Dantai' AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN, s.COURSE_CD ORDER BY s.D_KENSIN, s.COURSE_CD
"@ 40

Q '--- 7. 予約番号は日ごとに1からか、通し番号か (最小・最大) ---' @"
SELECT TOP 20 s.D_KENSIN AS 受診日, COUNT(*) AS 人数,
       MIN(LEN(LTRIM(RTRIM(s.UKE_NO_KENSA)))) AS 桁の最小, MAX(LEN(LTRIM(RTRIM(s.UKE_NO_KENSA)))) AS 桁の最大,
       MIN(LTRIM(RTRIM(s.UKE_NO_KENSA))) AS 番号の最小, MAX(LTRIM(RTRIM(s.UKE_NO_KENSA))) AS 番号の最大
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) <> '' AND s.D_KENSIN >= '2026/01/01'
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN DESC
"@ 30

W ''
W '=== 読み方 ==='
W '  3 で、普段 予約番号(UKE_NO_KENSA) と 受付(T_KANJA_G) のどちらに入っているかが分かる。'
W '  5・7 で番号の振り方 (日ごとに1から / 4桁ゼロ埋め など) が分かる。'
W '  6 でマイクロンの予約が取り込まれているかが分かる。0件ならまだ取込前。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
