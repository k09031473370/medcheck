<#
  10月の予約が健診ナビに入っているかを確かめる (check69.ps1)
  check69.bat をダブルクリックすると実行され、結果 r_check69.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check68 で分かったこと
    ・団体名に「福生」が入るのは3社だけで、全部9月の個別巡回だった
        【福生】嵯峨野株式会社 34人 9/14 / 福生コンクリート工業 25人 9/17 / 関東照明 27人 9/24
    ・10/2・10/6・10/7 の福生市商工会の分が見当たらない
    ・UKE_NO_KENSA は数値型だった (ISNULL(...,'') でエラーになっていた)
    ・T_COURSE1 のジョイン漏れがあった
  → 団体名に「福生」が入っていないだけかもしれないので、日付から探し直す。
#>
[CmdletBinding()]
param(
    [string]$From = '2026/09/25',
    [string]$To   = '2026/10/31'
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check69.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 10月の予約が健診ナビに入っているか ($From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. UKE_NO_KENSA ほか受付番号の列の型 ---' @"
SELECT c.name AS 列, ty.name AS 型, c.max_length AS 長さ, c.precision AS 桁, c.is_nullable AS NULL可
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KENSIN')
  AND c.name IN ('UKE_NO_KENSA','JUSIN_KEN_NO','UKE_NO','KEN_NO')
ORDER BY c.column_id
"@ 20

Q "--- 2. ★$From〜$To の受診日ごとの人数 (団体を問わず全部) ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日, COUNT(*) AS 人数,
       COUNT(DISTINCT s.DANTAI_CD1) AS 団体数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL THEN 0 ELSE 1 END) AS 受付番号あり
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN
"@ 60

Q "--- 3. ★10/2・10/6・10/7 に入っている団体 ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名, COUNT(*) AS 人数
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/10/02','2026/10/06','2026/10/07')
GROUP BY s.D_KENSIN, s.DANTAI_CD1, d.MEISYO1
ORDER BY s.D_KENSIN, COUNT(*) DESC
"@ 100

Q '--- 4. 団体名に「商工」「福生」が入る団体を全部さがす (日付を問わず) ---' @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 団体名,
       (SELECT COUNT(*) FROM T_KENSIN x WHERE x.DANTAI_CD1 = T_DANTAI1.DANTAI_CD1 AND x.F_TORIKESI = 0) AS 受診のべ人数,
       (SELECT MAX(CONVERT(varchar(10), x.D_KENSIN)) FROM T_KENSIN x WHERE x.DANTAI_CD1 = T_DANTAI1.DANTAI_CD1 AND x.F_TORIKESI = 0) AS 最後の受診日
FROM T_DANTAI1
WHERE MEISYO1 LIKE N'%商工%' OR MEISYO1 LIKE N'%福生%'
ORDER BY DANTAI_CD1
"@ 120

Q '--- 5. 9月の福生3社のコースと受付番号 (番号の入り方の見本) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(*) AS 人数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL THEN 0 ELSE 1 END) AS 受付番号あり,
       MIN(s.UKE_NO_KENSA) AS 番号の最小, MAX(s.UKE_NO_KENSA) AS 番号の最大
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.DANTAI_CD1 = s.DANTAI_CD1 AND c.COURSE_CD = s.COURSE_CD
WHERE s.F_TORIKESI = 0 AND d.MEISYO1 LIKE N'%福生%'
GROUP BY s.D_KENSIN, d.MEISYO1, s.COURSE_CD, c.MEISYO
ORDER BY s.D_KENSIN, d.MEISYO1
"@ 60

Q '--- 6. 9月の福生3社の人 先頭30人 (受付番号の実物) ---' @"
SELECT TOP 30 CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       s.UKE_NO_KENSA AS 予約番号,
       (SELECT TOP 1 q.KEN_NO FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付番号
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND d.MEISYO1 LIKE N'%福生%'
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@ 40

Q '--- 7. 直近で受付番号が入っている健診 (振り方の見本をさがす) ---' @"
SELECT TOP 20 CONVERT(varchar(10), s.D_KENSIN) AS 受診日, COUNT(*) AS 人数,
       MIN(s.UKE_NO_KENSA) AS 番号の最小, MAX(s.UKE_NO_KENSA) AS 番号の最大,
       COUNT(DISTINCT s.DANTAI_CD1) AS 団体数
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.UKE_NO_KENSA IS NOT NULL AND s.D_KENSIN >= '2026/04/01'
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN DESC
"@ 30

W ''
W '=== 読み方 ==='
W '  3 が本命。10/2・10/6・10/7 に団体が出てこなければ、'
W '  福生市商工会の予約はまだ健診ナビに取り込まれていない。'
W '  その場合、受付番号を入れる前に まず予約取込 が必要。'
W '  4 で商工会がどの団体CDで登録されているかを確認する。'
W '  5〜7 は、普段の受付番号の振り方を見るため。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
