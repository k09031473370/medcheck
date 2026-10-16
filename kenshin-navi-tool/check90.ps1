<#
  石屋: 血液の「枠が無かった項目」の枠と値を数える (check90.ps1)
  check90.bat をダブルクリックすると実行され、結果 r_check90.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  対象: eGFR 068676 / HbA1c 068689 / 血小板 069010 / MCV 069007 / MCH 069008 / MCHC 069009
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check90.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 血液項目の枠と値 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★項目ごと: 石屋 199 人のうち 枠がある人数 / 値が入っている人数 (枠が無い人数 = 199 - 枠あり) ---' @"
SELECT i.項目CD, i.項目名,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
         WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
           AND LTRIM(RTRIM(k.KOMOKU_CD)) = i.項目CD) AS 枠あり,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
         WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
           AND LTRIM(RTRIM(k.KOMOKU_CD)) = i.項目CD AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS 値あり
FROM (SELECT '068676' AS 項目CD, N'eGFR' AS 項目名 UNION ALL SELECT '068689', N'HbA1c' UNION ALL SELECT '069010', N'血小板'
      UNION ALL SELECT '069007', N'MCV' UNION ALL SELECT '069008', N'MCH' UNION ALL SELECT '069009', N'MCHC') i
ORDER BY i.項目CD
"@ 10

Q '--- 2. 受診日×コースごとの 枠がある人数 (項目別) ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, p.コース, COUNT(*) AS 人数,
       SUM(p.HbA1c) AS HbA1c, SUM(p.血小板) AS 血小板, SUM(p.MCV) AS MCV, SUM(p.MCH) AS MCH, SUM(p.MCHC) AS MCHC, SUM(p.eGFR) AS eGFR
FROM (
  SELECT s.PK_SEQ, s.D_KENSIN, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
         CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068689') THEN 1 ELSE 0 END AS HbA1c,
         CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069010') THEN 1 ELSE 0 END AS 血小板,
         CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069007') THEN 1 ELSE 0 END AS MCV,
         CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069008') THEN 1 ELSE 0 END AS MCH,
         CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069009') THEN 1 ELSE 0 END AS MCHC,
         CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068676') THEN 1 ELSE 0 END AS eGFR
  FROM T_KENSIN s
  WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
) p
GROUP BY p.D_KENSIN, p.コース ORDER BY 1, 2
"@ 30

W ''
W '=== 読み方 ==='
W '  1: 「枠あり」が 199 より少ない分は、コースにその項目が無い人。足すなら健診ナビで枠を足してから mapping_srl53_rest (残り項目だけ) で 3→4。'
W '  2: どのコースに無いかの内訳。'
notepad $out
