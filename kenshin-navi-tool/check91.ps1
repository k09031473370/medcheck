<#
  石屋: SRL (血液) の項目のうち、コースに枠が無いものをコース別に出す (check91.ps1)
  check91.bat をダブルクリックすると実行され、結果 r_check91.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  mapping_srl53.csv で入れている 25 項目について、受診日×コースごとに「人数」と「枠がある人数」を数え、
  枠が足りない (枠あり < 人数) 組み合わせだけを出す。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check91.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 SRL 項目の枠が無いコース $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$items = @"
  SELECT '068604' AS CD, N'ｱﾙﾌﾞﾐﾝ' AS NM, 1 AS JUN UNION ALL SELECT '068674', N'ｸﾚｱﾁﾆﾝ', 2 UNION ALL SELECT '068675', N'尿酸', 3
  UNION ALL SELECT '068705', N'総ｺﾚｽﾃﾛｰﾙ', 4 UNION ALL SELECT '068709', N'LDL', 5 UNION ALL SELECT '068708', N'HDL', 6
  UNION ALL SELECT 'R06-0002', N'中性脂肪', 7 UNION ALL SELECT '068636', N'AST', 8 UNION ALL SELECT '068637', N'ALT', 9
  UNION ALL SELECT '068645', N'ALP', 10 UNION ALL SELECT '068652', N'γ-GTP', 11 UNION ALL SELECT '068682', N'血糖', 12
  UNION ALL SELECT '068689', N'HbA1c', 13 UNION ALL SELECT '069003', N'白血球', 14 UNION ALL SELECT '069004', N'赤血球', 15
  UNION ALL SELECT '069005', N'血色素', 16 UNION ALL SELECT '069006', N'ﾍﾏﾄｸﾘｯﾄ', 17 UNION ALL SELECT '069010', N'血小板', 18
  UNION ALL SELECT '068676', N'eGFR', 19 UNION ALL SELECT '069007', N'MCV', 20 UNION ALL SELECT '069008', N'MCH', 21
  UNION ALL SELECT '069009', N'MCHC', 22 UNION ALL SELECT '069245', N'便潜血1', 23 UNION ALL SELECT '069246', N'便潜血2', 24
  UNION ALL SELECT '068782', N'総ﾋﾞﾘﾙﾋﾞﾝ', 25
"@

Q '--- 1. ★受診日×コースごとに、枠が足りない SRL 項目 (枠あり < 人数 のものだけ) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(DISTINCT s.PK_SEQ) AS 人数, i.NM AS 項目, i.CD AS 項目CD,
       COUNT(DISTINCT k.PK_SEQ) AS 枠あり, COUNT(DISTINCT s.PK_SEQ) - COUNT(DISTINCT k.PK_SEQ) AS 枠なし
FROM T_KENSIN s
CROSS JOIN ($items) i
LEFT JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = i.CD
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
GROUP BY s.D_KENSIN, s.COURSE_CD, c.MEISYO, i.NM, i.CD, i.JUN
HAVING COUNT(DISTINCT k.PK_SEQ) < COUNT(DISTINCT s.PK_SEQ)
ORDER BY 1, 2, i.JUN
"@ 200

Q '--- 2. 参考: コース別の人数 (3 日分合計) ---' @"
SELECT LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名, COUNT(*) AS 人数
FROM T_KENSIN s LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
GROUP BY s.COURSE_CD, c.MEISYO ORDER BY 1
"@ 20

W ''
W '=== 読み方 ==='
W '  1: 出てきた行が「そのコースのその項目は枠が無い (= SRL の値を入れられていない)」。出ていない項目は全員入っている。'
W '     同じコースでも事業所によって枠が違うことがあるので、枠なしが人数より少ない行は事業所差。'
W '     足すなら、健診ナビで枠を足してから mapping_srl53_rest (HbA1c・血小板・MCV・MCH・MCHC・eGFR) か、'
W '     それ以外の項目なら通常の SRL 53 列版で 3→4 (他の項目は同じ値を書き直すだけ)。'
notepad $out
