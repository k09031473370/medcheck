<#
  自動判定画面で石屋 199 人だけを選ぶ絞り方を探す (check110.ps1)
  コース (カンマ区切りで指定できる) と 予約No・受付No の範囲で、石屋と他団体が混ざらないかを見る。
  check110.bat → r_check110.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check110.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 199 人の絞り方 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$DAYS = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')"
$IS = "(s.UKE_NO_KENSA BETWEEN 4631 AND 4831)"

Q '--- 1. ★3 日分の受診者: コース × 石屋/他団体 の人数 (石屋のコースを他団体が使っていなければ、コースで絞れる) ---' @"
SELECT LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
       SUM(CASE WHEN $IS THEN 1 ELSE 0 END) AS 石屋,
       SUM(CASE WHEN $IS THEN 0 ELSE 1 END) AS 他団体
FROM T_KENSIN s WHERE $DAYS
GROUP BY LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) ORDER BY 2 DESC, 1
"@ 60

Q '--- 2. T_KENSIN の「番号」らしい列 ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME = 'T_KENSIN' AND (c.COLUMN_NAME LIKE '%NO%' OR c.COLUMN_NAME LIKE '%YOYAKU%' OR c.COLUMN_NAME LIKE '%KARUTE%') ORDER BY c.ORDINAL_POSITION
"@ 40

Q '--- 3. ★受診日ごと、番号の列の 石屋の最小〜最大 と、その範囲に入る他団体の人数 (0 なら範囲で絞れる) ---' @"
DECLARE @sql nvarchar(max) = N'';
SELECT @sql = @sql + CASE WHEN @sql = N'' THEN N'' ELSE N' UNION ALL ' END +
 N'SELECT ''' + c.COLUMN_NAME + N''' AS 列, d.YMD AS 受診日, d.MN AS 石屋の最小, d.MX AS 石屋の最大, d.N AS 石屋, ' +
 N'(SELECT COUNT(*) FROM T_KENSIN o WHERE o.F_TORIKESI = 0 AND CONVERT(varchar(10), o.D_KENSIN, 111) = d.YMD AND NOT (o.UKE_NO_KENSA BETWEEN 4631 AND 4831) AND o.' + c.COLUMN_NAME + N' BETWEEN d.MN AND d.MX) AS 範囲内の他団体 ' +
 N'FROM (SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, MIN(s.' + c.COLUMN_NAME + N') AS MN, MAX(s.' + c.COLUMN_NAME + N') AS MX, COUNT(*) AS N FROM T_KENSIN s ' +
 N'WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN (''2026/09/03'',''2026/09/09'',''2026/09/10'') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831 GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111)) d'
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME = 'T_KENSIN' AND c.DATA_TYPE IN ('numeric','int','decimal','bigint','smallint')
  AND (c.COLUMN_NAME LIKE '%NO%' OR c.COLUMN_NAME LIKE '%YOYAKU%');
SET @sql = @sql + N' ORDER BY 1, 2';
EXEC sp_executesql @sql;
"@ 60

W ''
W '=== 読み方 ==='
W '  1: 石屋のコースの「他団体」が 0 なら、自動判定画面の下の「コース」にカンマ区切りで入れて絞れる。'
W '  3: 「範囲内の他団体」が 0 の列なら、その列 (予約No など) の範囲で受診日ごとに絞れる。'
notepad $out
