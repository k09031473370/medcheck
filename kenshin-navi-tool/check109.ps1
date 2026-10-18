<#
  既往歴の病名コード・転帰コードのマスタを探す (check109.ps1)
  業者の 123 列ファイルの 病名1・2 (19 など) と 転帰 (05 など) が、健診ナビの BYOMEI_CD / CHIRYO と同じコードかを確かめる。
  check109.bat → r_check109.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check109.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 既往歴のコード $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 病名コード・転帰の列を持つ表 (マスタの候補) ---' @"
SELECT c.TABLE_NAME AS 表, c.COLUMN_NAME AS 列 FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.COLUMN_NAME LIKE '%BYOMEI%' OR c.COLUMN_NAME LIKE '%CHIRYO%' OR c.COLUMN_NAME LIKE '%TENKI%' OR c.TABLE_NAME LIKE '%BYOMEI%' OR c.TABLE_NAME LIKE '%KIOU%'
ORDER BY 1, 2
"@ 60

Q '--- 2. ★既往歴で実際に使われている 病名コード → 病名 (2026 年の受診者、本人分) ---' @"
SELECT LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) AS 病名CD, LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) AS 病名, COUNT(*) AS 件数
FROM T_KIOU k WHERE k.KOJIN_ID IN (SELECT s.KOJIN_ID FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2025/01/01')
  AND LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) <> ''
GROUP BY LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))), LTRIM(RTRIM(ISNULL(k.BYOMEI,'')))
ORDER BY CASE WHEN ISNUMERIC(LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,'')))) = 1 THEN CAST(LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) AS int) ELSE 9999 END, 3 DESC
"@ 120

W ''
W '=== 読み方 ==='
W '  2: 業者ファイルの 病名1・2 のコード (19, 16, 34, 06, 07, 47, 51, 33, 35, 25 …) と同じ番号で 高血圧 などが並んでいれば、同じコード体系。'
W '     転帰 (05, 09, 06, 10, 08, 07, 02) は check108 の表 5 の 治療 (5 = 薬剤治療中 など) と同じ番号のはず。'
notepad $out
