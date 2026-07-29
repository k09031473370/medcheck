<#
  受付番号の自動設定に向けた調査 (check2.ps1)
  check2.bat をダブルクリックすると実行され、結果 r_check2.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check2.txt'
$ymd = '2026/07/02'   # 調査対象の受診日

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$queries = @(
    @{ Title = "1. $ymd の受診者 (受付番号と個人ID)"
       Sql = "SELECT UKE_NO_KENSA, KOJIN_ID, DANTAI_CD1, COURSE_CD, F_UKETUKE FROM T_KENSIN WHERE D_KENSIN = '$ymd' ORDER BY KOJIN_ID"
       Max = 40 }
    @{ Title = '2. 個人ID(KOJIN_ID)を持つテーブル'
       Sql = "SELECT t.name AS TBL FROM sys.tables t JOIN sys.columns c ON c.object_id = t.object_id AND c.name = 'KOJIN_ID' ORDER BY t.name"
       Max = 40 }
    @{ Title = '3. 氏名・カナ・生年月日らしい列を持つテーブル'
       Sql = "SELECT t.name AS TBL, c.name AS COL FROM sys.tables t JOIN sys.columns c ON c.object_id = t.object_id WHERE c.name LIKE '%KANA%' OR c.name LIKE '%SIMEI%' OR c.name LIKE '%NAME%' OR c.name LIKE '%SEI_YMD%' OR c.name LIKE '%BIRTH%' OR c.name LIKE '%SYAIN%' ORDER BY t.name, c.name"
       Max = 60 }
)

"=== 受付番号の自動設定 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
