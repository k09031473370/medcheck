<#
  受付番号の自動設定に向けた調査 その2 (check2.ps1)
  check2.bat をダブルクリックすると実行され、結果 r_check2.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。個人情報そのものは出力しません(列名のみ)。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check2.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$queries = @(
    @{ Title = '1. 個人マスタ T_KOJIN1 の列一覧'
       Sql = "SELECT c.name AS COL, ty.name AS TYPE FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id WHERE c.object_id = OBJECT_ID('T_KOJIN1') ORDER BY c.column_id"
       Max = 80 }
    @{ Title = '2. T_KOJIN2 の列一覧'
       Sql = "SELECT c.name AS COL, ty.name AS TYPE FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id WHERE c.object_id = OBJECT_ID('T_KOJIN2') ORDER BY c.column_id"
       Max = 60 }
)

"=== 個人マスタ 列調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
