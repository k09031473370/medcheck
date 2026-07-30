<#
  社員番号・事業所の保存先を調べる (check5.ps1)
  check5.bat をダブルクリックすると実行され、結果 r_check5.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  目的: 結果票に「事業所名」「社員番号」を出すために、
        健診ナビのどの列に入るのかを確定させる。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check5.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$queries = @(
    @{ Title = '1. 社員番号らしい列を持つテーブル'
       Sql = "SELECT t.name AS TBL, c.name AS COL FROM sys.tables t JOIN sys.columns c ON c.object_id = t.object_id WHERE c.name LIKE '%SYAIN%' OR c.name LIKE '%SHAIN%' OR c.name LIKE '%SYOKUIN%' OR c.name LIKE '%BANGO%' OR c.name LIKE '%_NO' ORDER BY t.name, c.name"
       Max = 120 }
    @{ Title = '2. 個人マスタ2 T_KOJIN2 の列一覧'
       Sql = "SELECT c.name AS COL, ty.name AS TYPE FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id WHERE c.object_id = OBJECT_ID('T_KOJIN2') ORDER BY c.column_id"
       Max = 80 }
    @{ Title = '3. 受診情報 T_KENSIN の列一覧'
       Sql = "SELECT c.name AS COL, ty.name AS TYPE FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id WHERE c.object_id = OBJECT_ID('T_KENSIN') ORDER BY c.column_id"
       Max = 120 }
    @{ Title = '4. 7/2の受診者に社員番号らしい値が入っているか (T_KOJIN2)'
       Sql = "SELECT TOP 5 k2.* FROM T_KOJIN2 k2 JOIN T_KENSIN s ON s.KOJIN_ID = k2.KOJIN_ID WHERE s.D_KENSIN = '2026/07/02' AND s.UKE_NO_KENSA >= '4001'"
       Max = 5 }
)

"=== 社員番号・事業所の保存先 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
