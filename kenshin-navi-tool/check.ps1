<#
  健診ナビ 設定調査スクリプト (check.ps1)
  check.bat をダブルクリックすると実行され、結果 r_check.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$queries = @(
    @{ Title = '1. 妊娠・生理・既往歴・家族歴・自覚症状の項目コード'
       Sql = "SELECT KOMOKU_CD, MEISYO1, SYOKEN_CD FROM T_KOMOKU WHERE MEISYO1 LIKE N'%妊娠%' OR MEISYO1 LIKE N'%生理%' OR MEISYO1 LIKE N'%既往%' OR MEISYO1 LIKE N'%家族%' OR MEISYO1 LIKE N'%病名%' OR MEISYO1 LIKE N'%転帰%' OR MEISYO1 LIKE N'%自覚%' OR KOMOKU_CD LIKE '017107%' ORDER BY KOMOKU_CD"
       Max = 80 }
    @{ Title = '2. 自覚症状(017107A)の選択肢一覧'
       Sql = "SELECT KEKKA_CD, SYOKEN FROM T_SYOKEN2 WHERE SYOKEN_CD = (SELECT SYOKEN_CD FROM T_KOMOKU WHERE KOMOKU_CD = '017107A') ORDER BY KEKKA_CD"
       Max = 60 }
    @{ Title = '3. 診察(017101A)の項目設定'
       Sql = "SELECT KOMOKU_CD, MEISYO1, SYOKEN_CD FROM T_KOMOKU WHERE KOMOKU_CD LIKE '017101%' OR KOMOKU_CD LIKE '083%' ORDER BY KOMOKU_CD"
       Max = 40 }
)

"=== 健診ナビ 設定調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
