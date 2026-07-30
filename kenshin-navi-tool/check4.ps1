<#
  コースマスタの調査 (check4.ps1)
  check4.bat をダブルクリックすると実行され、結果 r_check4.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。個人情報は出力しません(コードと名称のみ)。

  目的: 予約取込ファイルに書くコースコード/コース名を、健診ナビの実際の値に合わせる。
        (7/2の受診者は COURSE_CD が YA / YB になっていた)
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check4.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$queries = @(
    @{ Title = '1. コースマスタ T_COURSE1 の列一覧'
       Sql = "SELECT c.name AS COL, ty.name AS TYPE FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id WHERE c.object_id = OBJECT_ID('T_COURSE1') ORDER BY c.column_id"
       Max = 60 }
    @{ Title = '2. リアンで使われているコース (YA / YB) の中身'
       Sql = "SELECT * FROM T_COURSE1 WHERE LTRIM(RTRIM(COURSE_CD)) IN ('YA','YB')"
       Max = 10 }
    @{ Title = '3. コース一覧 (コードと名称)'
       Sql = "SELECT * FROM T_COURSE1 ORDER BY COURSE_CD"
       Max = 120 }
)

"=== コースマスタ 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
