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
    @{ Title = '1. コースを持つテーブル/列'
       Sql = "SELECT t.name AS TBL, c.name AS COL FROM sys.tables t JOIN sys.columns c ON c.object_id = t.object_id WHERE c.name LIKE '%COURSE%' OR c.name LIKE '%KOSU%' ORDER BY t.name, c.name"
       Max = 60 }
    @{ Title = '2. 最近つかわれているコースコードと人数'
       Sql = "SELECT COURSE_CD, COUNT(*) AS NINZU, MIN(D_KENSIN) AS FROM_YMD, MAX(D_KENSIN) AS TO_YMD FROM T_KENSIN WHERE D_KENSIN >= '2026/04/01' GROUP BY COURSE_CD ORDER BY COUNT(*) DESC"
       Max = 60 }
    @{ Title = '3. コースマスタ (M_COURSE)'
       Sql = "SELECT * FROM M_COURSE"
       Max = 60 }
    @{ Title = '4. コースマスタ (T_COURSE)'
       Sql = "SELECT * FROM T_COURSE"
       Max = 60 }
)

"=== コースマスタ 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"※ 3か4のどちらかは「テーブルがありません」のエラーになります。それで正常です。" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
