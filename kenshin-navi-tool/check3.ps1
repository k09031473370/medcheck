<#
  自覚症状(017107)の選択肢一覧を取り出す (check3.ps1)
  check3.bat をダブルクリックすると実行され、結果 r_check3.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。個人情報は出力しません(項目名とコードのみ)。

  目的: リアンのフォームの自覚症状コード(020 鼻づまり / 110 動悸・息切れ など)のうち、
        まだ健診ナビ側の結果CDに割り当てられていないものを確定させる。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check3.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$queries = @(
    @{ Title = '1. 自覚症状の項目と、使っている選択肢リスト'
       Sql = "SELECT KOMOKU_CD, MEISYO1, SYOKEN_CD FROM T_KOMOKU WHERE KOMOKU_CD LIKE '017107%' ORDER BY KOMOKU_CD"
       Max = 20 }
    @{ Title = '2. 自覚症状の選択肢一覧 (この結果CDを対応表に書きます)'
       Sql = "SELECT s.SYOKEN_CD, s.KEKKA_CD, s.SYOKEN, s.HANTEI_KIGO FROM T_SYOKEN2 s WHERE LTRIM(RTRIM(s.SYOKEN_CD)) IN (SELECT DISTINCT LTRIM(RTRIM(SYOKEN_CD)) FROM T_KOMOKU WHERE KOMOKU_CD LIKE '017107%' AND SYOKEN_CD IS NOT NULL AND LTRIM(RTRIM(SYOKEN_CD)) <> '') ORDER BY s.SYOKEN_CD, s.KEKKA_CD"
       Max = 200 }
)

"=== 自覚症状の選択肢 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
