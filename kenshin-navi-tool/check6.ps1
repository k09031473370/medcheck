<#
  血液項目の項目コードを調べる (check6.ps1)
  check6.bat をダブルクリックすると実行され、結果 r_check6.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  目的: SRLのExcelの各列を健診ナビのどの項目に入れるかを確定させる。
        7/2の頃安さん(4002)は血液が既に入っているので、
        「値」から「どの項目コードが何の検査か」を突き合わせられる。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check6.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

$pk0702 = "(SELECT TOP 1 PK_SEQ FROM T_KENSIN WHERE D_KENSIN = '2026/07/02' AND LTRIM(RTRIM(UKE_NO_KENSA)) = '4002' AND F_TORIKESI = 0)"
$pk0705 = "(SELECT TOP 1 PK_SEQ FROM T_KENSIN WHERE D_KENSIN = '2026/07/05' AND LTRIM(RTRIM(UKE_NO_KENSA)) = '4002' AND F_TORIKESI = 0)"

$queries = @(
    @{ Title = '1. 7/2 頃安さんの「値が入っている」項目 (値から検査を特定する)'
       Sql = "SELECT k.KOMOKU_CD, m.MEISYO1 AS 項目名, k.KEKKA AS 結果 FROM T_KENSA k LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)) WHERE k.PK_SEQ = $pk0702 AND k.KEKKA IS NOT NULL AND LTRIM(RTRIM(k.KEKKA)) <> '' ORDER BY k.KOMOKU_CD"
       Max = 200 }
    @{ Title = '2. 7/5 頃安さんの「空いている」項目 (これから血液を入れる枠)'
       Sql = "SELECT k.KOMOKU_CD, m.MEISYO1 AS 項目名 FROM T_KENSA k LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)) WHERE k.PK_SEQ = $pk0705 AND (k.KEKKA IS NULL OR LTRIM(RTRIM(k.KEKKA)) = '') ORDER BY k.KOMOKU_CD"
       Max = 250 }
)

"=== 血液項目の項目コード 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
