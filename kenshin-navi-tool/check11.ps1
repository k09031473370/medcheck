<#
  「社員番号」がDBのどの欄に入るのかを突き止める (check11.ps1)
  check11.bat をダブルクリックすると実行され、結果 r_check11.txt がメモ帳で開きます。
  DBは読むだけ、帳票も読むだけで、一切変更しません。

  分かっていること:
    帳票の差込み文字 **社員番号 は健診ナビが用意している
    (701_新規 Microsoft Excel Worksheet.xlsx で使われている)。
    ただし、それがDBのどの欄を読んでいるかは分からない。

  やりかた:
    健診ナビの画面で、テスト用の人1人に「社員番号」らしい欄へ
    目印になる値 (既定 99999) を入れて保存してから、これを実行する。
    個人・受診まわりの全部の欄からその値を探して、入った場所を特定する。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check11.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }

Write-Host ''
Write-Host '健診ナビの画面で、テストの人に入れた値を教えてください。'
Write-Host '(まだ入れていない場合は、そのまま Enter を押せば帳票の調査だけ行います)'
$val = Read-Host '探す値 [既定 99999]'
if ([string]::IsNullOrWhiteSpace($val)) { $val = '99999' }
$val = $val.Trim()

"=== 社員番号の入り先さがし $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W "探す値: $val"

if (Test-Path $tool) {
    # 個人・受診まわりの全列から、その値を探す
    $sql = @"
IF OBJECT_ID('tempdb..#r') IS NOT NULL DROP TABLE #r;
CREATE TABLE #r (TBL sysname, COL sysname, 件数 int);
DECLARE @v nvarchar(100) = N'$($val -replace "'", "''")';
DECLARE @sql nvarchar(max) = N'';
SELECT @sql = @sql + N'INSERT INTO #r SELECT ''' + c.TABLE_NAME + N''',''' + c.COLUMN_NAME
     + N''', COUNT(*) FROM [' + c.TABLE_NAME + N'] WHERE CONVERT(nvarchar(200),['
     + c.COLUMN_NAME + N']) LIKE ''%'' + @v + ''%'' HAVING COUNT(*) > 0;'
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME IN ('T_KOJIN1','T_KOJIN2','T_KENSIN','T_KANJA_G')
  AND c.DATA_TYPE IN ('varchar','nvarchar','char','nchar','numeric','int','bigint','smallint','decimal');
EXEC sp_executesql @sql, N'@v nvarchar(100)', @v = @v;
SELECT TBL AS テーブル, COL AS 列, 件数 FROM #r ORDER BY 件数, TBL, COL;
DROP TABLE #r;
"@
    W ''
    W '--- 1. その値が入っている欄 (件数が1なら、そこが社員番号の欄) ---'
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows 60 *>&1 | Out-File $out -Append -Encoding Default
    W ''
    W '※ 何も出ないときは、まだ画面から入れていないか、別の値です。'
    W '※ 件数が多い欄(受付番号など)は、たまたま同じ数字が入っているだけなので無視してください。'
} else {
    W 'db_tool.ps1 がありません'
}

# ---- 701 の中身を見て、社員番号の隣に何が書いてあるか確かめる ----
W ''
W '--- 2. **社員番号 を使っている帳票の中身 ---'
$reader = Join-Path $dir 'xlsx_read.ps1'
if (-not (Test-Path $reader)) {
    W 'xlsx_read.ps1 がありません'
} else {
    . $reader
    $roots = @('\\KNSV\KenshinNavi\11_健康診断結果報告書', '\\KNSV\KenshinNavi', 'C:\KenshinNavi')
    $tpl = $null
    foreach ($r in $roots) {
        if (-not (Test-Path $r)) { continue }
        $hit = Get-ChildItem -Path $r -Filter '701_*.xls*' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { $tpl = $hit.FullName; break }
    }
    if (-not $tpl) {
        W '701_ で始まる帳票が見つかりませんでした。'
    } else {
        W "テンプレート: $tpl"
        W ''
        try {
            $rows = Read-Xlsx $tpl
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $line = @()
                for ($c = 0; $c -lt $rows[$i].Count; $c++) {
                    $v = [string]$rows[$i][$c]
                    if ($v.Trim() -ne '') { $line += ('{0}:{1}' -f ($c + 1), $v.Trim()) }
                }
                if ($line.Count -gt 0) { W ('{0,4}行  {1}' -f ($i + 1), ($line -join ' | ')) }
            }
        } catch { W ("読めませんでした: {0}" -f $_.Exception.Message) }
    }
}

W ''
W '=== 完了。この内容をチャットに貼り付けてください ==='
notepad $out
