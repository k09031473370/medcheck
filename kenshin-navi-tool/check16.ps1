﻿<#
  監督署へ出す「定期健康診断結果報告書(様式第6号)」を健診ナビで出せるか調べる (check16.ps1)
  check16.bat をダブルクリックすると実行され、結果 r_check16.txt がメモ帳で開きます。
  DBもファイルも読むだけで、一切変更しません。

  調べること:
    1. 健診ナビの帳票フォルダに、様式第6号らしいテンプレートがあるか
    2. 帳票の一覧 (何が出せる仕組みになっているか)
    3. 帳票の一覧を持つマスタがDBにあるか
    4. 有所見者数を集計できるか (判定記号がどう入っているか)
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check16.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 様式第6号を健診ナビで出せるか 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default

# ---- 1. 帳票フォルダを見る ----
$roots = @('\\KNSV\KenshinNavi', 'C:\KenshinNavi')
$KEYS  = @('様式', '監督署', '労基', '報告書', '第6号', '第六号', '定期健康診断結果')

W ''
W '--- 1. 様式第6号らしいテンプレートを探す ---'
$anyRoot = $false
foreach ($r in $roots) {
    if (-not (Test-Path $r)) { W "  (無い) $r"; continue }
    $anyRoot = $true
    W "  調べる場所: $r"

    $all = @(Get-ChildItem -Path $r -Include '*.xls', '*.xlsx', '*.xlsm', '*.rpt', '*.frx', '*.pdf' -Recurse -ErrorAction SilentlyContinue)
    W ("  帳票らしいファイル: {0} 個" -f $all.Count)

    $hit = @($all | Where-Object { $n = $_.Name; ($KEYS | Where-Object { $n -like "*$_*" }).Count -gt 0 })
    if ($hit.Count -eq 0) {
        W '  ★ 名前で見つかりませんでした。'
    } else {
        W '  ★ 名前が一致したもの:'
        foreach ($f in $hit) { W ("      {0}  ({1:N0} バイト, {2})" -f $f.FullName, $f.Length, $f.LastWriteTime.ToString('yyyy/MM/dd')) }
    }

    # xlsx の中身も見る。ファイル名が違っても中に「労働基準監督署」と書いてあれば分かる
    W ''
    W '  中身に「労働基準監督署」「様式第6号」が入っているファイル:'
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $inner = 0
    foreach ($f in $all) {
        if ($f.Extension -notin @('.xlsx', '.xlsm')) { continue }
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
            $found = $false
            foreach ($e in $zip.Entries) {
                if ($e.FullName -notmatch '\.xml$') { continue }
                $sr = New-Object System.IO.StreamReader($e.Open())
                $txt = $sr.ReadToEnd(); $sr.Close()
                if ($txt -match '労働基準監督署|様式第6号|様式第六号') { $found = $true; break }
            }
            $zip.Dispose()
            if ($found) { W ("      ★ {0}" -f $f.FullName); $inner++ }
        }
        catch { }
    }
    if ($inner -eq 0) { W '      見つかりませんでした。' }

    # 帳票フォルダの一覧 (何が出せるのかの手がかり)
    W ''
    W '  帳票フォルダにあるファイル (先頭60個):'
    foreach ($f in ($all | Sort-Object Name | Select-Object -First 60)) { W ("      {0}" -f $f.Name) }
    if ($all.Count -gt 60) { W ("      … ほか {0} 個" -f ($all.Count - 60)) }
}
if (-not $anyRoot) {
    W '  健診ナビのフォルダが見えません。'
    W '  自宅PCなど、サーバーに繋がっていない場所で実行した場合はこれで正常です。'
    W '  この部分は、クリニックのPCで実行してください。'
}

# ---- 2. DBに帳票の一覧を持つマスタがあるか ----
if (Test-Path $tool) {
    Q '--- 2. 帳票の一覧を持っていそうなテーブル ---' @"
SELECT TABLE_NAME AS テーブル
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
  AND (TABLE_NAME LIKE '%CHOHYO%' OR TABLE_NAME LIKE '%CYOHYO%'
       OR TABLE_NAME LIKE '%REPORT%' OR TABLE_NAME LIKE '%FORM%'
       OR TABLE_NAME LIKE '%INSATU%' OR TABLE_NAME LIKE '%PRINT%'
       OR TABLE_NAME LIKE '%YOUSIKI%' OR TABLE_NAME LIKE '%SYORUI%')
ORDER BY TABLE_NAME
"@ 60

    Q '--- 3. 判定記号の種類 (有所見者数を数えられるか) ---' @"
SELECT TOP 30 LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定記号, COUNT(*) AS 件数
FROM T_KENSA
WHERE LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) <> ''
GROUP BY LTRIM(RTRIM(ISNULL(HANTEI_KIGO,'')))
ORDER BY COUNT(*) DESC
"@ 40

    Q '--- 4. 様式第6号で数える項目が、DBのどの項目コードに当たるか ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%聴力%' OR MEISYO1 LIKE N'%視力%'
   OR MEISYO1 LIKE N'%胸部%' OR MEISYO1 LIKE N'%喀痰%'
   OR MEISYO1 LIKE N'%血圧%' OR MEISYO1 LIKE N'%貧血%'
   OR MEISYO1 LIKE N'%血色素%' OR MEISYO1 LIKE N'%赤血球%'
   OR MEISYO1 LIKE N'%GOT%' OR MEISYO1 LIKE N'%GPT%' OR MEISYO1 LIKE N'%GTP%'
   OR MEISYO1 LIKE N'%中性脂肪%' OR MEISYO1 LIKE N'%コレステロール%'
   OR MEISYO1 LIKE N'%血糖%' OR MEISYO1 LIKE N'%HbA1c%'
   OR MEISYO1 LIKE N'%尿糖%' OR MEISYO1 LIKE N'%尿蛋白%'
   OR MEISYO1 LIKE N'%心電図%'
ORDER BY KOMOKU_CD
"@ 200
}
else { W ''; W "db_tool.ps1 がありません: $dir" }

W ''
W '=== 読み方 ==='
W '  1 で ★ が付いたファイルがある → 健診ナビで様式第6号を出せます。'
W '     そのテンプレートの「健診機関」欄を差し替えれば、別のクリニック名で出せます。'
W '  1 で何も出ない → 健診ナビには様式第6号の機能が無い可能性が高いです。'
W '     その場合は 3・4 の情報をもとに、有所見者数の集計だけこちらで作れます。'
W '  3 の判定記号が A/B/C や 1/2/3 のように揃っていれば、有所見者を機械的に数えられます。'

notepad $out
