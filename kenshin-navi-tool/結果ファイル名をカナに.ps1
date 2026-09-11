<#
.SYNOPSIS
  結果報告書のExcelを、ファイル名の氏名をカナにしたコピーにする (結果ファイル名をカナに.ps1)

.DESCRIPTION
  健診ナビが出す結果報告書のファイル名は
      20260821_0006_秋葉達也_303_健康診断結果結果報告書_事業所あり_A3両面_LAP.xlsx
  のように漢字氏名が入る。これを、Excelの中のｶﾅ氏名のセルを読んで
      20260821_0006_アキバタツヤ_303_健康診断結果結果報告書_事業所あり_A3両面_LAP.xlsx
  というファイル名でコピーする。元のフォルダ・元のファイルは触らない。

  ・出力先は <元フォルダ>_カナ (無ければ作る)
  ・中身が空のファイル (未受診) はコピーしない。一覧に出すだけ
  ・カナは全角カタカナ、姓名の間の空白は詰める
  ・-KanaFirst を付けると アキバタツヤ_20260821_0006_303_… の形 (あいうえお順に並ぶ)

.EXAMPLE
  結果ファイル名をカナに.bat に、結果報告書の入ったフォルダをドラッグ＆ドロップ
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Folder,
    [string]$OutDir,
    [switch]$KanaFirst,
    [string]$CellMap
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression.FileSystem
if (-not $CellMap) { $CellMap = Join-Path $PSScriptRoot 'form\帳票303_セル対応.csv' }

function Read-XlsxCells([string]$path) {
    $NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $getXml = {
            param($name)
            foreach ($e in $zip.Entries) {
                if ($e.FullName -ne $name) { continue }
                $st = $e.Open(); $sr = New-Object System.IO.StreamReader($st, [System.Text.Encoding]::UTF8)
                $txt = $sr.ReadToEnd(); $sr.Close(); $st.Close()
                $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($txt); return ,$doc
            }
            return $null
        }
        $shared = @()
        $xs = & $getXml 'xl/sharedStrings.xml'
        if ($xs) {
            $m = New-Object System.Xml.XmlNamespaceManager($xs.NameTable); $m.AddNamespace('m', $NS)
            foreach ($si in $xs.SelectNodes('/m:sst/m:si', $m)) {
                $sb = New-Object System.Text.StringBuilder
                foreach ($t in $si.SelectNodes('m:t | m:r/m:t', $m)) { [void]$sb.Append($t.InnerText) }
                $shared += $sb.ToString()
            }
        }
        $cells = @{}
        foreach ($e in @($zip.Entries)) {
            if ($e.FullName -notmatch '^xl/worksheets/(sheet\d+\.xml)$') { continue }
            $sheet = $Matches[1]
            $doc = & $getXml $e.FullName
            $m = New-Object System.Xml.XmlNamespaceManager($doc.NameTable); $m.AddNamespace('m', $NS)
            foreach ($c in $doc.SelectNodes('/m:worksheet/m:sheetData/m:row/m:c', $m)) {
                $t = $c.GetAttribute('t'); $v = $c.SelectSingleNode('m:v', $m); $val = ''
                if ($t -eq 'inlineStr') { foreach ($tn in $c.SelectNodes('m:is//m:t', $m)) { $val += $tn.InnerText } }
                elseif ($t -eq 's') { if ($v) { $ix = -1; if ([int]::TryParse($v.InnerText, [ref]$ix) -and $ix -ge 0 -and $ix -lt $shared.Count) { $val = $shared[$ix] } } }
                elseif ($v) { $val = $v.InnerText }
                if ($val -ne '') { $cells[$sheet + '!' + $c.GetAttribute('r')] = $val }
            }
        }
        return $cells
    }
    finally { $zip.Dispose() }
}
# 半角カナ → 全角カタカナ。空白は詰める。ファイル名に使えない文字は除く
function To-KanaName([string]$s) {
    if ($null -eq $s) { return '' }
    $w = [Microsoft.VisualBasic.Strings]::StrConv($s, [Microsoft.VisualBasic.VbStrConv]::Wide, 1041)
    $w = $w -replace '[\s　]+', ''
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) { $w = $w.Replace([string]$ch, '') }
    return $w
}

if (-not $Folder) { $Folder = Read-Host '結果報告書(.xlsx)の入ったフォルダをドラッグ＆ドロップして Enter' }
$Folder = ($Folder -replace '^"|"$', '').Trim().TrimEnd('\')
if (-not (Test-Path $Folder)) { throw "見つかりません: $Folder" }
if (-not (Test-Path $CellMap)) { throw "帳票のセル対応表がありません: $CellMap" }
if (-not $OutDir) { $OutDir = $Folder + '_カナ' }
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
elseif ($OutDir -like '*_カナ') {
    # 自分が作った出力フォルダなら、前回の結果を消してから作り直す
    Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Extension -in '.xlsx', '.csv' } | Remove-Item -Force
}

$kanaCell = $null; $kanjiCell = $null; $checkCells = @()
foreach ($r in (Import-Csv -Path $CellMap -Encoding UTF8)) {
    $key = ($r.Sheet -replace '^Sheet', 'sheet') + '.xml!' + $r.Cell
    if ($r.Ryaku -eq '' -and $r.Kind -eq 'ｶﾅ氏名')  { $kanaCell  = $key }
    if ($r.Ryaku -eq '' -and $r.Kind -eq '漢字氏名') { $kanjiCell = $key }
    if ($r.Kind -eq '結果値_1' -and $r.Ryaku -in @('身長','体重','GOT','白血球数','視力右')) { $checkCells += $key }
}
if (-not $kanaCell) { throw 'セル対応表に **ｶﾅ氏名 がありません。' }

$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
if ($files.Count -eq 0) { throw "フォルダに .xlsx がありません: $Folder" }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 結果報告書のファイル名を カナ氏名 にしてコピー' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  元  : {0}  ({1} ファイル)" -f $Folder, $files.Count)
Write-Host ("  出力: {0}" -f $OutDir)
Write-Host ''

$done = @(); $skip = @(); $warn = @()
$n = 0
foreach ($fx in $files) {
    $n++
    try { $cells = Read-XlsxCells $fx.FullName } catch { $warn += ("{0} : 読めません ({1})" -f $fx.Name, $_.Exception.Message); continue }
    $filled = 0; foreach ($k in $checkCells) { if ($cells.ContainsKey($k)) { $filled++ } }
    if ($filled -eq 0) { $skip += $fx.Name; continue }     # 未受診 (中身が空)
    $kana = ''; if ($cells.ContainsKey($kanaCell)) { $kana = To-KanaName $cells[$kanaCell] }
    if ($kana -eq '') { $warn += ("{0} : カナ氏名が空なので元の名前のままコピー" -f $fx.Name); $newName = $fx.Name }
    else {
        # 20260821_0006__秋葉達也_303_… のように、氏名の前に空の区切りが入ることがある。
        # Excelの中の漢字氏名と同じ区切りを探して、そこをカナにする。
        $kanji = ''; if ($kanjiCell -and $cells.ContainsKey($kanjiCell)) { $kanji = ([string]$cells[$kanjiCell]) -replace '[\s　]+', '' }
        $parts = @($fx.BaseName -split '_')
        $ix = -1
        for ($i = 0; $i -lt $parts.Count; $i++) {
            if ($kanji -ne '' -and (($parts[$i] -replace '[\s　]+', '') -eq $kanji)) { $ix = $i; break }
        }
        if ($ix -lt 0) {
            # 見つからなければ、日付・番号の次の「空でない」区切りを氏名とみなす
            for ($i = 2; $i -lt $parts.Count; $i++) { if ($parts[$i] -ne '') { $ix = $i; break } }
        }
        if ($ix -lt 0) { $newName = $kana + '_' + $fx.Name }
        else {
            $parts[$ix] = $kana
            if ($KanaFirst) {
                $rest = @($parts | Where-Object { $_ -ne $kana })
                $newName = ($kana + '_' + ($rest -join '_')) + $fx.Extension
            } else {
                $newName = ($parts -join '_') + $fx.Extension
            }
        }
    }
    $dst = Join-Path $OutDir $newName
    if (Test-Path $dst) { $warn += ("{0} : 同じ名前が既にあるので上書き" -f $newName) }
    Copy-Item -LiteralPath $fx.FullName -Destination $dst -Force
    $done += New-Object PSObject -Property @{ 元 = $fx.Name; 新 = $newName }
    Write-Host ("  [{0}/{1}] {2}" -f $n, $files.Count, $newName) -ForegroundColor DarkGray
}

Write-Host ''
Write-Host ("コピー: {0} ファイル / 未受診でコピーしない: {1} ファイル / 注意: {2} 件" -f $done.Count, $skip.Count, $warn.Count) -ForegroundColor Green
if ($skip.Count -gt 0) { Write-Host ''; Write-Host '--- 未受診 (中身が空) なのでコピーしていないもの ---' -ForegroundColor Yellow; $skip | ForEach-Object { Write-Host ('  ' + $_) } }
if ($warn.Count -gt 0) { Write-Host ''; Write-Host '--- 注意 ---' -ForegroundColor Yellow; $warn | ForEach-Object { Write-Host ('  ' + $_) } }

$list = Join-Path $OutDir '_ファイル名対応.csv'
$done | Select-Object 元, 新 | Export-Csv -Path $list -NoTypeInformation -Encoding Default
Write-Host ''
Write-Host ("[一覧] 元の名前と新しい名前の対応: {0}" -f $list) -ForegroundColor DarkGray
Write-Host '※ 元のフォルダは変えていません。出力フォルダの中身を渡してください。' -ForegroundColor Yellow
Write-Host '※ _ファイル名対応.csv は当院の控えです。渡すときは外してください。' -ForegroundColor Yellow
