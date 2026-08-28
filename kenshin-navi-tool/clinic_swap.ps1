﻿<#
.SYNOPSIS
  帳票のクリニック名を別のクリニックに差し替える (clinic_swap.ps1)

.DESCRIPTION
  健診ナビが出した帳票(Excel)を読み、クリニック欄だけを
  別のクリニックの情報に置き換えた「別ファイル」を作る。
  結果票でも受診表でも、同じように使える。

  - 健診ナビも帳票テンプレートも一切触らない
  - 元のファイルも書き換えない (必ず別名で保存する)
  - 検査結果・氏名・判定には手を付けない。触るのはクリニック欄だけ

  クリニックの一覧: form\clinics.csv
  欄の位置は、いまのクリニック名を手がかりに自動で探すので、登録は不要。
  自動でうまくいかない帳票だけ form\clinic_cells.csv にセル位置を書けば、そちらが優先される。

.EXAMPLE
  # 何が変わるか見る (書き出しません)
  powershell -ExecutionPolicy Bypass -File clinic_swap.ps1 -In "C:\...\帳票.xlsx" -Clinic 〇〇クリニック

  # 書き出す
  powershell -ExecutionPolicy Bypass -File clinic_swap.ps1 -In "C:\...\帳票.xlsx" -Clinic 〇〇クリニック -Commit

  # フォルダごとまとめて
  powershell -ExecutionPolicy Bypass -File clinic_swap.ps1 -In "C:\...\出力フォルダ" -Clinic 〇〇クリニック -Commit
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$In,   # 帳票のファイル、またはフォルダ
    [string]$Clinic,                             # clinics.csv の Key (省略すると一覧を出す)
    [string]$OutDir,                             # 出力先 (省略すると元と同じ場所)
    [switch]$Commit                              # 付けると実際に書き出す
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$MapDir = Join-Path $PSScriptRoot 'form'
$FIELDS = @('法人名', 'クリニック名', '郵便', '住所', '建物', 'TEL', 'FAX')

function Read-CsvJp([string]$path) {
    if (-not (Test-Path $path)) { throw "対応表がありません: $path" }
    # UTF-8 (BOM有無どちらも) でも Shift_JIS でも読めるようにする。
    # Excel で編集すると Shift_JIS で保存されるため、両方を受け付ける必要がある。
    $raw = [System.IO.File]::ReadAllBytes($path)
    if ($raw.Length -ge 3 -and $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF) {
        $txt = [System.Text.Encoding]::UTF8.GetString($raw, 3, $raw.Length - 3)
    }
    else {
        # BOMが無い場合は、UTF-8として厳密に読めるかどうかで判定する
        try {
            $strict = New-Object System.Text.UTF8Encoding($false, $true)
            $txt = $strict.GetString($raw)
        }
        catch {
            $txt = [System.Text.Encoding]::GetEncoding(932).GetString($raw)
        }
    }
    # 1本の文字列のまま渡すと1行として扱われるので、行に分けてから渡す
    $lines = @($txt -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
    if ($lines.Count -lt 2) { return @() }
    return ($lines | ConvertFrom-Csv)
}

# ---- クリニック一覧 ----
$clinics = @(Read-CsvJp (Join-Path $MapDir 'clinics.csv') |
             Where-Object { ([string]$_.Key).Trim() -ne '' -and ([string]$_.クリニック名).Trim() -ne '' })
$layouts = @(Read-CsvJp (Join-Path $MapDir 'clinic_cells.csv') |
             Where-Object { ([string]$_.Layout).Trim() -ne '' })

if (-not $Clinic) {
    Write-Host ''
    Write-Host '=== 使えるクリニック ===' -ForegroundColor Cyan
    foreach ($c in $clinics) {
        Write-Host ("  {0,-10} {1} / {2}" -f $c.Key, $c.法人名, $c.クリニック名)
    }
    Write-Host ''
    Write-Host '  -Clinic <Key> を付けて実行してください。' -ForegroundColor Yellow
    Write-Host '  増やしたいときは form\clinics.csv に行を足します。' -ForegroundColor DarkGray
    return
}

$to = $clinics | Where-Object { $_.Key -eq $Clinic } | Select-Object -First 1
if (-not $to) { throw "clinics.csv に「$Clinic」がありません。-Clinic を付けずに実行すると一覧が出ます。" }

# ---- xlsx の中の1セルを書き換える ----
# 共有文字列(sharedStrings)は他のセルと使い回されているため、そこは触らない。
# 対象セルだけを「そのセル専用の文字列(inlineStr)」に置き換える。
function Set-CellText([string]$xml, [string]$cell, [string]$text) {
    $esc = [System.Security.SecurityElement]::Escape($text)
    if ($null -eq $esc) { $esc = '' }
    # 既存のセル (中身あり / 空タグ の両方)
    $re = [regex]("<c r=""$cell""([^>]*?)(/>|>.*?</c>)")
    $m = $re.Match($xml)
    if (-not $m.Success) { return @{ Xml = $xml; Ok = $false } }
    $attr = $m.Groups[1].Value
    # t= 属性 (型) だけ差し替え、s= (書式) はそのまま残す
    $attr = [regex]::Replace($attr, '\s+t="[^"]*"', '')
    $new = if ($esc -eq '') { "<c r=""$cell""$attr/>" }
           else { "<c r=""$cell""$attr t=""inlineStr""><is><t xml:space=""preserve"">$esc</t></is></c>" }
    return @{ Xml = $re.Replace($xml, [System.Text.RegularExpressions.MatchEvaluator]{ param($x) $new }, 1); Ok = $true }
}

function Get-CellText([string]$xml, [string]$cell, [string[]]$shared) {
    $m = [regex]::Match($xml, "<c r=""$cell""([^>]*?)(/>|>(.*?)</c>)", 'Singleline')
    if (-not $m.Success) { return '' }
    $body = $m.Groups[3].Value
    if ($m.Groups[1].Value -match 't="s"') {
        $v = [regex]::Match($body, '<v>(\d+)</v>')
        if ($v.Success) {
            $i = [int]$v.Groups[1].Value
            if ($i -lt $shared.Count) { return $shared[$i] }
        }
        return ''
    }
    $t = [regex]::Matches($body, '<t[^>]*>(.*?)</t>', 'Singleline')
    if ($t.Count -gt 0) { return (($t | ForEach-Object { $_.Groups[1].Value }) -join '') }
    $v = [regex]::Match($body, '<v>(.*?)</v>', 'Singleline')
    if ($v.Success) { return $v.Groups[1].Value }
    return ''
}

function Get-Shared([System.IO.Compression.ZipArchive]$zip) {
    $e = $zip.GetEntry('xl/sharedStrings.xml')
    if (-not $e) { return @() }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $x = $sr.ReadToEnd(); $sr.Close()
    $out = @()
    foreach ($si in [regex]::Matches($x, '<si>(.*?)</si>', 'Singleline')) {
        $out += (([regex]::Matches($si.Groups[1].Value, '<t[^>]*>(.*?)</t>', 'Singleline') |
                  ForEach-Object { $_.Groups[1].Value }) -join '')
    }
    return $out
}

# セル番地 (AA11) を 列番号・行番号 に分ける
function Split-Ref([string]$ref) {
    $m = [regex]::Match($ref, '^([A-Z]+)(\d+)$')
    if (-not $m.Success) { return $null }
    $col = 0
    foreach ($ch in $m.Groups[1].Value.ToCharArray()) { $col = $col * 26 + ([int][char]$ch - 64) }
    return @{ Col = $col; Row = [int]$m.Groups[2].Value }
}

# シート内の全セルを 番地→文字 の表にする
function Get-AllCells([string]$xml, [string[]]$shared) {
    $map = @{}
    foreach ($m in [regex]::Matches($xml, '<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', 'Singleline')) {
        $ref = $m.Groups[1].Value; $attr = $m.Groups[2].Value; $body = $m.Groups[3].Value
        $t = ''
        if ($attr -match 't="s"') {
            $v = [regex]::Match($body, '<v>(\d+)</v>')
            if ($v.Success) { $i = [int]$v.Groups[1].Value; if ($i -lt $shared.Count) { $t = $shared[$i] } }
        }
        elseif ($attr -match 't="(inlineStr|str)"') {
            $t = ([regex]::Matches($body, '<t[^>]*>(.*?)</t>', 'Singleline') | ForEach-Object { $_.Groups[1].Value }) -join ''
        }
        if ($t -ne '') { $map[$ref] = $t }
    }
    return $map
}

# クリニック名のセルを手がかりに、その周りにあるクリニック欄を自動で見つける。
# 帳票ごとにセル位置を登録しなくても済むようにするため。
# 受診者側にも同じ文字 (郵便番号など) があるので、クリニック名の近くだけを見る。
function Find-ClinicBlock($cells, $clinics) {
    foreach ($c in $clinics) {
        $name = [string]$c.クリニック名
        if ($name -eq '') { continue }
        $anchor = $null
        foreach ($ref in $cells.Keys) { if ($cells[$ref] -eq $name) { $anchor = $ref; break } }
        if (-not $anchor) { continue }

        $a = Split-Ref $anchor
        if (-not $a) { continue }
        $found = @{ 'クリニック名' = $anchor }
        foreach ($fld in $FIELDS) {
            if ($fld -eq 'クリニック名') { continue }
            $val = [string]$c.$fld
            if ($val -eq '') { continue }
            foreach ($ref in $cells.Keys) {
                if ($cells[$ref] -ne $val) { continue }
                $p = Split-Ref $ref
                if (-not $p) { continue }
                # クリニック名の近く (縦12行・横10列) にあるものだけを採用する
                if ([Math]::Abs($p.Row - $a.Row) -le 12 -and $p.Col -ge ($a.Col - 2) -and $p.Col -le ($a.Col + 10)) {
                    $found[$fld] = $ref
                    break
                }
            }
        }
        return @{ From = $c; Cells = $found }
    }
    return $null
}

function Read-Entry([System.IO.Compression.ZipArchive]$zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $t = $sr.ReadToEnd(); $sr.Close()
    return $t
}

function Write-Entry([System.IO.Compression.ZipArchive]$zip, [string]$name, [string]$text) {
    $e = $zip.GetEntry($name)
    if ($e) { $e.Delete() }
    $e = $zip.CreateEntry($name)
    $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($text); $sw.Close()
}

# ---- 対象ファイルを集める ----
if (-not (Test-Path $In)) { throw "見つかりません: $In" }
$targets = @()
if ((Get-Item $In).PSIsContainer) {
    $targets = @(Get-ChildItem -Path $In -Filter '*.xlsx' -File |
                 Where-Object { $_.Name -notlike '*_差替_*' })
} else {
    $targets = @(Get-Item $In)
}
if ($targets.Count -eq 0) { throw "xlsx が1つもありません: $In" }

Write-Host ''
Write-Host '=== 結果票のクリニック名を差し替える ===' -ForegroundColor Cyan
Write-Host ("  差し替え先 : {0} / {1}" -f $to.法人名, $to.クリニック名) -ForegroundColor Green
Write-Host ("  対象       : {0} 件" -f $targets.Count)
if (-not $Commit) { Write-Host '  ※ いまは確認だけです。書き出しません。' -ForegroundColor Yellow }
Write-Host ''

$done = 0; $skip = 0
foreach ($f in $targets) {
    # --- どの帳票かを判別する ---
    $zip = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
    try {
        $shared = Get-Shared $zip
        $hit = $null
        foreach ($ly in $layouts) {
            $sheet = "xl/worksheets/sheet$($ly.シート).xml"
            $xml = Read-Entry $zip $sheet
            if (-not $xml) { continue }
            $now = Get-CellText $xml ([string]$ly.判別セル) $shared
            if ($now -eq '') { continue }
            # いまのクリニック名が clinics.csv のどれかと一致すれば、その帳票と判断する
            $from = $clinics | Where-Object { $_.クリニック名 -eq $now } | Select-Object -First 1
            if ($from) {
                $cellMap = @{}
                foreach ($fld in $FIELDS) { if ([string]$ly.$fld -ne '') { $cellMap[$fld] = [string]$ly.$fld } }
                $hit = @{ Name = [string]$ly.Layout; Sheet = $sheet; From = $from; Cells = $cellMap }
                break
            }
        }

        # 登録が無い帳票 (受診表など) は、クリニック名を手がかりに自動で探す
        if (-not $hit) {
            foreach ($e in $zip.Entries) {
                if ($e.FullName -notmatch '^xl/worksheets/sheet\d+\.xml$') { continue }
                $xml = Read-Entry $zip $e.FullName
                if (-not $xml) { continue }
                $cells = Get-AllCells $xml $shared
                $blk = Find-ClinicBlock $cells $clinics
                if ($blk) {
                    $hit = @{ Name = '自動判別'; Sheet = $e.FullName; From = $blk.From; Cells = $blk.Cells }
                    break
                }
            }
        }
    }
    finally { $zip.Dispose() }

    if (-not $hit) {
        Write-Host ("  [とばす] {0}" -f $f.Name) -ForegroundColor DarkYellow
        Write-Host "           クリニック欄が見つかりませんでした (帳票の種類が違うかもしれません)" -ForegroundColor DarkGray
        $skip++
        continue
    }
    if ($hit.From.Key -eq $to.Key) {
        Write-Host ("  [とばす] {0} … すでに {1} です" -f $f.Name, $to.クリニック名) -ForegroundColor DarkYellow
        $skip++
        continue
    }

    Write-Host ("  {0}" -f $f.Name) -ForegroundColor White
    Write-Host ("      帳票 {0} / {1} → {2}" -f $hit.Name, $hit.From.クリニック名, $to.クリニック名) -ForegroundColor DarkGray

    # --- 出力先を決める。元のファイルには絶対に書かない ---
    $dir = if ($OutDir) { $OutDir } else { $f.DirectoryName }
    if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
    $outFile = Join-Path $dir ("{0}_差替_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($f.Name), $to.Key, $f.Extension)
    if ((Resolve-Path $f.FullName).Path -eq $outFile) { throw "出力先が元ファイルと同じです: $outFile" }

    if (-not $Commit) {
        foreach ($fld in $FIELDS) {
            if (-not $hit.Cells.ContainsKey($fld)) {
                Write-Host ("        {0,-12} (この帳票には無い)" -f $fld) -ForegroundColor DarkGray
                continue
            }
            Write-Host ("        {0,-12} {1,-5} 「{2}」→「{3}」" -f $fld, $hit.Cells[$fld], $hit.From.$fld, $to.$fld) -ForegroundColor DarkGray
        }
        Write-Host ("        書き出し先: {0}" -f $outFile) -ForegroundColor DarkGray
        $done++
        continue
    }

    Copy-Item -Path $f.FullName -Destination $outFile -Force
    $zip = [System.IO.Compression.ZipFile]::Open($outFile, 'Update')
    try {
        $xml = Read-Entry $zip $hit.Sheet
        $n = 0
        foreach ($fld in $FIELDS) {
            if (-not $hit.Cells.ContainsKey($fld)) { continue }
            $cell = $hit.Cells[$fld]
            $r = Set-CellText $xml $cell ([string]$to.$fld)
            if (-not $r.Ok) {
                Write-Host ("        [注意] セル {0} ({1}) が見つかりません" -f $cell, $fld) -ForegroundColor DarkYellow
                continue
            }
            $xml = $r.Xml; $n++
        }
        Write-Entry $zip $hit.Sheet $xml
    }
    finally { $zip.Dispose() }

    Write-Host ("        {0} 欄を差し替えて保存: {1}" -f $n, $outFile) -ForegroundColor Green
    $done++
}

Write-Host ''
if ($Commit) { Write-Host ("[完了] {0} 件を書き出しました (とばした {1} 件)" -f $done, $skip) -ForegroundColor Green }
else {
    Write-Host ("[確認] {0} 件が対象です (とばす {1} 件)" -f $done, $skip) -ForegroundColor Cyan
    Write-Host '実際に書き出すには -Commit を付けて再実行してください。' -ForegroundColor Yellow
}
Write-Host '元のファイルは変更していません。' -ForegroundColor DarkGray
