<#
.SYNOPSIS
  健診ナビが出した結果票 (Excel) の決まった文字を差し替えて、別名で保存する (text_swap.ps1)

.DESCRIPTION
  例: 裏面の「全国の医療生協では、…」を「当院では、…」にする。
  - 健診ナビもテンプレートも一切触らない。元のファイルも書き換えない (必ず「_当院」を付けた別ファイルに書く)
  - 差し替えるのはセル・図形 (テキストボックス) の中の指定した文字だけ。氏名・検査値・判定には触らない
  - 文字の一部だけ太字・色が違って、文字がいくつかに分かれて入っていても差し替えられる
  - まず何か所変わるかを出し、y を押したときだけ書き出す
  例:
    text_swap.ps1 -In "C:\...\出力フォルダ"                       (既定: 全国の医療生協 → 当院)
    text_swap.ps1 -In "C:\...\結果票.xlsx" -From 全国の医療生協 -To 当院 -Commit
#>
[CmdletBinding()]
param(
    [string]$In,
    [string]$From = '全国の医療生協',
    [string]$To = '当院',
    [string]$Suffix = '_当院',
    [string]$OutDir,
    [switch]$Commit
)
$ErrorActionPreference = 'Stop'
trap { Write-Host (" エラー: {0}`n   {1} 行目: {2}" -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber, ([string]$_.InvocationInfo.Line).Trim()) -ForegroundColor Red; break }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not $In) { $In = (Read-Host '結果票の Excel ファイルかフォルダ (ドラッグして Enter でも可)').Trim('"', ' ') }
if (-not (Test-Path $In)) { throw "見つかりません: $In" }
$files = @(if ((Get-Item $In).PSIsContainer) { Get-ChildItem $In -Filter *.xlsx -File | Where-Object { $_.Name -notlike '~$*' -and $_.BaseName -notlike "*$Suffix" } } else { Get-Item $In })
if ($files.Count -eq 0) { throw "xlsx がありません: $In" }

function XmlEsc([string]$s) { return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') }
function XmlUnesc([string]$s) { return [System.Net.WebUtility]::HtmlDecode($s) }

# 1 つのまとまり (共有文字列 <si>・セルの <is>・図形の段落 <a:p>) の中の文字の部品 (<t>/<a:t>) をつなげて探し、
# 見つかったら、始まりの部品に新しい文字を入れ、残りの部品から元の文字を消す
function Swap-Block([string]$block, [string]$tagRx, [ref]$count) {
    # ふりがな (<rPh>) は印刷されないので、差し替えるまとまりからは外す (ふりがなの文字が間に入って探せなくなるため)
    $block = [regex]::Replace($block, '<rPh\b.*?</rPh>', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $rx = New-Object System.Text.RegularExpressions.Regex("(<$tagRx(?:\s[^>]*)?>)(.*?)(</$tagRx>)", [System.Text.RegularExpressions.RegexOptions]::Singleline)
    for ($guard = 0; $guard -lt 20; $guard++) {
        $ms = $rx.Matches($block)
        if ($ms.Count -eq 0) { return $block }
        $texts = @(foreach ($m in $ms) { XmlUnesc $m.Groups[2].Value })
        $full = -join $texts
        $at = $full.IndexOf($From)
        if ($at -lt 0) { return $block }
        # どの部品の何文字目から始まるか
        $pos = 0; $si = 0
        while ($si -lt $texts.Count -and $pos + $texts[$si].Length -le $at) { $pos += $texts[$si].Length; $si++ }
        $off = $at - $pos
        $remain = $From.Length
        $new = @($texts)
        $new[$si] = $texts[$si].Substring(0, $off) + $To + $texts[$si].Substring([math]::Min($texts[$si].Length, $off + $remain))
        $remain -= [math]::Min($remain, $texts[$si].Length - $off)
        $k = $si + 1
        while ($remain -gt 0 -and $k -lt $texts.Count) {
            $cut = [math]::Min($remain, $texts[$k].Length)
            $new[$k] = $texts[$k].Substring($cut); $remain -= $cut; $k++
        }
        # 後ろから組み直す (位置がずれないように)
        $sb = $block
        for ($i = $ms.Count - 1; $i -ge 0; $i--) {
            if ($new[$i] -eq $texts[$i]) { continue }
            $open = $ms[$i].Groups[1].Value
            if ($open -notmatch 'xml:space=' -and $tagRx -eq 't') { $open = $open.Replace('<t', '<t xml:space="preserve"') }
            $sb = $sb.Substring(0, $ms[$i].Index) + $open + (XmlEsc $new[$i]) + $ms[$i].Groups[3].Value + $sb.Substring($ms[$i].Index + $ms[$i].Length)
        }
        $block = $sb
        $count.Value++
    }
    return $block
}
function Swap-Part([string]$xml, [string]$blockRx, [string]$tagRx, [ref]$count) {
    $brx = New-Object System.Text.RegularExpressions.Regex($blockRx, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $parts = New-Object System.Text.StringBuilder
    $last = 0
    foreach ($b in $brx.Matches($xml)) {
        [void]$parts.Append($xml.Substring($last, $b.Index - $last))
        $plain = XmlUnesc ([regex]::Replace($b.Value, '<[^>]+>', ''))
        if ($plain.Contains($From)) { [void]$parts.Append((Swap-Block $b.Value $tagRx $count)) } else { [void]$parts.Append($b.Value) }
        $last = $b.Index + $b.Length
    }
    [void]$parts.Append($xml.Substring($last))
    return $parts.ToString()
}
# 差し替える部品の一覧と、それぞれの「まとまり」と「文字の部品」
$TARGETS = @(
    @{ Rx = '^xl/sharedStrings\.xml$';   Block = '<si>.*?</si>';          Tag = 't' }
    @{ Rx = '^xl/worksheets/[^/]+\.xml$'; Block = '<is>.*?</is>';          Tag = 't' }
    @{ Rx = '^xl/drawings/[^/]+\.xml$';  Block = '<a:p>.*?</a:p>';        Tag = 'a:t' }
)
function Read-Entry($e) { $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8); try { return $sr.ReadToEnd() } finally { $sr.Dispose() } }

Write-Host ''
Write-Host (" 結果票の文字を差し替える: 「{0}」→「{1}」 (元のファイルは変えず、名前に {2} を付けた別ファイルに書く)" -f $From, $To, $Suffix) -ForegroundColor Cyan
$plan = @()
foreach ($f in $files) {
    $fs = [System.IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $n = 0; $changed = @{}
        foreach ($e in $zip.Entries) {
            foreach ($tg in $TARGETS) {
                if ($e.FullName -notmatch $tg.Rx) { continue }
                $xml = Read-Entry $e
                if (-not (XmlUnesc ([regex]::Replace($xml, '<[^>]+>', ''))).Contains($From)) { continue }
                $c = 0
                $out = Swap-Part $xml $tg.Block $tg.Tag ([ref]$c)
                if ($c -gt 0) { $changed[$e.FullName] = $out; $n += $c }
            }
        }
        $zip.Dispose()
    } finally { $fs.Dispose() }
    $plan += [PSCustomObject]@{ File = $f; Count = $n; Changed = $changed }
    Write-Host ("  {0,4} か所  {1}" -f $n, $f.Name)
}
$todo = @($plan | Where-Object { $_.Count -gt 0 })
if ($todo.Count -eq 0) { Write-Host (" 「{0}」は見つかりませんでした。何も書き出していません。" -f $From) -ForegroundColor Yellow; return }
if (-not $Commit) {
    $ans = Read-Host (" {0} ファイルを書き出しますか？ (y で書き出す)" -f $todo.Count)
    if ($ans -ne 'y') { Write-Host ' やめました。何も書き出していません。'; return }
}
foreach ($p in $todo) {
    $dir = $(if ($OutDir) { $OutDir } else { $p.File.DirectoryName })
    if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
    $dest = Join-Path $dir ($p.File.BaseName + $Suffix + $p.File.Extension)
    if ($dest -eq $p.File.FullName) { throw "元のファイルと同じ名前になります: $dest" }
    Copy-Item $p.File.FullName $dest -Force
    $zip = [System.IO.Compression.ZipFile]::Open($dest, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($name in $p.Changed.Keys) {
            $old = $zip.GetEntry($name); $old.Delete()
            $ne = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $sw = New-Object System.IO.StreamWriter($ne.Open(), (New-Object System.Text.UTF8Encoding $false))
            try { $sw.Write($p.Changed[$name]) } finally { $sw.Dispose() }
        }
    } finally { $zip.Dispose() }
    Write-Host ("  書き出し: {0}" -f $dest) -ForegroundColor Green
}
Write-Host ' 書き出したファイルを Excel で開いて、裏面の文が変わっているか・印刷の向き (A3 両面) を確かめてから印刷してください。' -ForegroundColor DarkGray
