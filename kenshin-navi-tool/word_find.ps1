<#
.SYNOPSIS
  Excel (xlsx) のどこに、ある文字が入っているかを調べる (word_find.ps1)

.DESCRIPTION
  結果票の裏面の「医療生協…」のような文字が、セル・図形 (テキストボックス)・ヘッダー/フッター・コメント の
  どこに入っているかを出す。差し替えの道具を作る前の下調べ用。
  ファイルは読むだけで、一切変更しない (健診ナビのテンプレートを指定しても書き換えない)。
  見つかった文字のまわりだけを出す (氏名や検査値は、探す文字を含まない限り出ない)。
  例:
    word_find.ps1 -In "C:\...\301_健康診断結果結果報告書_事業所あり_A3(両面).xlsx"
    word_find.ps1 -In "C:\...\出力フォルダ" -Word 生協
#>
[CmdletBinding()]
param(
    [string]$In,
    [string]$Word = '生協'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not $In) { $In = (Read-Host 'Excel ファイルかフォルダ (ドラッグして Enter でも可)').Trim('"', ' ') }
if (-not (Test-Path $In)) { throw "見つかりません: $In" }
$files = @(if ((Get-Item $In).PSIsContainer) { Get-ChildItem $In -Filter *.xlsx -File | Where-Object { $_.Name -notlike '~$*' } } else { Get-Item $In })
if ($files.Count -eq 0) { throw "xlsx がありません: $In" }

function Plain([string]$xml) {
    # タグを取って文字だけにする
    $t = [regex]::Replace($xml, '<[^>]+>', ' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ([regex]::Replace($t, '\s+', ' ')).Trim()
}
function Around([string]$text, [string]$w) {
    $i = $text.IndexOf($w)
    if ($i -lt 0) { return '' }
    $s = [math]::Max(0, $i - 30); $e = [math]::Min($text.Length, $i + $w.Length + 50)
    return $text.Substring($s, $e - $s)
}

foreach ($f in $files) {
    Write-Host ''
    Write-Host ("=== {0} ===" -f $f.Name) -ForegroundColor Cyan
    # 開いているファイルでも読めるよう、読み取り共有で開く
    $fs = [System.IO.File]::Open($f.FullName, 'Open', 'Read', 'ReadWrite')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        # シート名 (xl/worksheets/sheetN.xml → シート名)
        $sheetName = @{}
        $wbE = $zip.GetEntry('xl/workbook.xml'); $relE = $zip.GetEntry('xl/_rels/workbook.xml.rels')
        if ($wbE -and $relE) {
            $wb = (New-Object System.IO.StreamReader($wbE.Open())).ReadToEnd()
            $rel = (New-Object System.IO.StreamReader($relE.Open())).ReadToEnd()
            foreach ($m in [regex]::Matches($wb, '<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"')) {
                $rm = [regex]::Match($rel, 'Id="' + [regex]::Escape($m.Groups[2].Value) + '"[^>]*Target="([^"]+)"')
                if (-not $rm.Success) { $rm = [regex]::Match($rel, 'Target="([^"]+)"[^>]*Id="' + [regex]::Escape($m.Groups[2].Value) + '"') }
                if ($rm.Success) { $sheetName['xl/' + ($rm.Groups[1].Value -replace '^/?xl/', '')] = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) }
            }
        }
        # 共有文字列の何番目に入っているか
        $sharedHit = @{}
        $ssE = $zip.GetEntry('xl/sharedStrings.xml')
        if ($ssE) {
            $ss = (New-Object System.IO.StreamReader($ssE.Open())).ReadToEnd()
            $idx = 0
            foreach ($si in [regex]::Matches($ss, '<si>(.*?)</si>', 'Singleline')) {
                $t = Plain $si.Groups[1].Value
                if ($t.Contains($Word)) { $sharedHit[$idx] = $t }
                $idx++
            }
        }
        $found = 0
        foreach ($e in $zip.Entries) {
            if ($e.FullName -notmatch '\.xml$' -or $e.FullName -eq 'xl/sharedStrings.xml') { continue }
            $xml = (New-Object System.IO.StreamReader($e.Open())).ReadToEnd()
            $kind = switch -regex ($e.FullName) { 'worksheets/' { 'シート' } 'drawings/' { '図形・テキストボックス' } 'comments' { 'コメント' } 'charts/' { 'グラフ' } default { 'その他' } }
            $sn = $(if ($sheetName.ContainsKey($e.FullName)) { "「$($sheetName[$e.FullName])」" } else { '' })
            if ($e.FullName -match 'worksheets/') {
                # セル (共有文字列を参照しているもの・直接入っているもの)
                foreach ($c in [regex]::Matches($xml, '<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', 'Singleline')) {
                    $ref = $c.Groups[1].Value; $attr = $c.Groups[2].Value; $body = $c.Groups[3].Value
                    $txt = ''
                    if ($attr -match 't="s"') { $v = [regex]::Match($body, '<v>(\d+)</v>'); if ($v.Success -and $sharedHit.ContainsKey([int]$v.Groups[1].Value)) { $txt = $sharedHit[[int]$v.Groups[1].Value] } }
                    elseif ($body.Contains($Word)) { $txt = Plain $body }
                    if ($txt) { $found++; Write-Host ("  セル  {0}{1}  : {2}" -f $sn, $ref, (Around $txt $Word)) }
                }
                # ヘッダー・フッター
                foreach ($h in [regex]::Matches($xml, '<(oddHeader|oddFooter|evenHeader|evenFooter|firstHeader|firstFooter)>(.*?)</\1>', 'Singleline')) {
                    $t = [System.Net.WebUtility]::HtmlDecode($h.Groups[2].Value)
                    if ($t.Contains($Word)) { $found++; Write-Host ("  ヘッダー/フッター {0} {1} : {2}" -f $sn, $h.Groups[1].Value, (Around $t $Word)) }
                }
            } else {
                $t = Plain $xml
                if ($t.Contains($Word)) { $found++; Write-Host ("  {0}  {1} : {2}" -f $kind, $e.FullName, (Around $t $Word)) }
            }
        }
        if ($sharedHit.Count -gt 0 -and $found -eq 0) { Write-Host '  (共有文字列にはあるが、どのセルからも使われていない)' -ForegroundColor DarkGray }
        if ($found -eq 0 -and $sharedHit.Count -eq 0) {
            $imgs = @($zip.Entries | Where-Object { $_.FullName -match '^xl/media/' })
            Write-Host ("  「{0}」は文字としては見つかりませんでした。画像 {1} 個 (画像の中の文字は探せません)。" -f $Word, $imgs.Count) -ForegroundColor Yellow
        }
        $zip.Dispose()
    } finally { $fs.Dispose() }
}
Write-Host ''
Write-Host ' ファイルは読むだけで、何も変えていません。' -ForegroundColor DarkGray
