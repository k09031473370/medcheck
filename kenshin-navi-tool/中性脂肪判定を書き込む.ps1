<#
.SYNOPSIS
  結果報告書のExcelに、中性脂肪の判定 (H印 と 脂質代謝の判定) を書き込む (中性脂肪判定を書き込む.ps1)

.DESCRIPTION
  健診ナビの基準値マスタに 中性脂肪(R06-0002) の基準が無く、今回の報告書は
  中性脂肪に H の印も付かず、くくり「脂質代謝」の判定にも中性脂肪が入っていない。
  健診ナビ側は触らず (自動判定も走らせず)、出来上がった報告書のExcelだけを直す。

  やること (1人1ファイルの帳票303 .xlsx)
    1. 東振協CSVの中性脂肪 (列103) を読む。Excelの中性脂肪の値と同じことを確かめる
    2. 学会基準で判定する:  F 0〜29 / A 30〜149 / B 150〜299 / D 300〜499 / F 500以上
    3. 書き込む
         ・中性脂肪の H/L の欄 … 150以上は H、30未満は L、それ以外は空のまま
         ・くくり「脂質代謝」の判定 … 今の判定より中性脂肪の判定が悪いときだけ、その判定に上げる
           (A→B, A→D, B→D など。今の方が悪ければそのまま)
         ・中性脂肪（空腹時）の基準値の欄が空なら 30〜149 mg/dl を入れる
           (古いひな形で出した帳票は空になる。「〜」と単位は同じ帳票の総コレステロールの欄に合わせる)
    4. 総合判定は書き換えない。中性脂肪のせいで総合判定まで変わるはずの人がいたら
       「手で確認」として一覧に出すだけ (今回の 8/21 分は 0人)

  安全のため
    ・先に一覧を出して Y を押してから書く
    ・書く前に元ファイルを excel_tools\backup\中性脂肪判定_退避_<日時>\ にコピーする
    ・書き込むのはセルの値だけ (Excelで開いて値を入れて保存)。行や列・書式は触らない
    ・書いたあと、もう一度ファイルを読み直して、狙った値が入ったことを確かめる
    ・DBには接続しない

.EXAMPLE
  中性脂肪判定を書き込む.bat に 結果報告書のフォルダをドラッグ＆ドロップ → CSVを聞かれるので入れる
  そのあと「結果ファイル名をカナに.bat」をもう一度かけて、カナ名のフォルダを作り直す
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Folder,            # 結果報告書(.xlsx)の入ったフォルダ
    [Parameter(Position = 1)]
    [string]$Csv,               # 東振協のCSV
    [string]$Mark,              # H の印の文字。省略時は他の項目に付いている印を見て決める (普通は H)
    [string]$MarkLow,           # L の印の文字。省略時は L
    [string]$Kijun = '30〜149', # 中性脂肪の基準値の欄が空のとき入れる文字 (単位と「〜」は同じファイルの総コレステロールの欄に合わせる)
    [switch]$NoKijun,           # 基準値の欄には書かない
    [switch]$Yes,               # 確認なしで書く
    [string]$CellMap            # 省略時は form\帳票303_セル対応.csv
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression.FileSystem
try { [void][System.Text.Encoding]::GetEncoding(932) }
catch { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) }
if (-not $CellMap) { $CellMap = Join-Path $PSScriptRoot 'form\帳票303_セル対応.csv' }

$COL_TG = 103; $COL_KANJI = 4; $COL_KANA = 5; $COL_YMD = 1
$TG_RYAKU = @('中性脂肪未判別', '中性脂肪', '随時中性脂肪')   # 帳票のひな形によって中性脂肪の枠の略称が違う
$RANK = @{ 'A' = 1; 'B' = 2; 'C' = 3; 'D' = 4; 'E' = 5; 'F' = 6; 'G' = 7; 'J' = 7 }

# 学会基準 (2026/4/1 改定) の中性脂肪
function Judge-TG([double]$v) {
    if ($v -lt 30)  { return 'F' }
    if ($v -lt 150) { return 'A' }
    if ($v -lt 300) { return 'B' }
    if ($v -lt 500) { return 'D' }
    return 'F'
}

# ============================================================================
# 文字・CSV・xlsx の読み方 (結果Excel照合.ps1 と同じ)
# ============================================================================
function Narrow([string]$s) {
    if ($null -eq $s) { return '' }
    $t = [Microsoft.VisualBasic.Strings]::StrConv($s, [Microsoft.VisualBasic.VbStrConv]::Narrow, 1041)
    return ($t -replace '[\s　]+', ' ').Trim()
}
function NoSpace([string]$s) { return ((Narrow $s) -replace ' ', '') }
function Norm-Num([string]$s) {
    $t = (NoSpace $s) -replace '[^0-9.\-]', ''
    $d = 0.0
    if ($t -ne '' -and [double]::TryParse($t, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        return $d.ToString('0.####', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return (NoSpace $s)
}
function Read-CsvRows([string]$path) {
    $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($path, [System.Text.Encoding]::GetEncoding(932))
    $p.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters(',')
    $p.HasFieldsEnclosedInQuotes = $true
    $p.TrimWhiteSpace = $false
    $rows = @()
    while (-not $p.EndOfData) { $rows += ,$p.ReadFields() }
    $p.Close()
    return ,$rows
}
function F($fields, [int]$col) {
    if ($col -le 0 -or $col -gt $fields.Count) { return '' }
    return [string]$fields[$col - 1]
}
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
                elseif ($v) {
                    $raw = $v.InnerText; $d = 0.0
                    if ($t -ne 'str' -and $t -ne 'e' -and [double]::TryParse($raw, [System.Globalization.NumberStyles]::Float,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
                        $val = $d.ToString('0.##########', [System.Globalization.CultureInfo]::InvariantCulture)
                    } else { $val = $raw }
                }
                if ($val -ne '') { $cells[$sheet + '!' + $c.GetAttribute('r')] = $val }
            }
        }
        return $cells
    }
    finally { $zip.Dispose() }
}

# ============================================================================
# 入力
# ============================================================================
if (-not $Folder) { $Folder = Read-Host '結果報告書(.xlsx)の入ったフォルダをドラッグ＆ドロップして Enter' }
$Folder = ($Folder -replace '^"|"$', '').Trim().TrimEnd('\')
if (-not (Test-Path $Folder)) { throw "見つかりません: $Folder" }
if (-not $Csv) { $Csv = Read-Host '東振協のCSVをドラッグ＆ドロップして Enter' }
$Csv = ($Csv -replace '^"|"$', '').Trim()
if (-not (Test-Path $Csv)) { throw "見つかりません: $Csv" }
if (-not (Test-Path $CellMap)) { throw "帳票のセル対応表がありません: $CellMap" }
if ($Folder -like '*\Temp\*' -or $Folder -like '*\Temp*') { throw 'ZIPの中や一時フォルダのファイルには書けません。フォルダを取り出してから実行してください。' }

$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
if ($files.Count -eq 0) { throw "フォルダに .xlsx がありません: $Folder" }

# ---- 帳票のセル対応 ----
#   $tgVal[略称]  = @{ Sheet=2; Cell='CI41' }  (中性脂肪の結果値_1)
#   $tgHL[略称]   = @{ Sheet=2; Cell='CH41' }  (中性脂肪の HL_1)
$tgVal = @{}; $tgHL = @{}; $tgKijun = @(); $tcKijun = $null; $shishitsu = $null; $sogo = $null; $kanjiCell = $null; $kanaCell = $null
$hlKeys = @()      # 全項目の HL_1 のセル (印の文字を調べる用)
foreach ($r in (Import-Csv -Path $CellMap -Encoding UTF8)) {
    $sheetNo = [int]($r.Sheet -replace '\D', '')
    $ref = @{ Sheet = $sheetNo; Cell = $r.Cell; Key = ('sheet{0}.xml!{1}' -f $sheetNo, $r.Cell); Row = [int]($r.Cell -replace '\D', '') }
    if ($r.Ryaku -in $TG_RYAKU) {
        if ($r.Kind -eq '結果値_1') { $tgVal[$r.Ryaku] = $ref }
        if ($r.Kind -eq 'HL_1')    { $tgHL[$r.Ryaku]  = $ref }
        if ($r.Kind -eq '基単')    { $tgKijun += $ref }        # 基準値の欄。値の枠と同じ行のものを使う
    }
    if ($r.Ryaku -eq '総ｺﾚｽﾃﾛｰﾙ' -and $r.Kind -eq '基単') { $tcKijun = $ref }
    if ($r.Ryaku -eq '判定_脂質代謝' -and $r.Kind -eq '判定_1')  { $shishitsu = $ref }
    if ($r.Ryaku -eq '総合判定(編集)' -and $r.Kind -eq '判定_1') { $sogo = $ref }
    if ($r.Ryaku -eq '' -and $r.Kind -eq '漢字氏名' -and -not $kanjiCell) { $kanjiCell = $ref.Key }
    if ($r.Ryaku -eq '' -and $r.Kind -eq 'ｶﾅ氏名'  -and -not $kanaCell)  { $kanaCell  = $ref.Key }
    if ($r.Kind -eq 'HL_1') { $hlKeys += $ref.Key }
}
if ($tgVal.Count -eq 0) { throw 'セル対応表に 中性脂肪 の結果値の枠がありません。' }
if (-not $shishitsu) { throw 'セル対応表に [判定_脂質代謝]判定_1 がありません。' }
if (-not $kanjiCell) { throw 'セル対応表に **漢字氏名 がありません。' }

# ---- CSV ----
$rows = Read-CsvRows $Csv
if ($rows.Count -lt 2) { throw "CSVにデータがありません: $Csv" }
$head = $rows[0]; $data = $rows[1..($rows.Count - 1)]
if ($head.Count -lt $COL_TG) { throw ("列が足りません ({0}列)。東振協の250列のCSVを渡してください。" -f $head.Count) }
$ymd = (F $data[0] $COL_YMD) -replace '[^0-9]', ''
$csvBy = @{}; $csvKana = @{}
foreach ($f in $data) {
    $k = NoSpace (F $f $COL_KANJI); $a = NoSpace (F $f $COL_KANA)
    if ($k -ne '') { $csvBy[$k] = $f }
    if ($a -ne '') { $csvKana[$a] = $f }
}

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 結果報告書(Excel) に 中性脂肪の判定 (H印・脂質代謝) を書き込む' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  Excel: {0} ファイル ({1})" -f $files.Count, $Folder)
Write-Host ("  CSV  : {0} 人 ({1})" -f $data.Count, [System.IO.Path]::GetFileName($Csv))
Write-Host '  基準 : F 0〜29 / A 30〜149 / B 150〜299 / D 300〜499 / F 500以上 (学会 2026/4/1)'
Write-Host ''

# ============================================================================
# 1. 読んで、何を書くか決める (まだ書かない)
# ============================================================================
$plan = @()        # 1人1件
$marksSeen = @{}   # 他の項目に付いている印の文字
$n = 0
foreach ($fx in $files) {
    $n++
    Write-Host ("  [{0}/{1}] {2}" -f $n, $files.Count, $fx.Name) -ForegroundColor DarkGray
    try { $cells = Read-XlsxCells $fx.FullName }
    catch { $plan += New-Object PSObject -Property @{ File = $fx; 氏名 = $fx.Name; 状態 = ('読めません: ' + $_.Exception.Message); Writes = @() }; continue }

    foreach ($k in $hlKeys) { if ($cells.ContainsKey($k)) { $marksSeen[(NoSpace $cells[$k])] = 1 } }

    $xName = ''; if ($cells.ContainsKey($kanjiCell)) { $xName = NoSpace $cells[$kanjiCell] }
    $xKana = ''; if ($kanaCell -and $cells.ContainsKey($kanaCell)) { $xKana = NoSpace $cells[$kanaCell] }
    $label = $(if ($xName -ne '') { $xName } else { $fx.BaseName })
    $f = $null
    if ($xName -ne '' -and $csvBy.ContainsKey($xName)) { $f = $csvBy[$xName] }
    elseif ($xKana -ne '' -and $csvKana.ContainsKey($xKana)) { $f = $csvKana[$xKana] }
    if ($null -eq $f) { $plan += New-Object PSObject -Property @{ File = $fx; 氏名 = $label; 状態 = 'CSVに無い (未受診なら何もしない)'; Writes = @() }; continue }

    # 帳票側の中性脂肪: 値が入っている枠を使う
    $slot = $null; $xlVal = ''
    foreach ($ry in $TG_RYAKU) {
        if ($tgVal.ContainsKey($ry) -and $cells.ContainsKey($tgVal[$ry].Key)) { $v = Narrow $cells[$tgVal[$ry].Key]; if ($v -ne '') { $slot = $ry; $xlVal = $v; break } }
    }
    $csvVal = Narrow (F $f $COL_TG)
    if ($csvVal -eq '' -and $xlVal -eq '') { $plan += New-Object PSObject -Property @{ File = $fx; 氏名 = $label; 状態 = '中性脂肪なし (未検査)'; Writes = @() }; continue }
    if ($csvVal -eq '' -or $xlVal -eq '' -or (Norm-Num $csvVal) -ne (Norm-Num $xlVal)) {
        $plan += New-Object PSObject -Property @{ File = $fx; 氏名 = $label; 状態 = ("★CSV [{0}] と 帳票 [{1}] の中性脂肪が違う。書かない" -f $csvVal, $xlVal); Writes = @() }; continue
    }
    if (-not $tgHL.ContainsKey($slot)) { $plan += New-Object PSObject -Property @{ File = $fx; 氏名 = $label; 状態 = ("★[{0}] の HL の枠が無い。書かない" -f $slot); Writes = @() }; continue }
    $d = [double]::Parse((Norm-Num $csvVal), [System.Globalization.CultureInfo]::InvariantCulture)
    $han = Judge-TG $d

    # 今の値
    $hlKey = $tgHL[$slot].Key
    $hlNow = ''; if ($cells.ContainsKey($hlKey)) { $hlNow = NoSpace $cells[$hlKey] }
    $shNow = ''; if ($cells.ContainsKey($shishitsu.Key)) { $shNow = (NoSpace $cells[$shishitsu.Key]).ToUpper() }
    $sgNow = ''; if ($sogo -and $cells.ContainsKey($sogo.Key)) { $sgNow = (NoSpace $cells[$sogo.Key]).ToUpper() }

    # 印: 150以上 H / 30未満 L / それ以外 空
    $hlNew = ''
    if ($d -ge 150) { $hlNew = '@H' } elseif ($d -lt 30) { $hlNew = '@L' }   # 印の文字は後で決めるので仮
    # 脂質代謝: 中性脂肪の方が悪ければ上げる
    $rTG = $RANK[$han]; $rSH = 0; if ($RANK.ContainsKey($shNow)) { $rSH = $RANK[$shNow] }
    $shNew = $(if ($rTG -gt $rSH) { $han } else { $shNow })
    $rSG = 0; if ($RANK.ContainsKey($sgNow)) { $rSG = $RANK[$sgNow] }
    $note = ''
    if ($sogo -and $rTG -gt $rSG) { $note = ("★総合判定 {0} も {1} になるはず → 手で確認 (書かない)" -f $sgNow, $han) }

    # 基準値の欄 (値の枠と同じ行のもの)。空なら 30〜149 を、同じ帳票の総コレステロールの欄の書き方に合わせて入れる
    $kjNow = ''; $kjNew = ''; $kjRef = $null
    if (-not $NoKijun) {
        $kjRef = $tgKijun | Where-Object { $_.Sheet -eq $tgVal[$slot].Sheet -and $_.Row -eq $tgVal[$slot].Row } | Select-Object -First 1
        if ($kjRef) {
            if ($cells.ContainsKey($kjRef.Key)) { $kjNow = ([string]$cells[$kjRef.Key]).Trim() }
            if ($kjNow -eq '') {
                $tilde = '〜'; $unit = ' mg/dl'
                if ($tcKijun -and $cells.ContainsKey($tcKijun.Key)) {
                    $tc = [string]$cells[$tcKijun.Key]                     # 例: 140〜199 mg/dl
                    if ($tc -match '^\s*\d+(\D+?)\d+(.*)$') { $tilde = $Matches[1]; $unit = $Matches[2] }
                }
                $kjNew = $(if ($Kijun -match '^\s*(\d+)\D+(\d+)\s*$') { $Matches[1] + $tilde + $Matches[2] + $unit } else { $Kijun })
            }
        }
    }

    $writes = @()
    if ($hlNew -ne '' -and $hlNow -eq '')   { $writes += @{ Sheet = $tgHL[$slot].Sheet; Cell = $tgHL[$slot].Cell; Key = $hlKey; What = '中性脂肪の印'; Old = $hlNow; New = $hlNew; Check = @{ Key = $tgVal[$slot].Key; Val = $xlVal } } }
    elseif ($hlNew -ne '' -and $hlNow -ne '') { if ($note) { $note += ' / ' }; $note += ("印は既に [{0}]" -f $hlNow) }
    if ($shNew -ne $shNow) { $writes += @{ Sheet = $shishitsu.Sheet; Cell = $shishitsu.Cell; Key = $shishitsu.Key; What = '脂質代謝の判定'; Old = $shNow; New = $shNew; Check = @{ Key = $tgVal[$slot].Key; Val = $xlVal } } }
    if ($kjNew -ne '')     { $writes += @{ Sheet = $kjRef.Sheet; Cell = $kjRef.Cell; Key = $kjRef.Key; What = '中性脂肪の基準値'; Old = ''; New = $kjNew; Check = @{ Key = $tgVal[$slot].Key; Val = $xlVal } } }

    $plan += New-Object PSObject -Property @{
        File = $fx; 氏名 = $label; 中性脂肪 = $csvVal; 判定 = $han
        印 = $(if ($hlNew -eq '') { '' } else { "$hlNow→$hlNew" })
        脂質代謝 = $(if ($shNew -ne $shNow) { "$shNow→$shNew" } else { $shNow })
        基準値 = $(if ($kjNew -ne '') { "空→$kjNew" } else { $kjNow })
        総合判定 = $sgNow; 状態 = $(if ($writes.Count -gt 0) { '書く' } else { '変更なし' }); 備考 = $note; Writes = $writes
    }
}

# ---- 印の文字を決める ----
$seen = @($marksSeen.Keys | Where-Object { $_ -ne '' } | Sort-Object)
if (-not $Mark) {
    if ($seen.Count -eq 0 -or ($seen -contains 'H')) { $Mark = 'H' }
    else { throw ("他の項目に付いている印が [{0}] で、H ではありません。-Mark で印の文字を指定してください。" -f ($seen -join ' ')) }
}
if (-not $MarkLow) { $MarkLow = 'L' }
foreach ($p in $plan) { foreach ($w in $p.Writes) { if ($w.New -eq '@H') { $w.New = $Mark } elseif ($w.New -eq '@L') { $w.New = $MarkLow } } ; if ($p.印) { $p.印 = $p.印.Replace('@H', $Mark).Replace('@L', $MarkLow) } }

# ============================================================================
# 2. 一覧
# ============================================================================
$toWrite = @($plan | Where-Object { $_.Writes.Count -gt 0 })
$warn    = @($plan | Where-Object { $_.状態 -like '★*' -or $_.状態 -like '読めません*' })
$dist = @{}; foreach ($p in $plan) { if ($p.判定) { $dist[$p.判定] = 1 + $(if ($dist.ContainsKey($p.判定)) { $dist[$p.判定] } else { 0 }) } }

Write-Host ''
Write-Host ('--- 中性脂肪の判定 (学会基準で計算) ---') -ForegroundColor Cyan
foreach ($k in @('A','B','C','D','E','F')) { if ($dist.ContainsKey($k)) { Write-Host ("  {0} : {1,3} 人" -f $k, $dist[$k]) } }
Write-Host ("  他の項目に付いている印の文字: [{0}]  → 中性脂肪にも [{1}] を付けます" -f ($seen -join ' '), $Mark) -ForegroundColor DarkGray
Write-Host ''
$judged = @($toWrite | Where-Object { @($_.Writes | Where-Object { $_.What -ne '中性脂肪の基準値' }).Count -gt 0 })
$kjOnly = @($toWrite | Where-Object { @($_.Writes | Where-Object { $_.What -eq '中性脂肪の基準値' }).Count -gt 0 })
Write-Host ('--- 印 または 脂質代謝の判定 を書く人 ({0} 人) ---' -f $judged.Count) -ForegroundColor Cyan
if ($judged.Count -gt 0) {
    $judged | Select-Object 氏名, 中性脂肪, 判定, 印, 脂質代謝, 総合判定, 備考 | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
}
if ($kjOnly.Count -gt 0) {
    $sample = @($kjOnly | ForEach-Object { ($_.基準値 -replace '^空→', '') } | Select-Object -Unique)
    Write-Host ('--- 中性脂肪の基準値の欄が空なので [{0}] を入れる人: {1} 人 ---' -f ($sample -join ' / '), $kjOnly.Count) -ForegroundColor Cyan
    Write-Host ''
}
if ($warn.Count -gt 0) {
    Write-Host ('--- 注意 ({0} 件) ---' -f $warn.Count) -ForegroundColor Yellow
    $warn | ForEach-Object { Write-Host ("  {0} : {1}" -f $_.氏名, $_.状態) -ForegroundColor Yellow }
    Write-Host ''
}
$sogoNote = @($plan | Where-Object { $_.備考 -like '*総合判定*' })
if ($sogoNote.Count -gt 0) { Write-Host ("★ 総合判定まで変わるはずの人が {0} 人います。総合判定は書き換えないので、健診ナビで直してください。" -f $sogoNote.Count) -ForegroundColor Yellow; Write-Host '' }

if ($toWrite.Count -eq 0) { Write-Host '書き込むものがありません。'; return }

# 書き込みログ (先に控えを作っておく)
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$bkDir = Join-Path (Join-Path $PSScriptRoot 'backup') ("中性脂肪判定_退避_{0}" -f $stamp)
Write-Host ("  書く前の元ファイルは {0} にコピーします" -f $bkDir) -ForegroundColor DarkGray
Write-Host ''
if (-not $Yes) {
    $ans = Read-Host ("上の {0} 人のExcelにセルの値を書き込みます。Excelは閉じてありますか? よければ Y" -f $toWrite.Count)
    if ($ans -notmatch '^[Yy]') { Write-Host '中止しました。何も書いていません。' -ForegroundColor Yellow; return }
}

# ============================================================================
# 3. 書く (Excelを裏で起動して、セルの値だけ入れて保存)
# ============================================================================
[void](New-Item -ItemType Directory -Path $bkDir -Force)
$xl = $null
try { $xl = New-Object -ComObject Excel.Application }
catch { throw 'Excelが起動できません。このPCにExcelが入っているか確認してください。' }
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$okFiles = 0; $ng = @(); $log = @()
try {
    $n = 0
    foreach ($p in $toWrite) {
        $n++
        $fx = $p.File
        Copy-Item -LiteralPath $fx.FullName -Destination (Join-Path $bkDir $fx.Name) -Force
        $wb = $null
        try {
            $wb = $xl.Workbooks.Open($fx.FullName, 0, $false)
            foreach ($w in $p.Writes) {
                $ws = $wb.Worksheets.Item($w.Sheet)
                # 念のため、同じシートの中性脂肪の値が帳票と同じことを見てから書く (シートの取り違え防止)
                $chk = [string]$ws.Range(($w.Check.Key -split '!')[1]).MergeArea.Cells.Item(1, 1).Text
                if ((Norm-Num $chk) -ne (Norm-Num $w.Check.Val)) { throw ("シート{0} の中性脂肪が [{1}] で帳票と合わない" -f $w.Sheet, $chk) }
                $tgt = $ws.Range($w.Cell).MergeArea.Cells.Item(1, 1)      # 結合セルなら左上に書く
                $cur = [string]$tgt.Text
                if ((NoSpace $cur) -ne (NoSpace $w.Old)) { throw ("{0} {1} が [{2}] で想定 [{3}] と違う" -f $w.What, $w.Cell, $cur, $w.Old) }
                $tgt.Value2 = $w.New
                $log += New-Object PSObject -Property @{ 氏名 = $p.氏名; ファイル = $fx.Name; 項目 = $w.What; セル = ("Sheet{0}!{1}" -f $w.Sheet, $w.Cell); 前 = $w.Old; 後 = $w.New }
            }
            $wb.Save()
            $wb.Close($false); $wb = $null
            $okFiles++
            Write-Host ("  [{0}/{1}] {2}  {3}" -f $n, $toWrite.Count, $p.氏名, (($p.Writes | ForEach-Object { "{0} {1}→{2}" -f $_.What, $_.Old, $_.New }) -join ' / ')) -ForegroundColor DarkGray
        } catch {
            $ng += ("{0} : {1}" -f $p.氏名, $_.Exception.Message)
            try { if ($wb) { $wb.Close($false) } } catch {}
            # 失敗したら退避したものを戻す
            Copy-Item -LiteralPath (Join-Path $bkDir $fx.Name) -Destination $fx.FullName -Force
        }
    }
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ============================================================================
# 4. 読み直して確かめる
# ============================================================================
Write-Host ''
Write-Host '--- 読み直して確認 ---' -ForegroundColor Cyan
$bad = @()
foreach ($p in $toWrite) {
    if ($ng | Where-Object { $_ -like ($p.氏名 + ' :*') }) { continue }
    try { $cells = Read-XlsxCells $p.File.FullName } catch { $bad += ("{0} : 読み直せません ({1})" -f $p.氏名, $_.Exception.Message); continue }
    foreach ($w in $p.Writes) {
        $got = ''; if ($cells.ContainsKey($w.Key)) { $got = NoSpace $cells[$w.Key] }
        if ($got -ne (NoSpace $w.New)) { $bad += ("{0} : {1} が [{2}] (期待 [{3}])" -f $p.氏名, $w.What, $got, $w.New) }
    }
    # 中性脂肪の値そのものが変わっていないこと
    foreach ($ry in $TG_RYAKU) {
        if ($tgVal.ContainsKey($ry) -and $cells.ContainsKey($tgVal[$ry].Key)) {
            if ((Norm-Num $cells[$tgVal[$ry].Key]) -ne (Norm-Num $p.中性脂肪)) { $bad += ("{0} : 中性脂肪の値が [{1}] に変わっている" -f $p.氏名, $cells[$tgVal[$ry].Key]) }
        }
    }
}
$logPath = Join-Path $bkDir '_書き込み一覧.csv'
$log | Select-Object 氏名, ファイル, 項目, セル, 前, 後 | Export-Csv -Path $logPath -NoTypeInformation -Encoding Default

$color = $(if ($ng.Count -gt 0 -or $bad.Count -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ("書き込み: {0} ファイル ({1} セル) / 失敗: {2} / 確認NG: {3}" -f $okFiles, $log.Count, $ng.Count, $bad.Count) -ForegroundColor $color
if ($ng.Count -gt 0)  { Write-Host ''; Write-Host '--- 失敗 (元に戻してあります) ---' -ForegroundColor Yellow; $ng  | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow } }
if ($bad.Count -gt 0) { Write-Host ''; Write-Host '--- 確認NG ---' -ForegroundColor Yellow; $bad | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow } }
Write-Host ''
Write-Host ("[控え] 書く前のファイル: {0}" -f $bkDir) -ForegroundColor DarkGray
Write-Host ("[控え] 何をどう書いたか : {0}" -f $logPath) -ForegroundColor DarkGray
Write-Host '※ このあと「結果ファイル名をカナに.bat」をもう一度かけて、カナ名のフォルダを作り直してください。' -ForegroundColor Yellow
Write-Host '※ 健診ナビの中は変えていません (脂質代謝は元の判定のまま)。来年の「前回判定」には元の判定が出ます。' -ForegroundColor Yellow
