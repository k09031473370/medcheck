<#
.SYNOPSIS
  取込ファイルが対応表と合っているかを、書き込まずに確かめる (ファイル確認.ps1)

.DESCRIPTION
  届いた結果ファイル(CSV / xlsx)を読んで、次を報告する。
    ・どの対応表(レイアウト)と判定されるか
    ・列数が対応表と合っているか
    ・対応表の各項目に、実際どんな値が入っているか
    ・受付番号・氏名の入り方
    ・対応表に無い列、対応表にあるのに足りない列
    ・変換表(code_map)に無い値

  データベースには一切つながらない。ファイルを読むだけ。
  取込の前にこれを通しておけば、当日エラーで止まらずに済む。

.EXAMPLE
  ファイル確認.bat に取込ファイルをドラッグ＆ドロップ

  powershell -ExecutionPolicy Bypass -File ファイル確認.ps1 -Csv "C:\...\20260821.csv"
  powershell -ExecutionPolicy Bypass -File ファイル確認.ps1 -Csv "...\file.csv" -CsvEncoding UTF8
#>
[CmdletBinding()]
param(
    [string]$Csv,                     # 確かめたいファイル
    [string]$Mapping = 'auto',        # 対応表を指定したいとき
    [string]$CsvEncoding = 'SJIS',    # CSVの文字コード。文字化けするなら UTF8
    [switch]$NoHeader,                # 見出し行が無いファイル
    [string]$MapDir,
    [int]$Sample = 3                  # 何人目までを見本として出すか
)

$ErrorActionPreference = 'Stop'
if (-not $MapDir) { $MapDir = Join-Path $PSScriptRoot 'form' }

if (-not $Csv) {
    Write-Host 'ファイルを指定してください。ファイル確認.bat にドラッグ＆ドロップするのが簡単です。' -ForegroundColor Yellow
    return
}
if (-not (Test-Path $Csv)) { throw "ファイルが見つかりません: $Csv" }

# xlsx を読む部品
$XlsxLib = Join-Path $PSScriptRoot 'xlsx_read.ps1'
if (Test-Path $XlsxLib) { . $XlsxLib }

# form_import.ps1 から読み取りまわりを借りる (本番と同じ判定をさせるため)
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
# Get-LocalConnFile の手前まで借りる。ここまでに接続まわりは入らないので、DBには繋がらない。
# ただし Parse-CsvText と Test-XlsxEncrypted の間には
#     $XlsxLib = Join-Path $PSScriptRoot 'xlsx_read.ps1'
# という関数の外の行がある。Invoke-Expression の中では $PSScriptRoot が空になり
# Join-Path が失敗するので、そこは飛ばして2回に分けて借りる。
Invoke-Expression (Get-Part 'function Normalize-Text' '# .xlsx / .xlsm の読み込み')
Invoke-Expression (Get-Part 'function Test-XlsxEncrypted' 'function Get-LocalConnFile')

function Line($c = 'Gray') { Write-Host ('-' * 78) -ForegroundColor $c }

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("取込ファイルの確認  {0}" -f (Get-Date -Format 'yyyy/MM/dd HH:mm')) -ForegroundColor Cyan
Write-Host ("ファイル: {0}" -f $Csv) -ForegroundColor Cyan
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host '※ このスクリプトはデータベースに繋ぎません。ファイルを読むだけです。'

# ---------------------------------------------------------------- 1. 読む
$rows = Read-FormCsv $Csv $CsvEncoding
$maxCols = 0
foreach ($r in $rows) { if ($r.Count -gt $maxCols) { $maxCols = $r.Count } }

Write-Host ''
Write-Host '1. ファイルの形' -ForegroundColor White
Line
Write-Host ("  行数(空行を除く): {0}" -f $rows.Count)
Write-Host ("  列数(最大)      : {0}" -f $maxCols)
Write-Host ("  文字コード指定  : {0}" -f $CsvEncoding)

$head = $rows[0]
$garbled = 0
foreach ($h in $head) { if ($h -match '[\uFFFD]') { $garbled++ } }
if ($garbled -gt 0) {
    Write-Host ("  ★ 見出しが文字化けしています。-CsvEncoding UTF8 を付けて試してください。") -ForegroundColor Red
}

Write-Host ''
Write-Host '  1行目:'
for ($i = 0; $i -lt $head.Count; $i++) {
    Write-Host ("    {0,3}: {1}" -f ($i + 1), (Normalize-Text $head[$i]))
}

# ---------------------------------------------------------------- 2. 対応表
Write-Host ''
Write-Host '2. どの対応表(レイアウト)になるか' -ForegroundColor White
Line
$mapPath = $null
if ($Mapping -and $Mapping -ne 'auto') {
    $mapPath = if ([System.IO.Path]::IsPathRooted($Mapping)) { $Mapping } else { Join-Path $MapDir $Mapping }
    Write-Host ("  指定: {0}" -f (Split-Path $mapPath -Leaf))
}
else {
    $mapPath = Detect-MappingPath $Csv
    if ($mapPath) {
        Write-Host ("  ★ 自動判別: {0}" -f (Split-Path $mapPath -Leaf)) -ForegroundColor Green
    }
    else {
        Write-Host '  ★ 自動判別できませんでした。' -ForegroundColor Red
        Write-Host '     取込のときは、GUIの「レイアウト」で手で選ぶことになります。'
        Write-Host '     候補:'
        foreach ($f in (Get-ChildItem -Path $MapDir -Filter 'mapping*.csv' -File | Sort-Object Name)) {
            $d = @(Get-Content $f.FullName -TotalCount 8 -Encoding UTF8 | Where-Object { $_ -match '^#\s*DETECT' })
            Write-Host ("       {0}  {1}" -f $f.Name, ($d -join ' '))
        }
        Write-Host ''
        Write-Host '  ここで止めます。対応表が決まらないと中身の確認ができません。' -ForegroundColor Yellow
        return
    }
}
if (-not (Test-Path $mapPath)) { throw "対応表が見つかりません: $mapPath" }

$map = @(Import-Csv -Path $mapPath -Encoding UTF8)
$mapUse = @($map | Where-Object {
    (Normalize-Text $_.Col) -ne '' -and (Normalize-Text $_.Kind).ToUpper() -ne 'IGNORE'
})
Write-Host ("  対応表の行数: {0} (うち取込対象 {1})" -f $map.Count, $mapUse.Count)

# 列数の食い違い
$needMax = 0
foreach ($m in $map) {
    foreach ($cName in @('Col', 'Col2')) {
        $c = Normalize-Text $m.$cName
        if ($c -match '^\d+$' -and [int]$c -gt $needMax) { $needMax = [int]$c }
    }
}
Write-Host ("  対応表が使う一番右の列: {0}" -f $needMax)
if ($needMax -gt $maxCols) {
    Write-Host ("  ★ ファイルの列が足りません (ファイル {0} 列 / 対応表は {1} 列目まで使う)" -f $maxCols, $needMax) -ForegroundColor Red
    Write-Host '     レイアウト違いの可能性があります。'
}
elseif ($maxCols -gt $needMax) {
    Write-Host ("  ・ファイルの方が {0} 列多い (右側の列は使いません)" -f ($maxCols - $needMax)) -ForegroundColor DarkGray
}
else {
    Write-Host '  ・列数は対応表とぴったりです。' -ForegroundColor Green
}

# ---------------------------------------------------------------- 3. 見出し/データ分離
$idCols = @{ KenNo = 0; KenYmd = 0 }
foreach ($m in $map) {
    $k = (Normalize-Text $m.Kind).ToUpper()
    $c = Normalize-Text $m.Col
    if ($c -notmatch '^\d+$') { continue }
    if ($k -eq 'KENNO'  -and $idCols.KenNo  -eq 0) { $idCols.KenNo  = [int]$c }
    if ($k -eq 'KENYMD' -and $idCols.KenYmd -eq 0) { $idCols.KenYmd = [int]$c }
}
$split = Split-HeaderData $rows $idCols
$data = @($split.Data)

Write-Host ''
Write-Host '3. 人数' -ForegroundColor White
Line
Write-Host ("  データ行(人数): {0}" -f $data.Count)
if ($data.Count -eq 0) { Write-Host '  ★ データがありません。' -ForegroundColor Red; return }

# ---------------------------------------------------------------- 4. 受付番号・氏名
Write-Host ''
Write-Host '4. 受付番号と氏名' -ForegroundColor White
Line

$colKenNo  = $idCols.KenNo
$colKanji  = 0
$colKana   = 0
$colYmd    = $idCols.KenYmd
foreach ($m in $map) {
    $k = (Normalize-Text $m.Kind).ToUpper()
    $c = Normalize-Text $m.Col
    if ($c -notmatch '^\d+$') { continue }
    if ($k -eq 'NAMEKANJI' -and $colKanji -eq 0) { $colKanji = [int]$c }
    if ($k -eq 'NAMEKANA'  -and $colKana  -eq 0) { $colKana  = [int]$c }
}

if ($colKenNo -eq 0) {
    if ($colKanji -gt 0 -or $colKana -gt 0) {
        Write-Host '  受付番号の列 : ありません → 氏名で健診ナビの受診者を探して書き込みます。' -ForegroundColor Green
    }
    else {
        Write-Host '  ★ 受付番号の列も氏名の列もありません。誰の結果か特定できません。' -ForegroundColor Red
    }
}
else {
    $nos = @()
    $blank = 0
    foreach ($r in $data) {
        $v = Normalize-KenNo (Get-Field $r $colKenNo)
        if ($v -eq '') { $blank++ } else { $nos += $v }
    }
    $nums = @($nos | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
    Write-Host ("  受付番号の列 : {0} 列目" -f $colKenNo)
    Write-Host ("  入っている人 : {0} 人 (空欄 {1} 人)" -f $nos.Count, $blank)
    if ($nums.Count -gt 0) {
        Write-Host ("  範囲         : {0} 〜 {1}" -f ($nums | Measure-Object -Minimum).Minimum, ($nums | Measure-Object -Maximum).Maximum)
    }
    $script:dupCount = 0
    $dups = @($nos | Group-Object | Where-Object { $_.Count -gt 1 })
    $script:dupCount = $dups.Count
    if ($dups.Count -gt 0) {
        Write-Host ("  ★ 同じ受付番号が複数あります: {0}" -f (($dups | ForEach-Object { "$($_.Name)($($_.Count)件)" }) -join ' ')) -ForegroundColor Red
        Write-Host '     このまま取り込むと、別人の結果が上書きされます。'
    } else {
        Write-Host '  ・受付番号の重複はありません。' -ForegroundColor Green
    }
    $nonNum = @($nos | Where-Object { $_ -notmatch '^\d+$' })
    if ($nonNum.Count -gt 0) {
        Write-Host ("  ★ 数字でない受付番号: {0}" -f (($nonNum | Select-Object -First 5) -join ' ')) -ForegroundColor Red
    }
}

if ($colKanji -gt 0 -or $colKana -gt 0) {
    Write-Host ("  氏名の列     : 漢字 {0} 列目 / カナ {1} 列目" -f $colKanji, $colKana) -ForegroundColor Green
    $nameBlank = 0
    foreach ($r in $data) {
        $a = if ($colKanji -gt 0) { Normalize-Text (Get-Field $r $colKanji) } else { '' }
        $b = if ($colKana  -gt 0) { Normalize-Text (Get-Field $r $colKana)  } else { '' }
        if ($a -eq '' -and $b -eq '') { $nameBlank++ }
    }
    if ($nameBlank -gt 0) { Write-Host ("  ★ 氏名が空の人が {0} 人います。" -f $nameBlank) -ForegroundColor Yellow }
    else { Write-Host '  ・全員に氏名が入っています(取込時に照合できます)。' -ForegroundColor Green }
}
else {
    Write-Host '  氏名の列     : ありません' -ForegroundColor Yellow
    Write-Host '     取込時に氏名の突き合わせができません。受付番号だけが頼りになります。'
}

if ($colYmd -gt 0) {
    $ymds = @($data | ForEach-Object { Normalize-Text (Get-Field $_ $colYmd) } | Where-Object { $_ -ne '' } | Sort-Object -Unique)
    Write-Host ("  受診日の列   : {0} 列目 / 中身: {1}" -f $colYmd, (($ymds | Select-Object -First 5) -join ' '))
    if ($ymds.Count -gt 1) { Write-Host ("  ★ 受診日が {0} 種類あります。" -f $ymds.Count) -ForegroundColor Yellow }
}
else {
    Write-Host '  受診日の列   : ありません(取込のとき画面に日付を入れてください)' -ForegroundColor Yellow
}

# ---------------------------------------------------------------- 5. 項目ごとの中身
Write-Host ''
Write-Host '5. 項目ごとの中身 (対応表の順)' -ForegroundColor White
Line
Write-Host ('  {0,4} {1,-22} {2,-10} {3,6} {4}' -f '列', '項目', '種類', '入力率', '見本')
Line 'DarkGray'

$usedCols = @{}
$emptyItems = @()
foreach ($m in $map) {
    $kind = (Normalize-Text $m.Kind).ToUpper()
    $c = Normalize-Text $m.Col
    if ($c -notmatch '^\d+$') { continue }
    $ci = [int]$c
    $usedCols[$ci] = $true
    $c2 = Normalize-Text $m.Col2
    if ($c2 -match '^\d+$') { $usedCols[[int]$c2] = $true }
    if ($kind -eq 'IGNORE') { continue }

    $filled = 0
    $samples = @()
    foreach ($r in $data) {
        $v = Normalize-Text (Get-Field $r $ci)
        if ($v -ne '') {
            $filled++
            if ($samples.Count -lt $Sample -and ($samples -notcontains $v)) { $samples += $v }
        }
    }
    $rate = if ($data.Count -gt 0) { [math]::Round(100.0 * $filled / $data.Count) } else { 0 }
    $label = Normalize-Text $m.Label
    if ($label.Length -gt 22) { $label = $label.Substring(0, 21) + '…' }

    $color = 'Gray'
    if ($filled -eq 0) { $color = 'DarkGray'; $emptyItems += ("{0}列 {1}" -f $ci, (Normalize-Text $m.Label)) }
    elseif ($rate -eq 100) { $color = 'Green' }

    Write-Host ('  {0,4} {1,-22} {2,-10} {3,5}% {4}' -f $ci, $label, $kind, $rate, ($samples -join ' / ')) -ForegroundColor $color
}

# ---------------------------------------------------------------- 6. 食い違い
Write-Host ''
Write-Host '6. 食い違いの確認' -ForegroundColor White
Line

$unmapped = @()
for ($i = 1; $i -le $maxCols; $i++) {
    if ($usedCols.ContainsKey($i)) { continue }
    $anyVal = $false
    foreach ($r in $data) { if ((Normalize-Text (Get-Field $r $i)) -ne '') { $anyVal = $true; break } }
    if ($anyVal) {
        $h = if ($i -le $head.Count) { Normalize-Text $head[$i - 1] } else { '' }
        $unmapped += ("{0}列 [{1}]" -f $i, $h)
    }
}
if ($unmapped.Count -gt 0) {
    Write-Host ("  ★ 対応表に無いのに値が入っている列が {0} 個あります:" -f $unmapped.Count) -ForegroundColor Yellow
    foreach ($u in $unmapped) { Write-Host ("      {0}" -f $u) -ForegroundColor Yellow }
    Write-Host '     この列の値は取り込まれません。取り込むべき項目なら対応表に足す必要があります。'
}
else {
    Write-Host '  ・値の入っている列は、すべて対応表に載っています。' -ForegroundColor Green
}

if ($emptyItems.Count -gt 0) {
    Write-Host ''
    Write-Host ("  ・対応表にあるが全員空の項目が {0} 個あります (実施していない検査なら正常):" -f $emptyItems.Count) -ForegroundColor DarkGray
    foreach ($e in ($emptyItems | Select-Object -First 20)) { Write-Host ("      {0}" -f $e) -ForegroundColor DarkGray }
    if ($emptyItems.Count -gt 20) { Write-Host ("      … ほか {0} 個" -f ($emptyItems.Count - 20)) -ForegroundColor DarkGray }
}

# ---------------------------------------------------------------- 7. 変換表に無い値
Write-Host ''
Write-Host '7. 値の確認 (取込のときエラーになるもの)' -ForegroundColor White
Line
$codeMap = $null
try { $codeMap = Load-CodeMap } catch { }
$valueMap = @{}
try { $valueMap = Load-ValueMap } catch { }

# 問診の選択肢は本来 健診ナビ のマスタで決まる。ここではDBに繋がないので、
# monshin_map.csv に出てくるラベルを手がかりにして「見慣れない値」を拾うだけにする。
$monshinLabels = @{}
try {
    foreach ($item in (Load-MonshinMap).Values) {
        foreach ($lab in $item.Values) { if ($lab -ne '') { $monshinLabels[$lab] = $true } }
    }
} catch { }

$bad = @()      # 確実にエラーになるもの
$suspect = @()  # 要確認 (最終判定は健診ナビのマスタ次第)

foreach ($m in $map) {
    $kind = (Normalize-Text $m.Kind).ToUpper()
    $c = Normalize-Text $m.Col
    if ($c -notmatch '^\d+$') { continue }
    $ci = [int]$c
    $label = Normalize-Text $m.Label
    $cmName = Normalize-Text $m.CodeMap

    $seen = @{}
    foreach ($r in $data) {
        $v = Normalize-Text (Get-Field $r $ci)
        if ($v -eq '' -or $seen.ContainsKey($v)) { continue }
        $seen[$v] = $true

        # 変換表(code_map.csv)を使う項目
        if ($cmName -ne '' -and $codeMap) {
            $conv = $null
            try { $conv = Convert-Code $m $v } catch { $conv = $null }
            if ($null -eq $conv) {
                $bad += ("{0}列 {1}: 「{2}」 (変換表 {3})" -f $ci, $label, $v, $cmName)
            }
            continue
        }

        # 尿・聴力は決まった書き方しか受け付けない
        if ($kind -eq 'NYOU' -or $kind -eq 'CHORYOKU') {
            $r2 = Convert-FormValue $valueMap $kind $v
            if ($r2.Status -eq 'CONVERR') {
                $bad += ("{0}列 {1}: 「{2}」 ({3}は決まった書き方だけです)" -f $ci, $label, $v, $kind)
            }
            continue
        }

        # 問診
        if ($kind -eq 'MONSHIN' -and $monshinLabels.Count -gt 0) {
            $v2 = $v
            if ($valueMap.ContainsKey('MONSHIN') -and $valueMap['MONSHIN'].ContainsKey($v)) { $v2 = $valueMap['MONSHIN'][$v] }
            if (-not $monshinLabels.ContainsKey($v2)) {
                $suspect += ("{0}列 {1}: 「{2}」" -f $ci, $label, $v)
            }
            continue
        }

        # 数値項目に数字でないものが混じっていないか
        if ($kind -eq 'VALUE' -and $v -notmatch '^[<>]?\s*-?[0-9]+(\.[0-9]+)?$') {
            $suspect += ("{0}列 {1}: 「{2}」 (数値のはず)" -f $ci, $label, $v)
        }
    }
}

if ($bad.Count -gt 0) {
    Write-Host ("  ★ 確実にエラーになる値が {0} 件あります:" -f $bad.Count) -ForegroundColor Red
    foreach ($b in ($bad | Select-Object -First 30)) { Write-Host ("      {0}" -f $b) -ForegroundColor Red }
    if ($bad.Count -gt 30) { Write-Host ("      … ほか {0} 件" -f ($bad.Count - 30)) -ForegroundColor Red }
    Write-Host '     担当者に連絡して、form\code_map.csv に足してもらってください。'
}
else {
    Write-Host '  ・変換表に無い値はありません。' -ForegroundColor Green
}

if ($suspect.Count -gt 0) {
    Write-Host ''
    Write-Host ("  ▲ 要確認が {0} 件あります (見慣れない書き方の値):" -f $suspect.Count) -ForegroundColor Yellow
    foreach ($s in ($suspect | Select-Object -First 30)) { Write-Host ("      {0}" -f $s) -ForegroundColor Yellow }
    if ($suspect.Count -gt 30) { Write-Host ("      … ほか {0} 件" -f ($suspect.Count - 30)) -ForegroundColor Yellow }
    Write-Host '     問診の選択肢は健診ナビのマスタで決まるので、ここでは確定できません。'
    Write-Host '     取込のプレビューで「変換表に無いコード」と出るかどうかで最終判断してください。'
}

# ---------------------------------------------------------------- まとめ
Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host 'まとめ' -ForegroundColor Cyan
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  レイアウト : {0}" -f (Split-Path $mapPath -Leaf))
Write-Host ("  人数       : {0} 人" -f $data.Count)
$ng = @()
if ($needMax -gt $maxCols) { $ng += 'ファイルの列が足りない (レイアウト違いの疑い)' }
if ($colKenNo -eq 0 -and $colKanji -le 0 -and $colKana -le 0) { $ng += '受付番号も氏名も無く、誰の結果か特定できない' }
if ($script:dupCount -gt 0){ $ng += '受付番号が重複している' }
if ($bad.Count -gt 0)      { $ng += ('変換表に無い値が {0} 件' -f $bad.Count) }
if ($ng.Count -eq 0) {
    if ($suspect.Count -gt 0) {
        Write-Host ("  ▲ 要確認が {0} 件あります。見てから取り込んでください。" -f $suspect.Count) -ForegroundColor Yellow
    }
    else {
        Write-Host '  ★ そのまま取り込めます。' -ForegroundColor Green
    }
}
else {
    Write-Host '  ★ 直してから取り込んでください:' -ForegroundColor Red
    foreach ($x in $ng) { Write-Host ("      ・{0}" -f $x) -ForegroundColor Red }
}
Write-Host ''
