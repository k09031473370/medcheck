<#
.SYNOPSIS
  受診票 OCR v5 の Excel を、表面 (問診・計測・自覚症状) と裏面 (既往歴・家族歴) まとめて健診ナビに入れる (jushinhyo_all.ps1)

.DESCRIPTION
  受診票一括取込.bat に Excel をドラッグ → 受診日を入れる (複数はカンマ区切り) → 一覧を見て y。
  v5 (80 列) と v5.2・v5.3 (88 列・受診日の列つき) のどちらの Excel でも使える。
  ・Excel に「受診日」の列があり、全員に年月日がそろって入っていれば、受診日は聞かずにその日付を使う。
  ・「24」のように日だけの人・空の人がいるときは受診日を聞く。日だけの人は、入れた受診日のうち日が合うものに割り当てる。
  ・受診日が決まった人は、その受診日の予約だけを見る。決まらない人は、入れた受診日を順に見て受付番号で探す。
    そのとき同じ受付番号が 2 日以上にある (日ごとに番号を振り直す健診) と別の人に入る恐れがあるので、何も書かずに止める。
  中でやっていること (どちらも今までの道具をそのまま呼ぶだけ。新しい書き込みの作りは無い):
    1. 受診日ごとに form_import.ps1 のプレビュー (表面・自覚症状)。その日の人だけが対象になり、ほかの日の人は「見つかりません」で飛ぶ
    2. kiou_v5.ps1 の一覧 (既往歴・家族歴。全部の受診日をまとめて)
    3. y を押すと、受診日ごとに form_import.ps1 -Commit、続けて kiou_v5.ps1 -Commit
  ・1 項目でもエラーのある人は、表面を書かずに最後に一覧で出す (Excel を直してもう一度流せば入る。入った人は同じ値になるだけ)。
    -Force を付けると、エラーの項目だけ飛ばして残りを書く (取込画面の「エラーを飛ばす」と同じ)。
  ・くわしい内容は out\受診票一括_<日時>_*.txt に残る (氏名が入るので院内だけで使う)。
  ・元に戻すのは今までどおり: 表面は 元に戻す.bat、既往歴・家族歴は 受診票既往歴_元に戻す.bat。
  ・Excel に「住所」の列があれば (v5.3)、住所も一緒に入れるかを聞く (jusho_v53.ps1 を呼ぶ)。
    1 = 空の人だけ入れる / 2 = 受診票の住所で上書きする / それ以外 = 住所は入れない。元に戻すのは 受診票住所_元に戻す.bat。
#>
[CmdletBinding()]
param(
    [string]$File,
    [string]$Ymd,                 # 受診日。複数はカンマ区切り (例 2026/11/22,2026/11/23,2026/11/24)
    [switch]$Force,               # エラーの項目だけ飛ばして、その人の残りの項目を書く
    [switch]$SkipKiou,            # 既往歴・家族歴を入れない (表面だけ)
    [string]$JushoMode            # 住所の入れ方: empty = 空の人だけ / overwrite = 上書き / none = 入れない。指定しなければ聞く
)
$ErrorActionPreference = 'Stop'
$here   = $PSScriptRoot
$coreScript = Join-Path $here 'form_import.ps1'
$kiouScript = Join-Path $here 'kiou_v5.ps1'
$jushoScript = Join-Path $here 'jusho_v53.ps1'
foreach ($need in @($coreScript, $kiouScript, (Join-Path $here 'xlsx_read.ps1'), (Join-Path $here 'form\mapping_shibaura_ai_v5.csv'), (Join-Path $here 'form\mapping_shibaura_ai_v53.csv'), (Join-Path $here 'form\code_map.csv'))) {
    if (-not (Test-Path $need)) { throw "必要なファイルがありません: $need" }
}
# form_import.ps1 から Excel の読み込みの部品を借りる (kiou_v5.ps1 と同じやり方)
$src = Get-Content $coreScript -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text'     '$XlsxLib = Join-Path')
. (Join-Path $here 'xlsx_read.ps1')
Invoke-Expression (Get-Part 'function Test-XlsxEncrypted' 'function Detect-MappingPath')
function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
# Excel の「受診日」の欄: 年月日がそろっていれば yyyy/MM/dd、そうでなければ ''
function Get-FullYmd([string]$raw) {
    $v = ([string]$raw).Normalize([System.Text.NormalizationForm]::FormKC).Trim()
    if ($v -eq '') { return '' }
    $y = 0; $mo = 0; $dd = 0
    if     ($v -match '^(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})\s*日?$') { $y = [int]$Matches[1]; $mo = [int]$Matches[2]; $dd = [int]$Matches[3] }
    elseif ($v -match '^(\d{4})(\d{2})(\d{2})$')                                        { $y = [int]$Matches[1]; $mo = [int]$Matches[2]; $dd = [int]$Matches[3] }
    elseif ($v -match '^(?:R|令和)\s*(\d{1,2})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})\s*日?$') { $y = 2018 + [int]$Matches[1]; $mo = [int]$Matches[2]; $dd = [int]$Matches[3] }
    elseif ($v -match '^\d{5}(\.0+)?$') { $n = Normalize-Ymd $v; if ($n) { return $n } else { return '' } }   # Excel の日付の数
    else { return '' }
    try { return (New-Object datetime($y, $mo, $dd)).ToString('yyyy/MM/dd') } catch { return '' }
}
# 年が無い書き方: 「24」→ @{ M = 0; D = 24 }、「11/24」→ @{ M = 11; D = 24 }。読めなければ $null
function Get-PartYmd([string]$raw) {
    $v = ([string]$raw).Normalize([System.Text.NormalizationForm]::FormKC).Trim()
    if ($v -match '^(\d{1,2})\s*日?$') { $dd = [int]$Matches[1]; if ($dd -ge 1 -and $dd -le 31) { return @{ M = 0; D = $dd } } }
    if ($v -match '^(\d{1,2})\s*[/\-.月]\s*(\d{1,2})\s*日?$') { $mo = [int]$Matches[1]; $dd = [int]$Matches[2]; if ($mo -ge 1 -and $mo -le 12 -and $dd -ge 1 -and $dd -le 31) { return @{ M = $mo; D = $dd } } }
    return $null
}
function Quote([string]$s) { return '"' + ($s -replace '"', '\"') + '"' }
# form_import.ps1 を別プロセスで動かし、出力の文字を返す (取込画面 form_import_gui.ps1 と同じやり方)
function Run-Core([string[]]$coreArgs) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $argStr = '-NoProfile -ExecutionPolicy Bypass -File ' + (Quote $coreScript)
    foreach ($a in $coreArgs) { if ($a -like '-*') { $argStr += ' ' + $a } else { $argStr += ' ' + (Quote $a) } }
    $psi.Arguments = $argStr
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $enc = [System.Text.Encoding]::GetEncoding(932)
    $psi.StandardOutputEncoding = $enc
    $psi.StandardErrorEncoding = $enc
    $proc = [System.Diagnostics.Process]::Start($psi)
    $oT = $proc.StandardOutput.ReadToEndAsync()
    $eT = $proc.StandardError.ReadToEndAsync()
    $proc.WaitForExit()
    $text = $oT.Result
    if (-not [string]::IsNullOrWhiteSpace($eT.Result)) { $text += "`r`n[エラー出力]`r`n" + $eT.Result }
    return $text
}
# 「…受付番号 (3人): 83, 106, 113」の形の行から番号を取り出す
function Get-NoList([string]$text, [string]$label) {
    $found = New-Object 'System.Collections.Generic.List[string]'
    foreach ($m in [regex]::Matches($text, [regex]::Escape($label) + '\s*\((\d+)人\):\s*([0-9,\s]+)')) {
        foreach ($n in ($m.Groups[2].Value -split '[,\s]+')) { if ($n -match '^\d+$') { [void]$found.Add([string][int]$n) } }
    }
    return ,@($found | Select-Object -Unique)
}
function Get-Summary([string]$text) {
    $r = @{ People = -1; Items = -1; Wrote = -1; Layout = ''; NotFound = @(); Err = @(); NameNg = @(); Locked = @(); Ok = $false }
    $m = [regex]::Match($text, '【全体】(\d+) 人を確認 / 書込可能 (\d+) 件')
    if ($m.Success) { $r.People = [int]$m.Groups[1].Value; $r.Items = [int]$m.Groups[2].Value; $r.Ok = $true }
    $m = [regex]::Match($text, '【全体】(\d+) 人を処理 / 書込 (\d+) 人 / 書込項目 (\d+) 件')
    if ($m.Success) { $r.People = [int]$m.Groups[1].Value; $r.Wrote = [int]$m.Groups[2].Value; $r.Items = [int]$m.Groups[3].Value; $r.Ok = $true }
    $m = [regex]::Match($text, '\[自動判別\]\s*(\S+)')
    if ($m.Success) { $r.Layout = $m.Groups[1].Value }
    $r.NotFound = Get-NoList $text '受診者が見つからない受付番号'
    $r.Err      = Get-NoList $text 'エラーのある受付番号'
    $r.NameNg   = Get-NoList $text '氏名が一致せずスキップした受付番号'
    $r.Locked   = Get-NoList $text '★書込できなかった受付番号'
    return $r
}

# ---- 入力 ----
if (-not $File) { $File = Read-Host '受診票 OCR v5 の Excel をドラッグ＆ドロップして Enter' }
$File = ($File -replace '^"|"$', '').Trim()
if (-not (Test-Path $File)) { throw "ファイルが見つかりません: $File" }
# ---- Excel の受付番号と「受診日」の欄を読む ----
$rows = Read-FormCsv $File 'SJIS'
$head = $rows[0]
$colNo = 0; $colYmd = 0; $colAddr = 0
for ($ci = 1; $ci -le $head.Count; $ci++) {
    $hn = Normalize-Text (Get-Field $head $ci)
    if ($hn -eq '受付NO' -and $colNo -eq 0) { $colNo = $ci }
    if ($hn -eq '受診日' -and $colYmd -eq 0) { $colYmd = $ci }
    if ($hn -eq '住所' -and $colAddr -eq 0) { $colAddr = $ci }
}
if ($colNo -eq 0) { throw '見出しに「受付NO」がありません。受診票 OCR の Excel ではないようです。' }
$persons = @()
foreach ($f in $rows[1..($rows.Count - 1)]) {
    $u = Normalize-KenNo (Get-Field $f $colNo)
    if ($u -notmatch '^\d+$') { continue }
    $rawY = $(if ($colYmd -gt 0) { Normalize-Text (Get-Field $f $colYmd) } else { '' })
    $persons += @{ No = [string][int]$u; Raw = $rawY; Full = (Get-FullYmd $rawY); Part = (Get-PartYmd $rawY); Day = '' }
}
if ($persons.Count -eq 0) { throw '受付番号のある行がありません。' }
$dupNo = @($persons | Group-Object { $_.No } | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($dupNo.Count -gt 0) { throw ("Excel の中に同じ受付番号が 2 行以上あります: {0}。直してからやり直してください。" -f ($dupNo -join ', ')) }
$fullDates = @($persons | Where-Object { $_.Full -ne '' } | ForEach-Object { $_.Full } | Sort-Object -Unique)
$nFull = @($persons | Where-Object { $_.Full -ne '' }).Count
$nPart = @($persons | Where-Object { $_.Full -eq '' -and $null -ne $_.Part }).Count
$nNone = $persons.Count - $nFull - $nPart

Write-Host ''
Write-Host ("  Excel: {0} 人 (受付番号のある行)" -f $persons.Count)
if ($colYmd -gt 0) {
    Write-Host ("  Excel の「受診日」の欄: 年月日がそろっている {0} 人 / 日だけ・月日だけ {1} 人 / 空・読めない {2} 人" -f $nFull, $nPart, $nNone)
    if ($fullDates.Count -gt 0) { Write-Host ("  Excel から読めた受診日: {0}" -f ($fullDates -join ', ')) }
}
if (-not $Ymd) {
    if ($colYmd -gt 0 -and $nFull -eq $persons.Count) {
        Write-Host '  → 全員の受診日が Excel に入っているので、その日付を使います。' -ForegroundColor Green
        $Ymd = ($fullDates -join ',')
    } else {
        $Ymd = Read-Host '受診日 (複数はカンマ区切り。例 2026/11/22,2026/11/23,2026/11/24)'
    }
}
# ---- 住所 (v5.3 の Excel にだけ列がある) を一緒に入れるか ----
$jushoOn = $false
if ($colAddr -gt 0) {
    if (-not (Test-Path $jushoScript)) {
        Write-Host '  (Excel に住所の列がありますが、jusho_v53.ps1 が無いので住所は入れません)' -ForegroundColor DarkGray
    } else {
        if (-not $JushoMode) {
            Write-Host ''
            Write-Host '  この Excel には住所の列があります。住所も健診ナビに入れますか。'
            Write-Host '    1 = 空の人だけ入れる (健診ナビに住所が入っている人には触らない)'
            Write-Host '    2 = 受診票の住所で上書きする'
            Write-Host '    それ以外 (Enter だけ) = 住所は入れない'
            $ansJ = ([string](Read-Host '  番号')).Normalize([System.Text.NormalizationForm]::FormKC).Trim()
            if ($ansJ -eq '1') { $JushoMode = 'empty' } elseif ($ansJ -eq '2') { $JushoMode = 'overwrite' } else { $JushoMode = 'none' }
        }
        $JushoMode = $JushoMode.Trim().ToLower()
        if ($JushoMode -eq 'empty' -or $JushoMode -eq 'overwrite') { $jushoOn = $true }
        elseif ($JushoMode -ne 'none') { throw '-JushoMode は empty・overwrite・none のどれかです。' }
    }
}
$dateList = @()
foreach ($part in ($Ymd -split '[,、\s]+')) {
    if ($part.Trim() -eq '') { continue }
    $dt = [datetime]::MinValue
    if (-not [datetime]::TryParse($part.Trim(), [ref]$dt)) { throw ("受診日「{0}」を解釈できません。2026/11/22 のように入力してください。" -f $part) }
    $ds = $dt.ToString('yyyy/MM/dd')
    if ($dateList -notcontains $ds) { $dateList += $ds }
}
foreach ($fd in $fullDates) { if ($dateList -notcontains $fd) { $dateList += $fd } }   # Excel に年月日がある人の日は必ず含める
if ($dateList.Count -eq 0) { throw '受診日が入っていません。' }
$dateList = @($dateList | Sort-Object)
# 1 人ずつ受診日を決める: 年月日あり → その日 / 日だけ・月日だけ → 入れた受診日のうち合うものが 1 つならその日 / それ以外 → 決めない (受付番号で探す)
foreach ($ps in $persons) {
    if ($ps.Full -ne '') { $ps.Day = $ps.Full; continue }
    if ($null -ne $ps.Part) {
        $hit = @($dateList | Where-Object { $dx = [datetime]::ParseExact($_, 'yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture); ($dx.Day -eq $ps.Part.D) -and ($ps.Part.M -eq 0 -or $dx.Month -eq $ps.Part.M) })
        if ($hit.Count -eq 1) { $ps.Day = $hit[0] }
    }
}
$onlyOf = @{}
foreach ($d in $dateList) { $onlyOf[$d] = @($persons | Where-Object { $_.Day -eq $d -or $_.Day -eq '' } | ForEach-Object { $_.No }) }
$undecided = @($persons | Where-Object { $_.Day -eq '' } | ForEach-Object { $_.No })
$outDir = Join-Path $here 'out'
if (-not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir) }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host (' 受診票の一括取込 (表面・自覚症状 → 既往歴・家族歴' + $(if ($jushoOn) { ' → 住所' } else { '' }) + ')') -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($File))
Write-Host ("  受診日  : {0}" -f ($dateList -join ', '))
foreach ($d in $dateList) { Write-Host ("    {0}: Excel の受診日がこの日の人 {1} 人" -f $d, @($persons | Where-Object { $_.Day -eq $d }).Count) -ForegroundColor DarkGray }
if ($undecided.Count -gt 0) { Write-Host ("    受診日が決まらない人 {0} 人 → 上の受診日を順に見て、受付番号で探します" -f $undecided.Count) -ForegroundColor DarkGray }
if ($jushoOn) { Write-Host ("  住所    : {0}" -f $(if ($JushoMode -eq 'empty') { '空の人だけ入れる' } else { '受診票の住所で上書きする' })) }
Write-Host ''

# ---- 1. 表面・自覚症状のプレビュー (受診日ごと) ----
Write-Host '【1】表面・自覚症状の確認 (まだ書き込みません。受診日ごとに 1〜2 分かかります)' -ForegroundColor Cyan
$pre = @{}
foreach ($d in $dateList) {
    if ($onlyOf[$d].Count -eq 0) { $pre[$d] = @{ People = 0; Items = 0; Wrote = 0; Layout = ''; NotFound = @(); Err = @(); NameNg = @(); Locked = @(); Ok = $true }; Write-Host ("  {0}: 対象の人がいません" -f $d) -ForegroundColor DarkGray; continue }
    Write-Host ("  {0} を確認中..." -f $d) -ForegroundColor DarkGray
    $text = Run-Core @('-Csv', $File, '-KenYmd', $d, '-Only', ($onlyOf[$d] -join ','))
    $logFile = Join-Path $outDir ("受診票一括_{0}_{1}_確認.txt" -f $stamp, ($d -replace '/', ''))
    [System.IO.File]::WriteAllText($logFile, $text, (New-Object System.Text.UTF8Encoding($true)))
    $sm = Get-Summary $text
    if (-not $sm.Ok) {
        Write-Host ("  ★{0} の確認が最後まで進みませんでした。内容: {1}" -f $d, $logFile) -ForegroundColor Red
        $tail = @($text -split "`r?`n" | Where-Object { $_.Trim() -ne '' } | Select-Object -Last 8)
        foreach ($tl in $tail) { Write-Host ('    ' + $tl) -ForegroundColor Red }
        throw '表面の確認でエラーになったため、何も書き込まずに終わります。'
    }
    $pre[$d] = $sm
    Write-Host ("  {0}: {1} 人 / 書ける項目 {2} 件 / エラーのある人 {3} 人{4}" -f $d, $sm.People, $sm.Items, $sm.Err.Count,
        $(if ($sm.Err.Count -gt 0) { ' (受付番号 ' + ($sm.Err -join ', ') + ')' } else { '' }))
    if ($sm.NameNg.Count -gt 0) { Write-Host ("    ★氏名が一致しない受付番号: {0}" -f ($sm.NameNg -join ', ')) -ForegroundColor Red }
}
$layout = ''
foreach ($d in $dateList) { if ($pre[$d].Layout -ne '') { $layout = $pre[$d].Layout; break } }
if ($layout -ne '' -and $layout -notlike '*ai_v5*') {
    Write-Host ("  ★レイアウトが「{0}」と判別されました。受診票 OCR v5 の Excel (80 列) ではないかもしれません。" -f $layout) -ForegroundColor Red
}
# 1 人ずつ、どの受診日で見つかったかを数える (その日の対象で、「見つからない」に入っていなければ見つかった)
$foundOn = @{}
foreach ($ps in $persons) { $foundOn[$ps.No] = @() }
foreach ($d in $dateList) { foreach ($n in $onlyOf[$d]) { if ($pre[$d].NotFound -notcontains $n) { $foundOn[$n] += $d } } }
# 受診日が決まらない人が 2 日以上で見つかった = 日ごとに番号を振り直している。別の人に入る恐れがあるので止める
$twice = @($undecided | Where-Object { $foundOn[$_].Count -ge 2 } | Sort-Object { [int]$_ })
if ($twice.Count -gt 0) {
    Write-Host ''
    Write-Host ("★同じ受付番号の人が 2 日以上にいます ({0} 人): {1}" -f $twice.Count, (($twice | Select-Object -First 30) -join ', ')) -ForegroundColor Red
    Write-Host '  どの日の人か決められないので、何も書き込まずに終わります。' -ForegroundColor Red
    Write-Host '  → Excel の「受診日」の列に年月日を入れるか、Excel を受診日ごとに分けて、受診日を 1 日だけ指定してください。' -ForegroundColor Red
    return
}
# どこにも見つからない人。受診日が決まっていた人は、その日に番号が無かったということ
$nowhere = @($persons | Where-Object { $foundOn[$_.No].Count -eq 0 } | ForEach-Object { $(if ($_.Day -ne '') { "{0} ({1})" -f $_.No, $_.Day.Substring(5) } else { $_.No }) })
$totalPeople = 0; foreach ($d in $dateList) { $totalPeople += $pre[$d].People }
$errAll = @(); foreach ($d in $dateList) { foreach ($n in $pre[$d].Err) { $errAll += ("{0} ({1})" -f $n, $d.Substring(5)) } }
Line
Write-Host ("  表面: 合わせて {0} 人が対象。" -f $totalPeople) -ForegroundColor Green
if ($nowhere.Count -gt 0) {
    Write-Host ("  ★どの受診日にも見つからない受付番号 ({0} 人): {1}" -f $nowhere.Count, ($nowhere -join ', ')) -ForegroundColor Yellow
    Write-Host '    → 受診日が足りない・違う (かっこの中は Excel の受診日) か、健診ナビに受付番号が入っていません。この人たちは入りません。' -ForegroundColor Yellow
}
if ($errAll.Count -gt 0) {
    if ($Force) { Write-Host ("  ★エラーの項目がある人 ({0} 人): {1} → エラーの項目だけ飛ばして、残りを書きます (-Force)。" -f $errAll.Count, ($errAll -join ', ')) -ForegroundColor Yellow }
    else        { Write-Host ("  ★エラーの項目がある人 ({0} 人): {1} → この人の表面は書きません。確認の内容は out フォルダの「…_確認.txt」で受付番号を検索。" -f $errAll.Count, ($errAll -join ', ')) -ForegroundColor Yellow }
}
Line

# ---- 2. 既往歴・家族歴の一覧 ----
$ymdArg = ($dateList -join ',')
if (-not $SkipKiou) {
    Write-Host ''
    Write-Host '【2】既往歴・家族歴の確認 (まだ書き込みません)' -ForegroundColor Cyan
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $kiouScript -File $File -Ymd $ymdArg -PreviewOnly
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) { throw '既往歴・家族歴の確認でエラーになったため、何も書き込まずに終わります。' }
}
if ($jushoOn) {
    Write-Host ''
    Write-Host '【2b】住所の確認 (まだ書き込みません)' -ForegroundColor Cyan
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $jushoScript -File $File -Ymd $ymdArg -Mode $JushoMode -PreviewOnly
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) { throw '住所の確認でエラーになったため、何も書き込まずに終わります。' }
}

# ---- 3. 確認して書く ----
Write-Host ''
Write-Host '※ ここまで、まだ何も書き込んでいません。' -ForegroundColor Yellow
$ans = Read-Host ('表面と既往歴・家族歴' + $(if ($jushoOn) { '・住所' } else { '' }) + 'をまとめて書き込みますか (y = 書き込む / それ以外 = やめる)')
if (($ans -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }

Write-Host ''
Write-Host '【3】表面・自覚症状を書き込み中' -ForegroundColor Cyan
$wroteTotal = 0; $itemTotal = 0; $errAfter = @(); $lockAfter = @(); $failed = $false
foreach ($d in $dateList) {
    if ($onlyOf[$d].Count -eq 0 -or $pre[$d].People -eq 0) { continue }
    Write-Host ("  {0} を書き込み中..." -f $d) -ForegroundColor DarkGray
    $coreArgs = @('-Csv', $File, '-KenYmd', $d, '-Only', ($onlyOf[$d] -join ','), '-Commit')
    if ($Force) { $coreArgs += '-Force' }
    $text = Run-Core $coreArgs
    $logFile = Join-Path $outDir ("受診票一括_{0}_{1}_書込.txt" -f $stamp, ($d -replace '/', ''))
    [System.IO.File]::WriteAllText($logFile, $text, (New-Object System.Text.UTF8Encoding($true)))
    $sm = Get-Summary $text
    if (-not $sm.Ok -or $sm.Wrote -lt 0) {
        Write-Host ("  ★{0} の書き込みが最後まで進みませんでした。内容: {1}" -f $d, $logFile) -ForegroundColor Red
        $tail = @($text -split "`r?`n" | Where-Object { $_.Trim() -ne '' } | Select-Object -Last 8)
        foreach ($tl in $tail) { Write-Host ('    ' + $tl) -ForegroundColor Red }
        $failed = $true
        continue
    }
    $wroteTotal += $sm.Wrote; $itemTotal += $sm.Items
    foreach ($n in $sm.Err)    { $errAfter  += ("{0} ({1})" -f $n, $d.Substring(5)) }
    foreach ($n in $sm.Locked) { $lockAfter += ("{0} ({1})" -f $n, $d.Substring(5)) }
    Write-Host ("  {0}: {1} 人に書きました" -f $d, $sm.Wrote) -ForegroundColor Green
}
if ($failed) {
    Write-Host ''
    Write-Host '★表面の書き込みで止まった受診日があるため、既往歴・家族歴は書かずに終わります。上の内容を確認してください。' -ForegroundColor Red
    Write-Host '  (書けた受診日の分はそのまま入っています。原因を直してからもう一度流せば、残りが入ります。)' -ForegroundColor Red
    return
}

if (-not $SkipKiou) {
    Write-Host ''
    Write-Host '【4】既往歴・家族歴を書き込み中' -ForegroundColor Cyan
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $kiouScript -File $File -Ymd $ymdArg -Commit
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) { Write-Host '★既往歴・家族歴の書き込みでエラーになりました。上の内容を確認してください (表面は入っています)。' -ForegroundColor Red }
}

if ($jushoOn) {
    Write-Host ''
    Write-Host '【5】住所を書き込み中' -ForegroundColor Cyan
    & "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File $jushoScript -File $File -Ymd $ymdArg -Mode $JushoMode -Commit
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) { Write-Host '★住所の書き込みでエラーになりました。上の内容を確認してください (表面と既往歴・家族歴は入っています。住所は何も変わっていません)。' -ForegroundColor Red }
}

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host (" 終わりました。表面: {0} 人 / {1} 項目" -f $wroteTotal, $itemTotal) -ForegroundColor Green
if ($errAfter.Count -gt 0 -and -not $Force) { Write-Host (" ★エラーの項目があって表面を書かなかった人 ({0} 人): {1}" -f $errAfter.Count, ($errAfter -join ', ')) -ForegroundColor Yellow }
if ($lockAfter.Count -gt 0) { Write-Host (" ★健診ナビの画面で開かれていて書けなかった人 ({0} 人): {1} → 画面を閉じてから、もう一度流してください" -f $lockAfter.Count, ($lockAfter -join ', ')) -ForegroundColor Red }
if ($nowhere.Count -gt 0) { Write-Host (" ★どの受診日にも見つからなかった受付番号 ({0} 人): {1}" -f $nowhere.Count, ($nowhere -join ', ')) -ForegroundColor Yellow }
Write-Host (" くわしい内容: {0}\受診票一括_{1}_*.txt" -f $outDir, $stamp)
Write-Host ' このあと、健診ナビで自動判定をかけてください。' -ForegroundColor Yellow
Write-Host ('=' * 76) -ForegroundColor Cyan
