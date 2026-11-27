<#
.SYNOPSIS
  受診票 OCR v5 の Excel を、表面 (問診・計測・自覚症状) と裏面 (既往歴・家族歴) まとめて健診ナビに入れる (jushinhyo_all.ps1)

.DESCRIPTION
  受診票一括取込.bat に Excel をドラッグ → 受診日を入れる (複数はカンマ区切り) → 一覧を見て y。
  中でやっていること (どちらも今までの道具をそのまま呼ぶだけ。新しい書き込みの作りは無い):
    1. 受診日ごとに form_import.ps1 のプレビュー (表面・自覚症状)。その日の人だけが対象になり、ほかの日の人は「見つかりません」で飛ぶ
    2. kiou_v5.ps1 の一覧 (既往歴・家族歴。全部の受診日をまとめて)
    3. y を押すと、受診日ごとに form_import.ps1 -Commit、続けて kiou_v5.ps1 -Commit
  ・1 項目でもエラーのある人は、表面を書かずに最後に一覧で出す (Excel を直してもう一度流せば入る。入った人は同じ値になるだけ)。
    -Force を付けると、エラーの項目だけ飛ばして残りを書く (取込画面の「エラーを飛ばす」と同じ)。
  ・くわしい内容は out\受診票一括_<日時>_*.txt に残る (氏名が入るので院内だけで使う)。
  ・元に戻すのは今までどおり: 表面は 元に戻す.bat、既往歴・家族歴は 受診票既往歴_元に戻す.bat。
#>
[CmdletBinding()]
param(
    [string]$File,
    [string]$Ymd,                 # 受診日。複数はカンマ区切り (例 2026/11/22,2026/11/23,2026/11/24)
    [switch]$Force,               # エラーの項目だけ飛ばして、その人の残りの項目を書く
    [switch]$SkipKiou             # 既往歴・家族歴を入れない (表面だけ)
)
$ErrorActionPreference = 'Stop'
$here   = $PSScriptRoot
$coreScript = Join-Path $here 'form_import.ps1'
$kiouScript = Join-Path $here 'kiou_v5.ps1'
foreach ($need in @($coreScript, $kiouScript, (Join-Path $here 'form\mapping_shibaura_ai_v5.csv'), (Join-Path $here 'form\code_map.csv'))) {
    if (-not (Test-Path $need)) { throw "必要なファイルがありません: $need" }
}
function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
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
if (-not $Ymd) { $Ymd = Read-Host '受診日 (複数はカンマ区切り。例 2026/11/22,2026/11/23,2026/11/24)' }
$dateList = @()
foreach ($part in ($Ymd -split '[,、\s]+')) {
    if ($part.Trim() -eq '') { continue }
    $dt = [datetime]::MinValue
    if (-not [datetime]::TryParse($part.Trim(), [ref]$dt)) { throw ("受診日「{0}」を解釈できません。2026/11/22 のように入力してください。" -f $part) }
    $ds = $dt.ToString('yyyy/MM/dd')
    if ($dateList -notcontains $ds) { $dateList += $ds }
}
if ($dateList.Count -eq 0) { throw '受診日が入っていません。' }
$outDir = Join-Path $here 'out'
if (-not (Test-Path $outDir)) { [void](New-Item -ItemType Directory -Path $outDir) }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 受診票の一括取込 (表面・自覚症状 → 既往歴・家族歴)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($File))
Write-Host ("  受診日  : {0}" -f ($dateList -join ', '))
Write-Host ''

# ---- 1. 表面・自覚症状のプレビュー (受診日ごと) ----
Write-Host '【1】表面・自覚症状の確認 (まだ書き込みません。受診日ごとに 1〜2 分かかります)' -ForegroundColor Cyan
$pre = @{}
foreach ($d in $dateList) {
    Write-Host ("  {0} を確認中..." -f $d) -ForegroundColor DarkGray
    $text = Run-Core @('-Csv', $File, '-KenYmd', $d)
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
$layout = $pre[$dateList[0]].Layout
if ($layout -ne '' -and $layout -notlike '*ai_v5*') {
    Write-Host ("  ★レイアウトが「{0}」と判別されました。受診票 OCR v5 の Excel (80 列) ではないかもしれません。" -f $layout) -ForegroundColor Red
}
# どの受診日にも見つからない受付番号 = 全部の日の「見つからない」に入っている番号
$nowhere = $null
foreach ($d in $dateList) {
    if ($null -eq $nowhere) { $nowhere = @($pre[$d].NotFound) } else { $nowhere = @($nowhere | Where-Object { $pre[$d].NotFound -contains $_ }) }
}
$nowhere = @($nowhere | Sort-Object { [int]$_ })
$totalPeople = 0; foreach ($d in $dateList) { $totalPeople += $pre[$d].People }
$errAll = @(); foreach ($d in $dateList) { foreach ($n in $pre[$d].Err) { $errAll += ("{0} ({1})" -f $n, $d.Substring(5)) } }
Line
Write-Host ("  表面: 合わせて {0} 人が対象。" -f $totalPeople) -ForegroundColor Green
if ($nowhere.Count -gt 0) {
    Write-Host ("  ★どの受診日にも見つからない受付番号 ({0} 人): {1}" -f $nowhere.Count, ($nowhere -join ', ')) -ForegroundColor Yellow
    Write-Host '    → 受診日が足りないか、健診ナビに受付番号が入っていません。この人たちは入りません。' -ForegroundColor Yellow
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

# ---- 3. 確認して書く ----
Write-Host ''
Write-Host '※ ここまで、まだ何も書き込んでいません。' -ForegroundColor Yellow
$ans = Read-Host '表面と既往歴・家族歴をまとめて書き込みますか (y = 書き込む / それ以外 = やめる)'
if (($ans -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }

Write-Host ''
Write-Host '【3】表面・自覚症状を書き込み中' -ForegroundColor Cyan
$wroteTotal = 0; $itemTotal = 0; $errAfter = @(); $lockAfter = @(); $failed = $false
foreach ($d in $dateList) {
    Write-Host ("  {0} を書き込み中..." -f $d) -ForegroundColor DarkGray
    $coreArgs = @('-Csv', $File, '-KenYmd', $d, '-Commit')
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

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host (" 終わりました。表面: {0} 人 / {1} 項目" -f $wroteTotal, $itemTotal) -ForegroundColor Green
if ($errAfter.Count -gt 0 -and -not $Force) { Write-Host (" ★エラーの項目があって表面を書かなかった人 ({0} 人): {1}" -f $errAfter.Count, ($errAfter -join ', ')) -ForegroundColor Yellow }
if ($lockAfter.Count -gt 0) { Write-Host (" ★健診ナビの画面で開かれていて書けなかった人 ({0} 人): {1} → 画面を閉じてから、もう一度流してください" -f $lockAfter.Count, ($lockAfter -join ', ')) -ForegroundColor Red }
if ($nowhere.Count -gt 0) { Write-Host (" ★どの受診日にも見つからなかった受付番号 ({0} 人): {1}" -f $nowhere.Count, ($nowhere -join ', ')) -ForegroundColor Yellow }
Write-Host (" くわしい内容: {0}\受診票一括_{1}_*.txt" -f $outDir, $stamp)
Write-Host ' このあと、健診ナビで自動判定をかけてください。' -ForegroundColor Yellow
Write-Host ('=' * 76) -ForegroundColor Cyan
