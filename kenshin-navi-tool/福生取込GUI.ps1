<#
.SYNOPSIS
  福生商工会 当日の取込 (画面付き) (福生取込GUI.ps1)
  福生取込GUI.bat をダブルクリック → JSON を選ぶ (または窓にドラッグ) → [実行]

.DESCRIPTION
  中身は 福生取込.bat と同じ。既にある3本を順に呼ぶだけで、新しい書込処理は無い。
    福生取込.ps1 -CheckOnly   … 事前チェック (団体のコース・対応表)
    受付結果取込.ps1          … 受付番号 (プレビュー → 画面で「はい」→ 書込)
    予約取込Excel作成.ps1     … 予約が無い人の Excel → 健診ナビの予約データ取込へ (人がボタン2回)
    受付結果取込.ps1          … いま登録した人の受付番号
  画面は進み具合と、各ツールの出力をそのまま見せる。
#>
[CmdletBinding()]
param([string]$Json)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$dir = $PSScriptRoot
$ps  = (Get-Command powershell).Source
$enc932 = [System.Text.Encoding]::GetEncoding(932)

# ============================================================================
# 画面
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = '福生 当日取込'
$form.Size = New-Object System.Drawing.Size(980, 760)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Yu Gothic UI', 10)
$form.AllowDrop = $true

$lblJson = New-Object System.Windows.Forms.Label
$lblJson.Text = '受付アプリの JSON'; $lblJson.Location = New-Object System.Drawing.Point(16, 18); $lblJson.AutoSize = $true
$txtJson = New-Object System.Windows.Forms.TextBox
$txtJson.Location = New-Object System.Drawing.Point(150, 14); $txtJson.Size = New-Object System.Drawing.Size(660, 26)
$txtJson.Text = $Json
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '参照...'; $btnBrowse.Location = New-Object System.Drawing.Point(820, 12); $btnBrowse.Size = New-Object System.Drawing.Size(120, 30)
$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = 'JSON をこの窓にドラッグしてもよい'; $lblHint.ForeColor = [System.Drawing.Color]::Gray
$lblHint.Location = New-Object System.Drawing.Point(150, 44); $lblHint.AutoSize = $true

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '実行'; $btnRun.Location = New-Object System.Drawing.Point(16, 72); $btnRun.Size = New-Object System.Drawing.Size(160, 44)
$btnRun.Font = New-Object System.Drawing.Font('Yu Gothic UI', 12, [System.Drawing.FontStyle]::Bold)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(222, 235, 247)

$steps = @('1. 事前チェック (団体のコース・対応表)', '2. 受付番号を入れる (予約がある人)', '3. 予約が無い人の Excel → 健診ナビで登録', '4. 受付番号を入れる (いま登録した人)', '5. ★ 画面で対応する人')
$stepLabels = @()
for ($i = 0; $i -lt $steps.Count; $i++) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = '－  ' + $steps[$i]; $l.AutoSize = $true
    $l.Location = New-Object System.Drawing.Point(200, (70 + $i * 24))
    $form.Controls.Add($l); $stepLabels += $l
}

$log = New-Object System.Windows.Forms.RichTextBox
$log.Location = New-Object System.Drawing.Point(16, 200); $log.Size = New-Object System.Drawing.Size(930, 470)
$log.Font = New-Object System.Drawing.Font('MS Gothic', 10); $log.ReadOnly = $true; $log.WordWrap = $false
$log.BackColor = [System.Drawing.Color]::White
$log.Anchor = 'Top,Bottom,Left,Right'

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = 'ログをコピー'; $btnCopy.Location = New-Object System.Drawing.Point(16, 680); $btnCopy.Size = New-Object System.Drawing.Size(140, 32); $btnCopy.Anchor = 'Bottom,Left'
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = '閉じる'; $btnClose.Location = New-Object System.Drawing.Point(806, 680); $btnClose.Size = New-Object System.Drawing.Size(140, 32); $btnClose.Anchor = 'Bottom,Right'

$form.Controls.AddRange(@($lblJson, $txtJson, $btnBrowse, $lblHint, $btnRun, $log, $btnCopy, $btnClose))

# ============================================================================
# 道具
# ============================================================================
function Log([string]$t, [string]$color = 'Black') {
    $log.SelectionStart = $log.TextLength; $log.SelectionLength = 0
    $log.SelectionColor = [System.Drawing.Color]::FromName($color)
    $log.AppendText($t + "`r`n")
    $log.SelectionColor = $log.ForeColor
    $log.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}
function Step([int]$i, [string]$state, [string]$extra = '') {
    $mark = switch ($state) { 'run' { '▶' } 'ok' { '✓' } 'ng' { '✗' } 'skip' { '－' } default { '－' } }
    $stepLabels[$i].Text = "$mark  $($steps[$i])" + $(if ($extra) { "   $extra" } else { '' })
    $stepLabels[$i].ForeColor = switch ($state) { 'run' { [System.Drawing.Color]::DarkOrange } 'ok' { [System.Drawing.Color]::Green } 'ng' { [System.Drawing.Color]::Red } default { [System.Drawing.Color]::Black } }
    [System.Windows.Forms.Application]::DoEvents()
}
# 子の PowerShell を動かして、出力を1行ずつログに流す。戻り値 = @{ Code=終了コード; Lines=出力 }
function Run-Child([string]$script, [string[]]$argList) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $ps
    $quoted = @($argList | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } })
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $dir $script) + '" ' + ($quoted -join ' ')
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.RedirectStandardInput = $true
    $psi.StandardOutputEncoding = $enc932; $psi.StandardErrorEncoding = $enc932
    $psi.WorkingDirectory = $dir
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.StandardInput.Close()   # Read-Host が来ても止まらないように閉じる
    $lines = New-Object System.Collections.Generic.List[string]
    while (-not $proc.StandardOutput.EndOfStream) {
        $line = $proc.StandardOutput.ReadLine()
        $lines.Add($line)
        $color = 'Black'
        if ($line -match '★|エラー|失敗|見つかりません') { $color = 'DarkRed' }
        elseif ($line -match '^\[書込\]|読み直し確認|^\[出力\]|OK ') { $color = 'Green' }
        elseif ($line -match '^---|^===') { $color = 'Navy' }
        Log ('    ' + $line) $color
    }
    $err = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    if ($err.Trim() -ne '') { Log ('    ' + ($err.Trim() -replace "`r?`n", "`r`n    ")) 'DarkRed'; foreach ($e in ($err -split "`r?`n")) { $lines.Add($e) } }
    if ($proc.ExitCode -ne 0) { Log ("    ※ {0} が終了コード {1} で終わりました (上の赤い行が原因です)" -f $script, $proc.ExitCode) 'DarkRed' }
    Save-Log
    return @{ Code = $proc.ExitCode; Lines = $lines }
}
# ログをファイルに残す (excel_tools\log\福生取込GUI_日時.txt)
$script:LogFile = $null
function Save-Log {
    try {
        $ld = Join-Path $dir 'log'
        if (-not (Test-Path $ld)) { [void](New-Item -ItemType Directory -Path $ld) }
        if (-not $script:LogFile) { $script:LogFile = Join-Path $ld ('福生取込GUI_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.txt') }
        [System.IO.File]::WriteAllText($script:LogFile, $log.Text, [System.Text.Encoding]::UTF8)
    } catch { }
}
# 子ツールが落ちたときに、原因になりそうな行だけ抜き出す
function Tail-Error($lines, [int]$n = 8) {
    $arr = @($lines)
    $hit = @($arr | Where-Object { $_ -match 'エラー|例外|Exception|失敗|ありません|できません|不正|throw|At line|発生場所' })
    if ($hit.Count -eq 0) { $hit = $arr }
    return (($hit | Select-Object -Last $n) -join "`r`n")
}
function Ask([string]$msg, [string]$title = '確認') {
    return [System.Windows.Forms.MessageBox]::Show($form, $msg, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
}
function Tell([string]$msg, [string]$title = '福生 当日取込') {
    [void][System.Windows.Forms.MessageBox]::Show($form, $msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# 受付番号: プレビュー → 人数を聞く → 書込
function Do-UkeNo([int]$stepNo, [string]$json) {
    Step $stepNo 'run'
    Log ''; Log ('--- プレビュー ---') 'Navy'
    $r = Run-Child '受付結果取込.ps1' @('-Json', $json, '-PreviewOnly')
    $n = -1
    foreach ($l in $r.Lines) { if ($l -match '\[更新予定\]\s+(\d+)\s*人') { $n = [int]$Matches[1] } }
    if ($r.Code -ne 0 -or $n -lt 0) {
        Step $stepNo 'ng' 'エラー'
        Tell ("受付結果取込 がエラーで止まりました。原因:`r`n`r`n" + (Tail-Error $r.Lines) + "`r`n`r`nログは " + $script:LogFile) 'エラー'
        return $false
    }
    if ($n -eq 0) { Step $stepNo 'ok' '更新なし'; return $true }
    $ans = Ask ("上のプレビューのとおり、{0} 人に受付番号を入れます。`r`n`r`nよければ「はい」" -f $n) '受付番号を入れる'
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { Step $stepNo 'skip' '中止'; Log '  中止しました。何も書いていません。' 'DarkOrange'; return $false }
    Log ''; Log ('--- 書込 ---') 'Navy'
    $r2 = Run-Child '受付結果取込.ps1' @('-Json', $json, '-Commit')
    $ok = @($r2.Lines | Where-Object { $_ -match '読み直し確認' }).Count -gt 0
    if ($ok) { Step $stepNo 'ok' ("{0} 人" -f $n) }
    else {
        Step $stepNo 'ng' '書込を確認できません'
        Tell ("書込が確認できませんでした。原因:`r`n`r`n" + (Tail-Error $r2.Lines) + "`r`n`r`n何も書いていないか、書いた分は backup に控えがあります。ログは " + $script:LogFile) 'エラー'
    }
    return $ok
}

# ============================================================================
# 実行
# ============================================================================
$btnBrowse.Add_Click({
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Filter = '受付アプリの JSON (*.json)|*.json|すべて (*.*)|*.*'
    if ($d.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $txtJson.Text = $d.FileName }
})
$form.Add_DragEnter({ param($s, $e) if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy } })
$form.Add_DragDrop({ param($s, $e) $f = @($e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)); if ($f.Count -gt 0) { $txtJson.Text = [string]$f[0] } })
$btnCopy.Add_Click({
    if (-not $log.Text) { return }
    # リモートデスクトップ越しだとクリップボードが一瞬使えないことがあるので、3回やり直す
    $done = $false
    for ($i = 0; $i -lt 3 -and -not $done; $i++) {
        try { [System.Windows.Forms.Clipboard]::SetDataObject($log.Text, $true, 5, 100); $done = $true }
        catch { Start-Sleep -Milliseconds 300 }
    }
    if ($done) { Tell 'ログをコピーしました。貼り付けできます。' }
    else {
        Save-Log
        Tell ("クリップボードが使えなかったので、ログをメモ帳で開きます。`r`n" + $script:LogFile)
        if ($script:LogFile -and (Test-Path $script:LogFile)) { Start-Process notepad.exe -ArgumentList "`"$($script:LogFile)`"" }
    }
})
$btnClose.Add_Click({ $form.Close() })

$btnRun.Add_Click({
    $json = $txtJson.Text.Trim().Trim('"')
    if (-not $json -or -not (Test-Path $json)) { Tell 'JSON を選んでください。'; return }
    if ([System.IO.Path]::GetExtension($json).ToLower() -ne '.json') { Tell 'Excel ではなく、受付アプリの JSON を選んでください。'; return }
    $btnRun.Enabled = $false; $log.Clear()
    for ($i = 0; $i -lt $steps.Count; $i++) { Step $i 'none' }
    try {
        # 今日の日付 = JSON の一番新しい受付日
        $doc = Get-Content -LiteralPath $json -Raw -Encoding UTF8 | ConvertFrom-Json
        $ymd = [string](@($doc.people | Where-Object { $_.actual -and $_.actual.checked_in_date } | ForEach-Object { [string]$_.actual.checked_in_date } | Sort-Object -Descending) | Select-Object -First 1)
        if (-not $ymd) { Tell 'JSON に受付済みの人がいません。'; return }
        $ymd = ($ymd -replace '-', '/')
        $recv = @($doc.people | Where-Object { $_.actual -and $null -ne $_.actual.reception_number -and (([string]$_.actual.checked_in_date) -replace '-', '/') -eq $ymd }).Count
        Log ("福生 当日取込  {0}   受付した人 {1} 人   ({2}  revision {3})" -f $ymd, $recv, (Split-Path $json -Leaf), $doc.revision) 'Navy'

        # ---- 1. 事前チェック ----
        Step 0 'run'
        $r = Run-Child '福生取込.ps1' @('-Json', $json, '-Ymd', $ymd, '-CheckOnly')
        if ($r.Code -eq 0) { Step 0 'ok' }
        elseif ($r.Code -ne 2) {
            Step 0 'ng' 'エラー'
            Tell ("事前チェックがエラーで止まりました (DB に繋がらない等)。原因:`r`n`r`n" + (Tail-Error $r.Lines) + "`r`n`r`nログは " + $script:LogFile) 'エラー'
            return
        }
        else {
            Step 0 'ng' '要対応'
            $ans = Ask "事前チェックで問題があります (ログの ★ を見てください)。`r`n`r`n健診ナビで直してからやり直すのが確実です。`r`n直さずに進めると、その人は健診ナビの取込で赤になります (他の人は入ります)。`r`n`r`nそれでも進めますか？" '事前チェック'
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { Log '止めました。直してから [実行] をもう一度。' 'DarkOrange'; return }
        }

        # ---- 2. 受付番号 ----
        [void](Do-UkeNo 1 $json)

        # ---- 3. Excel → 健診ナビ ----
        Step 2 'run'
        Log ''; Log '--- 予約が無い人の Excel ---' 'Navy'
        $r = Run-Child '予約取込Excel作成.ps1' @('-Json', $json, '-Ymd', $ymd)
        $xlsx = $null
        foreach ($l in $r.Lines) { if ($l -match '\[出力\]\s+(.+?\.xlsx)') { $xlsx = $Matches[1].Trim() } }
        if ($r.Code -ne 0) {
            Step 2 'ng' 'エラー'; Step 3 'skip'
            Tell ("予約取込Excel作成 がエラーで止まりました。原因:`r`n`r`n" + (Tail-Error $r.Lines) + "`r`n`r`nログは " + $script:LogFile) 'エラー'
        }
        elseif (-not $xlsx) {
            Step 2 'ok' '新しく登録する人なし'; Step 3 'skip'
        }
        else {
            # 健診ナビの予約データ取込を開く (exe が form\yoyaku_import_app.txt にあれば)
            $appTxt = Join-Path $dir 'form\yoyaku_import_app.txt'; $opened = $false
            if (Test-Path $appTxt) {
                $exe = (Get-Content $appTxt -Encoding UTF8 | Where-Object { $_.Trim() -ne '' -and -not $_.Trim().StartsWith('#') } | Select-Object -First 1)
                if ($exe -and (Test-Path $exe.Trim())) { try { Start-Process -FilePath $exe.Trim(); $opened = $true } catch { } }
            }
            try { Start-Process explorer.exe -ArgumentList "/select,`"$xlsx`"" } catch { }
            Step 2 'run' '健診ナビで登録してください'
            $msg = "Excel ができました:`r`n$xlsx`r`n`r`n" +
                   $(if ($opened) { "健診ナビの「予約データ取込」を開きました。`r`n" } else { "健診ナビのメニューから「予約データ取込」を開いてください。`r`n" }) +
                   "`r`n1. エクスプローラーで選ばれている Excel を、予約データ取込にドロップ`r`n" +
                   "2. 赤いセルが無ければ 「個人マスタ登録」 → 「予約登録」`r`n" +
                   "   (赤があれば、その行を Excel から消して保存し、ドロップし直す)`r`n" +
                   "3. 「予約登録中…」が終わるまで待つ`r`n`r`n終わったら OK を押してください。"
            Tell $msg '健診ナビで登録'
            Step 2 'ok'
            # ---- 4. 受付番号 (いま登録した人) ----
            [void](Do-UkeNo 3 $json)
        }

        # ---- 5. まとめ ----
        Step 4 'run'
        Log ''; Log '--- 画面で対応する人 ---' 'Navy'
        Log '  上のログで ★ が付いた人を、健診ナビの画面で対応してください。' 'DarkRed'
        if ($xlsx) {
            $memo = [System.IO.Path]::ChangeExtension($xlsx, '.確認一覧.txt')
            if (Test-Path $memo) { Get-Content $memo -Encoding Default | ForEach-Object { Log ('  ' + $_) 'DarkOrange' } }
        }
        Step 4 'ok'
        Log ''; Log '完了。' 'Green'; Save-Log; Log ("ログ: " + $script:LogFile) 'Gray'
        Tell "完了しました。`r`n`r`nログの ★ の人を健診ナビの画面で対応してください。`r`n「更新する人」が登録した人数より少なければ、その人は「見つかりません」に出ています。健診ナビで予約ができているか確認して、もう一度 [実行] してください。"
    }
    catch {
        Log ('エラー: ' + $_.Exception.Message) 'DarkRed'
        Log ('  場所: ' + $_.InvocationInfo.PositionMessage) 'DarkRed'
        Save-Log
        Tell ('エラーで止まりました:' + "`r`n`r`n" + $_.Exception.Message + "`r`n`r`nログは " + $script:LogFile) 'エラー'
    }
    finally { $btnRun.Enabled = $true }
})

[void]$form.ShowDialog()
