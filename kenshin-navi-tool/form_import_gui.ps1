<#
.SYNOPSIS
  健診ナビ11 検査結果一括取込ツール GUI版 (form_import_gui.ps1)

.DESCRIPTION
  form_import.ps1 (同じフォルダに必要) のGUIフロントエンド。
  - Excelフォーム(.xlsx)を直接指定可能 (ExcelのCOM機能で一時CSVに自動変換)
  - プレビュー / 書込 / 列確認(-Inspect) / 枠一覧(-DumpItems) / 所見マスタ(-DumpSyoken)
  起動方法: 右クリック→「PowerShellで実行」、または
    powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import_gui.ps1
  ショートカットを作る場合のリンク先:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\kenshin-navi\form_import_gui.ps1
#>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$coreScript = Join-Path $scriptDir 'form_import.ps1'
if (-not (Test-Path $coreScript)) {
    [void][System.Windows.Forms.MessageBox]::Show("form_import.ps1 が見つかりません:`n$coreScript", 'エラー', 'OK', 'Error')
    return
}

# ============================================================================
# コア呼び出し
# ============================================================================

function Quote([string]$s) { return '"' + ($s -replace '"', '\"') + '"' }

# form_import.ps1 を別プロセスで実行し、出力テキストを返す
function Run-Core([string[]]$coreArgs) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
    $argStr = '-NoProfile -ExecutionPolicy Bypass -File ' + (Quote $coreScript)
    foreach ($a in $coreArgs) {
        if ($a -like '-*') { $argStr += ' ' + $a } else { $argStr += ' ' + (Quote $a) }
    }
    $psi.Arguments = $argStr
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $enc = [System.Text.Encoding]::GetEncoding(932)
    $psi.StandardOutputEncoding = $enc
    $psi.StandardErrorEncoding = $enc
    $p = [System.Diagnostics.Process]::Start($psi)
    $oT = $p.StandardOutput.ReadToEndAsync()
    $eT = $p.StandardError.ReadToEndAsync()
    $p.WaitForExit()
    $out = $oT.Result
    $err = $eT.Result
    if ((-not [string]::IsNullOrWhiteSpace($err))) { $out += "`r`n[エラー出力]`r`n" + $err }
    return $out
}

# .xlsx/.xlsm を Excel COM で一時CSV(SJIS)に変換して、そのパスを返す
function Convert-ExcelToCsv([string]$xlsxPath) {
    $tmp = Join-Path $env:TEMP ('form_import_' + [System.IO.Path]::GetFileNameWithoutExtension($xlsxPath) + '.csv')
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($xlsxPath, 0, $true)   # 読み取り専用で開く
        if (Test-Path $tmp) { Remove-Item $tmp -Force }
        $wb.Worksheets.Item(1).SaveAs($tmp, 6)             # 6 = xlCSV (先頭シートのみ)
        return $tmp
    }
    finally {
        if ($wb) { $wb.Close($false) | Out-Null }
        if ($excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
    }
}

# ============================================================================
# GUI 構築
# ============================================================================

$form = New-Object System.Windows.Forms.Form
$form.Text = '健診ナビ 検査結果取込 (form_import GUI)'
$form.Size = New-Object System.Drawing.Size(960, 680)
$form.MinimumSize = New-Object System.Drawing.Size(760, 480)
$form.StartPosition = 'CenterScreen'

function New-Label([string]$text, [int]$x, [int]$y, [int]$w) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, $y)
    $l.Size = New-Object System.Drawing.Size($w, 20)
    return $l
}

# --- 1段目: ファイル選択 ---
$form.Controls.Add((New-Label 'フォーム (Excel/CSV):' 12 15 130))
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(145, 12)
$txtFile.Size = New-Object System.Drawing.Size(670, 24)
$txtFile.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '参照...'
$btnBrowse.Location = New-Object System.Drawing.Point(825, 10)
$btnBrowse.Size = New-Object System.Drawing.Size(105, 26)
$btnBrowse.Anchor = 'Top,Right'
$form.Controls.Add($btnBrowse)

# --- 2段目: 受付番号・受診日・オプション ---
$form.Controls.Add((New-Label '受付番号:' 12 48 70))
$txtOnly = New-Object System.Windows.Forms.TextBox
$txtOnly.Location = New-Object System.Drawing.Point(85, 45)
$txtOnly.Size = New-Object System.Drawing.Size(90, 24)
$form.Controls.Add($txtOnly)

$form.Controls.Add((New-Label '受診日:' 195 48 55))
$txtYmd = New-Object System.Windows.Forms.TextBox
$txtYmd.Location = New-Object System.Drawing.Point(252, 45)
$txtYmd.Size = New-Object System.Drawing.Size(110, 24)
$form.Controls.Add($txtYmd)
$form.Controls.Add((New-Label '(例 2026/07/02。空なら全員/CSVの日付列)' 368 48 280))

$chkUtf8 = New-Object System.Windows.Forms.CheckBox
$chkUtf8.Text = 'CSVはUTF-8'
$chkUtf8.Location = New-Object System.Drawing.Point(650, 45)
$chkUtf8.Size = New-Object System.Drawing.Size(110, 24)
$form.Controls.Add($chkUtf8)

$chkForce = New-Object System.Windows.Forms.CheckBox
$chkForce.Text = 'エラー行を飛ばして書込 (-Force)'
$chkForce.Location = New-Object System.Drawing.Point(760, 45)
$chkForce.Size = New-Object System.Drawing.Size(190, 24)
$chkForce.Anchor = 'Top,Right'
$form.Controls.Add($chkForce)

# --- 3段目: ボタン ---
$btnInspect = New-Object System.Windows.Forms.Button
$btnInspect.Text = '1. 列確認'
$btnInspect.Location = New-Object System.Drawing.Point(12, 80)
$btnInspect.Size = New-Object System.Drawing.Size(110, 32)
$form.Controls.Add($btnInspect)

$btnDump = New-Object System.Windows.Forms.Button
$btnDump.Text = '2. 枠一覧(DB)'
$btnDump.Location = New-Object System.Drawing.Point(128, 80)
$btnDump.Size = New-Object System.Drawing.Size(110, 32)
$form.Controls.Add($btnDump)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = '3. プレビュー'
$btnPreview.Location = New-Object System.Drawing.Point(244, 80)
$btnPreview.Size = New-Object System.Drawing.Size(120, 32)
$btnPreview.Font = New-Object System.Drawing.Font($btnPreview.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnPreview)

$btnCommit = New-Object System.Windows.Forms.Button
$btnCommit.Text = '4. 書込実行'
$btnCommit.Location = New-Object System.Drawing.Point(370, 80)
$btnCommit.Size = New-Object System.Drawing.Size(120, 32)
$btnCommit.ForeColor = [System.Drawing.Color]::Firebrick
$btnCommit.Font = New-Object System.Drawing.Font($btnCommit.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnCommit)

$form.Controls.Add((New-Label '所見マスタ:' 530 87 75))
$cmbSyoken = New-Object System.Windows.Forms.ComboBox
$cmbSyoken.DropDownStyle = 'DropDownList'
[void]$cmbSyoken.Items.AddRange(@('SHIN (診察)','GANTEI (眼底)','ZK011 (心電図)','ZK020 (胸部X線 部位)','ZK021 (胸部X線 所見)','ZK030 (胃部X線 部位)','ZK031 (胃部X線 所見)','ZK041 (腹部エコー)'))
$cmbSyoken.SelectedIndex = 0
$cmbSyoken.Location = New-Object System.Drawing.Point(605, 84)
$cmbSyoken.Size = New-Object System.Drawing.Size(190, 26)
$form.Controls.Add($cmbSyoken)

$btnSyoken = New-Object System.Windows.Forms.Button
$btnSyoken.Text = '表示'
$btnSyoken.Location = New-Object System.Drawing.Point(800, 80)
$btnSyoken.Size = New-Object System.Drawing.Size(60, 32)
$form.Controls.Add($btnSyoken)

# --- 出力欄 ---
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap = $false
$txtOut.Font = New-Object System.Drawing.Font('MS Gothic', 9)
$txtOut.Location = New-Object System.Drawing.Point(12, 122)
$txtOut.Size = New-Object System.Drawing.Size(918, 505)
$txtOut.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($txtOut)

# ============================================================================
# イベント
# ============================================================================

function Append-Out([string]$text) {
    $txtOut.AppendText($text.TrimEnd() + "`r`n`r`n")
    $txtOut.SelectionStart = $txtOut.Text.Length
    $txtOut.ScrollToCaret()
}

$allButtons = @($btnBrowse, $btnInspect, $btnDump, $btnPreview, $btnCommit, $btnSyoken)

function Invoke-Busy([scriptblock]$work) {
    foreach ($b in $allButtons) { $b.Enabled = $false }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    try { & $work }
    catch { Append-Out ("[エラー] " + $_.Exception.Message) }
    finally {
        foreach ($b in $allButtons) { $b.Enabled = $true }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# 入力ファイルを解決 (.xlsx/.xlsm なら一時CSVに変換)
function Resolve-CsvPath {
    $f = $txtFile.Text.Trim().Trim('"')
    if ($f -eq '') { throw 'フォームのファイルを選択してください。' }
    if (-not (Test-Path $f)) { throw "ファイルが見つかりません: $f" }
    $ext = [System.IO.Path]::GetExtension($f).ToLower()
    if ($ext -eq '.xlsx' -or $ext -eq '.xlsm' -or $ext -eq '.xls') {
        Append-Out '[変換] Excel → CSV に変換しています...'
        [System.Windows.Forms.Application]::DoEvents()
        $csv = Convert-ExcelToCsv $f
        $chkUtf8.Checked = $false   # Excel COM の CSV は SJIS
        return $csv
    }
    return $f
}

# 共通引数 (受付番号・受診日・エンコーディング)
function Get-CommonArgs {
    $a = @()
    if ($txtOnly.Text.Trim() -ne '') { $a += @('-Only', $txtOnly.Text.Trim()) }
    if ($txtYmd.Text.Trim()  -ne '') { $a += @('-KenYmd', $txtYmd.Text.Trim()) }
    if ($chkUtf8.Checked) { $a += @('-CsvEncoding', 'UTF8') }
    return ,$a
}

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel/CSV|*.xlsx;*.xlsm;*.xls;*.csv|すべて|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { $txtFile.Text = $dlg.FileName }
})

$btnInspect.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        Append-Out (Run-Core (@('-Csv', $csv, '-Inspect') + (Get-CommonArgs)))
    }
})

$btnDump.Add_Click({
    Invoke-Busy {
        if ($txtOnly.Text.Trim() -eq '' -or $txtYmd.Text.Trim() -eq '') {
            throw '枠一覧には受付番号と受診日の両方を入力してください。'
        }
        Append-Out (Run-Core @('-DumpItems', '-Only', $txtOnly.Text.Trim(), '-KenYmd', $txtYmd.Text.Trim()))
    }
})

$btnPreview.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        Append-Out (Run-Core (@('-Csv', $csv) + (Get-CommonArgs)))
    }
})

$btnCommit.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        $who = if ($txtOnly.Text.Trim() -ne '') { '受付番号 ' + $txtOnly.Text.Trim() + ' の1名' } else { 'CSVの全員' }
        $msg = "$who にDB書込を実行します。`r`n先にプレビューで内容を確認しましたか?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg, '書込の確認', 'YesNo', 'Warning', 'Button2')
        if ($r -ne 'Yes') { Append-Out '[中止] 書込をキャンセルしました。'; return }
        $a = @('-Csv', $csv, '-Commit') + (Get-CommonArgs)
        if ($chkForce.Checked) { $a += '-Force' }
        Append-Out (Run-Core $a)
    }
})

$btnSyoken.Add_Click({
    Invoke-Busy {
        $cd = ($cmbSyoken.SelectedItem -split ' ')[0]
        Append-Out (Run-Core @('-DumpSyoken', $cd))
    }
})

Append-Out @'
使い方:
 1. 「参照...」でフォーム(.xlsx または .csv)を選択
 2. 受付番号(例 4001)と受診日(例 2026/07/02)を入力
 3. 「1.列確認」で列の並びを確認 → form\mapping.csv を整備 (初回のみ)
 4. 「2.枠一覧(DB)」で KOMOKU_CD を確認 → mapping.csv に記入 (初回のみ)
 5. 「3.プレビュー」で書込内容を確認
 6. 問題なければ「4.書込実行」 → 健診ナビで「自動判定」を実行
'@

[void]$form.ShowDialog()
