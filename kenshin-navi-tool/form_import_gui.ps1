<#
.SYNOPSIS
  健診ナビ11 検査結果一括取込ツール GUI版 (form_import_gui.ps1)

.DESCRIPTION
  form_import.ps1 (同じフォルダに必要) のGUIフロントエンド。
  - Excelフォーム(.xlsx)を直接指定可能 (ExcelのCOM機能で一時CSVに自動変換)
  - プレビュー / 書込 / 列確認(-Inspect) / 枠一覧(-DumpItems) / 所見マスタ(-DumpSyoken)
  起動方法: 右クリック→「PowerShellで実行」、または
    powershell -ExecutionPolicy Bypass -File C:\Users\User\Documents\excel_tools\form_import_gui.ps1
  ショートカットを作る場合のリンク先:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File C:\Users\User\Documents\excel_tools\form_import_gui.ps1
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
$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($txtOnly, "空欄=ファイル内の全員。複数指定はカンマ区切り (例: 4005,4009)")
$form.Controls.Add($txtOnly)

$form.Controls.Add((New-Label '受診日:' 195 48 55))
$txtYmd = New-Object System.Windows.Forms.TextBox
$txtYmd.Location = New-Object System.Drawing.Point(252, 45)
$txtYmd.Size = New-Object System.Drawing.Size(110, 24)
$form.Controls.Add($txtYmd)


$form.Controls.Add((New-Label 'レイアウト:' 380 48 70))
$cmbMap = New-Object System.Windows.Forms.ComboBox
$cmbMap.DropDownStyle = 'DropDownList'
$cmbMap.Location = New-Object System.Drawing.Point(448, 45)
$cmbMap.Size = New-Object System.Drawing.Size(210, 26)
[void]$cmbMap.Items.Add((New-Object PSObject -Property @{ Label = '自動判別 (おすすめ)'; File = 'auto' }))
foreach ($f in (Get-ChildItem (Join-Path $scriptDir 'form') -Filter 'mapping*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
    $label = $f.Name
    foreach ($ln in (Get-Content $f.FullName -TotalCount 5 -Encoding UTF8)) {
        if ($ln -match '^#\s*LABEL\s*=\s*(.+)$') { $label = $Matches[1].Trim(); break }
    }
    [void]$cmbMap.Items.Add((New-Object PSObject -Property @{ Label = $label; File = $f.Name }))
}
$cmbMap.DisplayMember = 'Label'
$cmbMap.ValueMember = 'File'
if ($cmbMap.Items.Count -gt 0) { $cmbMap.SelectedIndex = 0 }
$form.Controls.Add($cmbMap)

$chkNoHdr = New-Object System.Windows.Forms.CheckBox
$chkNoHdr.Text = '見出し行なし(手動)'
$chkNoHdr.Checked = $false
$chkNoHdr.Location = New-Object System.Drawing.Point(668, 45)
$chkNoHdr.Size = New-Object System.Drawing.Size(140, 24)
$form.Controls.Add($chkNoHdr)

$chkUtf8 = New-Object System.Windows.Forms.CheckBox
$chkUtf8.Text = 'CSVはUTF-8'
$chkUtf8.Location = New-Object System.Drawing.Point(812, 45)
$chkUtf8.Size = New-Object System.Drawing.Size(110, 24)
$form.Controls.Add($chkUtf8)

$chkForce = New-Object System.Windows.Forms.CheckBox
$chkForce.Text = 'エラーを飛ばす'
$chkForce.Location = New-Object System.Drawing.Point(834, 72)
$chkForce.Size = New-Object System.Drawing.Size(110, 24)
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

$btnUkeNo = New-Object System.Windows.Forms.Button
$btnUkeNo.Text = '受付番号を設定'
$btnUkeNo.Location = New-Object System.Drawing.Point(12, 118)
$btnUkeNo.Size = New-Object System.Drawing.Size(226, 30)
$form.Controls.Add($btnUkeNo)

$btnYoyaku = New-Object System.Windows.Forms.Button
$btnYoyaku.Text = '予約取込ファイルを作成'
$btnYoyaku.Location = New-Object System.Drawing.Point(244, 118)
$btnYoyaku.Size = New-Object System.Drawing.Size(180, 30)
$form.Controls.Add($btnYoyaku)

$form.Controls.Add((New-Label '← 受付番号の設定 / 予約取込用のExcelを作成' 430 124 320))

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

# --- 4段目: 名簿でしぼり込む (血液など、名簿外の人が混ざるファイル用) ---
$form.Controls.Add((New-Label '名簿でしぼり込む:' 12 156 115))
$txtRoster = New-Object System.Windows.Forms.TextBox
$txtRoster.Location = New-Object System.Drawing.Point(130, 153)
$txtRoster.Size = New-Object System.Drawing.Size(575, 24)
$txtRoster.Anchor = 'Top,Left,Right'
$tip.SetToolTip($txtRoster, "空欄=ファイル内の全員を取込。名簿(リアン等)を指定すると、その名簿に載っている人だけを取り込みます。")
$form.Controls.Add($txtRoster)

$btnRoster = New-Object System.Windows.Forms.Button
$btnRoster.Text = '名簿を選ぶ'
$btnRoster.Location = New-Object System.Drawing.Point(712, 151)
$btnRoster.Size = New-Object System.Drawing.Size(100, 26)
$btnRoster.Anchor = 'Top,Right'
$form.Controls.Add($btnRoster)

$btnRosterClear = New-Object System.Windows.Forms.Button
$btnRosterClear.Text = 'クリア'
$btnRosterClear.Location = New-Object System.Drawing.Point(818, 151)
$btnRosterClear.Size = New-Object System.Drawing.Size(70, 26)
$btnRosterClear.Anchor = 'Top,Right'
$form.Controls.Add($btnRosterClear)

# --- 出力欄 ---
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap = $false
$txtOut.Font = New-Object System.Drawing.Font('MS Gothic', 9)
$txtOut.Location = New-Object System.Drawing.Point(12, 188)
$txtOut.Size = New-Object System.Drawing.Size(918, 439)
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

$allButtons = @($btnBrowse, $btnInspect, $btnDump, $btnPreview, $btnCommit, $btnSyoken, $btnUkeNo, $btnYoyaku, $btnRoster, $btnRosterClear)

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
    if ($ext -eq '.xls') {
        throw "古い形式(.xls)は読めません。Excelで開いて .xlsx として保存し直してください。`n$f"
    }
    # .xlsx / .xlsm は form_import.ps1 が直接読む (Excelを使わないので固まらない)
    return $f
}

# 共通引数 (受付番号・受診日・エンコーディング)
function Get-CommonArgs {
    $a = @()
    if ($txtOnly.Text.Trim() -ne '') { $a += @('-Only', $txtOnly.Text.Trim()) }
    if ($txtYmd.Text.Trim()  -ne '') { $a += @('-KenYmd', $txtYmd.Text.Trim()) }
    if ($chkUtf8.Checked) { $a += @('-CsvEncoding', 'UTF8') }
    if ($chkNoHdr.Checked) { $a += '-NoHeader' }
    if ($cmbMap.SelectedItem) { $a += @('-Mapping', [string]$cmbMap.SelectedItem.File) }
    if ($txtRoster.Text.Trim() -ne '') { $a += @('-Roster', $txtRoster.Text.Trim().Trim('"')) }
    return ,$a
}

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel/CSV|*.xlsx;*.xlsm;*.xls;*.csv|すべて|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { $txtFile.Text = $dlg.FileName }
})

$btnRoster.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel/CSV|*.xlsx;*.xlsm;*.xls;*.csv|すべて|*.*'
    $dlg.Title = '名簿ファイル(リアン等)を選んでください'
    if ($dlg.ShowDialog() -eq 'OK') { $txtRoster.Text = $dlg.FileName }
})

$btnRosterClear.Add_Click({ $txtRoster.Text = '' })

$btnInspect.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        Append-Out (Run-Core (@('-Csv', $csv, '-Inspect') + (Get-CommonArgs)))
    }
})

$btnDump.Add_Click({
    Invoke-Busy {
        # ファイルが選ばれていれば、受付番号・受診日が空欄でもファイルから拾う
        $a = @('-DumpItems')
        $f = $txtFile.Text.Trim().Trim('"')
        if ($f -ne '' -and (Test-Path $f)) {
            $a += @('-Csv', (Resolve-CsvPath))
            if ($chkUtf8.Checked)  { $a += @('-CsvEncoding', 'UTF8') }
            if ($chkNoHdr.Checked) { $a += '-NoHeader' }
            if ($cmbMap.SelectedItem) { $a += @('-Mapping', [string]$cmbMap.SelectedItem.File) }
        }
        elseif ($txtOnly.Text.Trim() -eq '' -or $txtYmd.Text.Trim() -eq '') {
            throw 'ファイルを選ぶか、受付番号と受診日の両方を入力してください。'
        }
        if ($txtOnly.Text.Trim() -ne '') { $a += @('-Only', $txtOnly.Text.Trim()) }
        if ($txtYmd.Text.Trim()  -ne '') { $a += @('-KenYmd', $txtYmd.Text.Trim()) }
        Append-Out (Run-Core $a)
    }
})

# 受診日欄に入力があり、かつファイルにも日付の列がある場合は、
# どちらを使うのかを必ず確認する (ファイルの日付を黙って無視しないため)
function Confirm-YmdOverride([string]$csv) {
    $typed = $txtYmd.Text.Trim()
    if ($typed -eq '') { return $true }

    $a = @('-Csv', $csv, '-ShowYmd', '-KenYmd', $typed)
    if ($chkUtf8.Checked)  { $a += @('-CsvEncoding', 'UTF8') }
    if ($chkNoHdr.Checked) { $a += '-NoHeader' }
    if ($cmbMap.SelectedItem) { $a += @('-Mapping', [string]$cmbMap.SelectedItem.File) }
    $out = Run-Core $a

    $fileYmd = ''; $inYmd = ''
    foreach ($ln in ($out -split "`r?`n")) {
        if ($ln -match '^FILEYMD=(.*)$')  { $fileYmd = $Matches[1].Trim() }
        if ($ln -match '^INPUTYMD=(.*)$') { $inYmd   = $Matches[1].Trim() }
    }

    if ($inYmd -eq '') {
        [void][System.Windows.Forms.MessageBox]::Show(
            "受診日「$typed」を解釈できません。`r`n2026/07/12 のように入力してください。",
            '受診日の書き方', 'OK', 'Error')
        return $false
    }
    # 日付の列が無いファイル(芝浦巡回AIデータ等)は、受診日欄が唯一の情報源なので確認不要
    if ($fileYmd -eq '') { return $true }
    if ($fileYmd -eq $inYmd) { return $true }

    $msg = "このファイルには受診日の列があります。`r`n`r`n" +
           "  ファイルの日付 : $fileYmd`r`n" +
           "  入力した受診日 : $inYmd`r`n`r`n" +
           "入力した $inYmd を使い、ファイルの日付は無視します。`r`n" +
           "よろしいですか?`r`n`r`n" +
           "（ファイルの日付 $fileYmd を使いたい場合は「いいえ」を押し、受診日欄を空欄にしてください）"
    $r = [System.Windows.Forms.MessageBox]::Show($msg, '受診日の確認', 'YesNo', 'Warning', 'Button2')
    if ($r -ne 'Yes') {
        Append-Out "[中止] 受診日の確認でキャンセルしました。受診日欄を空欄にすると、ファイルの日付 $fileYmd を使います。"
        return $false
    }
    Append-Out "[受診日] 入力された $inYmd を使います (ファイルの日付 $fileYmd は使いません)。"
    return $true
}

$btnPreview.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        if (-not (Confirm-YmdOverride $csv)) { return }
        Append-Out (Run-Core (@('-Csv', $csv) + (Get-CommonArgs)))
    }
})

$btnCommit.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        if (-not (Confirm-YmdOverride $csv)) { return }
        $who = if ($txtOnly.Text.Trim() -ne '') { '受付番号 ' + $txtOnly.Text.Trim() + ' の1名' } else { 'CSVの全員' }
        $msg = "$who にDB書込を実行します。`r`n先にプレビューで内容を確認しましたか?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg, '書込の確認', 'YesNo', 'Warning', 'Button2')
        if ($r -ne 'Yes') { Append-Out '[中止] 書込をキャンセルしました。'; return }
        $a = @('-Csv', $csv, '-Commit') + (Get-CommonArgs)
        if ($chkForce.Checked) { $a += '-Force' }
        Append-Out (Run-Core $a)
    }
})

$btnUkeNo.Add_Click({
    Invoke-Busy {
        $csv = Resolve-CsvPath
        if (-not (Confirm-YmdOverride $csv)) { return }
        $a = @('-Csv', $csv, '-SetUkeNo') + (Get-CommonArgs)
        Append-Out (Run-Core $a)
        $r = [System.Windows.Forms.MessageBox]::Show(
            "上のプレビューを確認しました。`r`n「OK」の人の受付番号を健診ナビへ設定しますか?",
            '受付番号の設定', 'YesNo', 'Warning', 'Button2')
        if ($r -ne 'Yes') { Append-Out '[中止] 受付番号の設定をキャンセルしました。'; return }
        Append-Out (Run-Core ($a + '-Commit'))
    }
})

$btnYoyaku.Add_Click({
    Invoke-Busy {
        $f = $txtFile.Text.Trim().Trim('"')
        if ($f -eq '') { throw 'フォームのファイルを選択してください。' }
        $yo = Join-Path $scriptDir 'yoyaku_export.ps1'
        if (-not (Test-Path $yo)) { throw "yoyaku_export.ps1 が見つかりません: $yo" }
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $arg = '-NoProfile -ExecutionPolicy Bypass -File ' + (Quote $yo) + ' -Csv ' + (Quote $f)
        if ($chkNoHdr.Checked) { $arg += ' -NoHeader' }
        $psi.Arguments = $arg
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.CreateNoWindow = $true
        $enc = [System.Text.Encoding]::GetEncoding(932)
        $psi.StandardOutputEncoding = $enc; $psi.StandardErrorEncoding = $enc
        $p = [System.Diagnostics.Process]::Start($psi)
        $oT = $p.StandardOutput.ReadToEndAsync(); $eT = $p.StandardError.ReadToEndAsync()
        $p.WaitForExit()
        $o = $oT.Result
        if (-not [string]::IsNullOrWhiteSpace($eT.Result)) { $o += "`r`n[エラー出力]`r`n" + $eT.Result }
        Append-Out $o
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
