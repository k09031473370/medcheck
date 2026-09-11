<#
  フォルダ内のExcel全部のパスワードを一括で外したコピーを作る (Excelのパスワード一括解除.ps1)
  受け取った側が使う。このbatと同じフォルダにある .xlsx が対象。
  ・元のファイルは触らない。「解除済み」フォルダに保存する
  ・Excelを裏で動かすので、Excelを閉じてから実行する
#>
[CmdletBinding()]
param([string]$Folder)
$ErrorActionPreference = 'Stop'
if (-not $Folder) { $Folder = $PSScriptRoot }
$Folder = ($Folder -replace '^"|"$', '').Trim().TrimEnd('\')
$OutDir = Join-Path $Folder '解除済み'
$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
if ($files.Count -eq 0) { Write-Host "このフォルダに .xlsx がありません: $Folder"; Read-Host 'Enterで閉じる'; return }

Write-Host ''
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host ' Excel のパスワードを一括で外す' -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host ("  対象: {0} ファイル" -f $files.Count)
Write-Host ("  保存先: {0}" -f $OutDir)
Write-Host ''
$s = Read-Host -Prompt 'パスワード (入力中は表示されません)' -AsSecureString
$b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
try { $pw = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }

if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
try { $xl = New-Object -ComObject Excel.Application } catch { throw 'Excelが起動できません。' }
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$done = 0; $ng = @()
try {
    $n = 0
    foreach ($fx in $files) {
        $n++
        $dst = Join-Path $OutDir $fx.Name
        $wb = $null
        try {
            # Open(FileName, UpdateLinks, ReadOnly, Format, Password)
            $wb = $xl.Workbooks.Open($fx.FullName, 0, $true, 5, $pw)
            $wb.SaveAs($dst, 51, '')      # パスワード無しで保存
            $wb.Close($false)
            $done++
            Write-Host ("  [{0}/{1}] {2}" -f $n, $files.Count, $fx.Name) -ForegroundColor DarkGray
        } catch {
            $ng += ("{0} : {1}" -f $fx.Name, $_.Exception.Message)
            try { if ($wb) { $wb.Close($false) } } catch {}
        }
    }
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$pw = $null
Write-Host ''
Write-Host ("解除して保存: {0} ファイル / 失敗: {1} 件" -f $done, $ng.Count) -ForegroundColor $(if ($ng.Count -gt 0) { 'Yellow' } else { 'Green' })
if ($ng.Count -gt 0) {
    $ng | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow }
    Write-Host '  ※ パスワードが違うと失敗します。' -ForegroundColor Yellow
}
Write-Host ''
Write-Host ("「解除済み」フォルダに、パスワード無しのファイルができました: {0}" -f $OutDir) -ForegroundColor Green
Read-Host 'Enterで閉じる'
