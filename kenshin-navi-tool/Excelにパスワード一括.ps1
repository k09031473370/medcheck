<#
.SYNOPSIS
  フォルダ内のExcel全部に、同じ「開くパスワード」を一括で付けたコピーを作る (Excelにパスワード一括.ps1)

.DESCRIPTION
  Excelの「パスワードを使用して暗号化」(ファイルを開くときのパスワード) を、
  フォルダ内の .xlsx すべてに同じパスワードで付ける。
  Excelの暗号化は中身がAESで暗号化されるので、受け取る側はExcelだけで開ける (追加ソフト不要)。
  7-Zip などを入れたくないときの代わり。

  ・元のファイルは触らない。<元フォルダ>_PW に保存する
  ・パスワードは画面に出ないように2回入力して確認する
  ・Excelを裏で動かすので、Excelを閉じてから実行する
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Folder,
    [string]$OutDir
)
$ErrorActionPreference = 'Stop'
if (-not $Folder) { $Folder = Read-Host 'Excelの入ったフォルダをドラッグ＆ドロップして Enter' }
$Folder = ($Folder -replace '^"|"$', '').Trim().TrimEnd('\')
if (-not (Test-Path $Folder)) { throw "見つかりません: $Folder" }
if (-not $OutDir) { $OutDir = $Folder + '_PW' }

$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
if ($files.Count -eq 0) { throw "フォルダに .xlsx がありません: $Folder" }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' Excel にパスワードを一括で付ける' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  元  : {0}  ({1} ファイル)" -f $Folder, $files.Count)
Write-Host ("  出力: {0}" -f $OutDir)
Write-Host ''
Write-Host '  パスワードの決め方: 12文字以上、英大文字・小文字・数字・記号をまぜる' -ForegroundColor DarkGray
Write-Host '  ※ 入力中は画面に出ません。伝えるときは電話かSMSで (メールでは送らない)' -ForegroundColor DarkGray
Write-Host ''

function Read-Plain([string]$prompt) {
    $s = Read-Host -Prompt $prompt -AsSecureString
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}
$pw1 = Read-Plain 'パスワード'
$pw2 = Read-Plain 'もう一度'
if ($pw1 -ne $pw2) { throw '2回の入力が違います。やり直してください。' }
if ($pw1.Length -lt 8) { throw '短すぎます。8文字以上 (できれば12文字以上) にしてください。' }
if ($pw1.Length -gt 15) { throw 'Excelのパスワードは15文字までです。' }

if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }

$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
} catch { throw 'Excelが起動できません。このPCにExcelが入っているか確認してください。' }
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$done = 0; $ng = @()
try {
    $n = 0
    foreach ($fx in $files) {
        $n++
        $dst = Join-Path $OutDir $fx.Name
        try {
            $wb = $xl.Workbooks.Open($fx.FullName, 0, $true)     # 読み取りで開く
            # 51 = xlOpenXMLWorkbook (.xlsx)。3つ目が「開くときのパスワード」
            $wb.SaveAs($dst, 51, $pw1)
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
$pw1 = $null; $pw2 = $null

Write-Host ''
Write-Host ("パスワード付きで保存: {0} ファイル / 失敗: {1} 件" -f $done, $ng.Count) -ForegroundColor $(if ($ng.Count -gt 0) { 'Yellow' } else { 'Green' })
if ($ng.Count -gt 0) { $ng | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow } }
Write-Host ''
Write-Host ("[出力] {0}" -f $OutDir) -ForegroundColor Green
Write-Host '※ 出力フォルダのファイルを1つ開いて、パスワードを聞かれることを確認してください。' -ForegroundColor Yellow
Write-Host '※ 説明書(Word)は Word で「ファイル → 情報 → 文書の保護 → パスワードを使用して暗号化」で同じパスワードを。' -ForegroundColor Yellow
Write-Host '※ そのあと出力フォルダを右クリック →「ZIPファイルに圧縮」(Windows標準) でまとめてメール添付。ZIP自体のパスワードは不要です。' -ForegroundColor Yellow
