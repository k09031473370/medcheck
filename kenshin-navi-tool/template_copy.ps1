<#
.SYNOPSIS
  結果報告書のテンプレート (例 301) をコピーし、文字を差し替えた新しい番号 (例 305) のテンプレートを足す (template_copy.ps1)

.DESCRIPTION
  例: 301 の裏面「全国の医療生協では、…」を「当院では、…」にした 305 を作り、健診ナビの結果報告書の画面から 305 を選んで印刷できるようにする。
  - 元のテンプレート (301 など) は読むだけで、一切書き換えない
  - 新しいファイルを 1 つ足すだけ。同じ番号のファイルが既にあれば何もしない (上書きしない)
  - 足す前に out\ に出来上がりを置くので、Excel で開いて確かめられる。y を押したときだけテンプレートのフォルダに置く
  - 元に戻すときは -Undo 305 (足したファイルだけを消す。消したものは out\ に控えを残す)
  例:
    template_copy.ps1                                   (301 → 305、全国の医療生協 → 当院)
    template_copy.ps1 -Src 303 -No 306
    template_copy.ps1 -Undo 305
#>
[CmdletBinding()]
param(
    [string]$Src = '301',
    [string]$No = '305',
    [string]$From = '全国の医療生協',
    [string]$To = '当院',
    [string]$Folder = '\\KNSV\KenshinNavi\11_健康診断結果報告書',
    [string]$Undo
)
$ErrorActionPreference = 'Stop'
trap { Write-Host (" エラー: {0}`n   {1} 行目: {2}" -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber, ([string]$_.InvocationInfo.Line).Trim()) -ForegroundColor Red; break }
$OutDir = Join-Path $PSScriptRoot 'out'
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
if (-not (Test-Path $Folder)) { throw "テンプレートのフォルダが見つかりません: $Folder" }

if ($Undo) {
    if ($Undo -notmatch '^\d{3}$') { throw '番号は 3 けたで (例 305)' }
    if ($Undo -in @('101','301','302','303')) { throw "$Undo は元からあるテンプレートなので消しません" }
    $hit = @(Get-ChildItem $Folder -Filter "$($Undo)_*.xlsx" -File)
    if ($hit.Count -eq 0) { Write-Host " $Undo のテンプレートはありません。"; return }
    foreach ($h in $hit) {
        if ($h.Name -notmatch '_当院\.xlsx$') { Write-Host " $($h.Name) はこの道具で足したものではない (名前が _当院 で終わらない) ので消しません。" -ForegroundColor Yellow; continue }
        $ans = Read-Host " $($h.Name) を消しますか？ (y で消す。控えを out\ に残します)"
        if ($ans -ne 'y') { Write-Host ' やめました。'; continue }
        Copy-Item $h.FullName (Join-Path $OutDir ("消した控え_" + $h.Name)) -Force
        Remove-Item $h.FullName
        Write-Host " 消しました: $($h.FullName)" -ForegroundColor Green
    }
    return
}

if ($Src -notmatch '^\d{3}$' -or $No -notmatch '^\d{3}$') { throw '番号は 3 けたで (例 301・305)' }
$srcFile = @(Get-ChildItem $Folder -Filter "$($Src)_*.xlsx" -File)
if ($srcFile.Count -ne 1) { throw "$Src のテンプレートが 1 つに決まりません ($($srcFile.Count) 個): $Folder" }
$srcFile = $srcFile[0]
$exist = @(Get-ChildItem $Folder -Filter "$($No)_*.xlsx" -File)
if ($exist.Count -gt 0) { throw "$No のテンプレートは既にあります: $($exist[0].Name)。別の番号にするか、-Undo $No で消してから作り直してください" }
$newName = $No + $srcFile.Name.Substring($Src.Length) -replace '\.xlsx$', '_当院.xlsx'

Write-Host ''
Write-Host (" {0} をコピーして「{1}」→「{2}」にしたテンプレート {3} を作る (元の {0} は変えない)" -f $Src, $From, $To, $No) -ForegroundColor Cyan
Write-Host "   元: $($srcFile.FullName)"
Write-Host "   新: $(Join-Path $Folder $newName)"

# 1. 元を out\ にコピー (読むだけ)
$work = Join-Path $OutDir ("作業_" + $srcFile.Name)
$fs = [System.IO.File]::Open($srcFile.FullName, 'Open', 'Read', 'ReadWrite')
try { $ws = [System.IO.File]::Create($work); try { $fs.CopyTo($ws) } finally { $ws.Dispose() } } finally { $fs.Dispose() }

# 2. text_swap で文字を差し替えた別ファイルを作る
$swap = Join-Path $PSScriptRoot 'text_swap.ps1'
if (-not (Test-Path $swap)) { throw "text_swap.ps1 が同じフォルダにありません: $PSScriptRoot" }
& $swap -In $work -From $From -To $To -Suffix '_差替' -OutDir $OutDir -Commit
$made = Join-Path $OutDir ([IO.Path]::GetFileNameWithoutExtension($work) + '_差替.xlsx')
Remove-Item $work -ErrorAction SilentlyContinue
if (-not (Test-Path $made)) { throw "文字の差し替えができませんでした (「$From」が見つからない など)。何もしていません" }
$preview = Join-Path $OutDir $newName
Move-Item $made $preview -Force
Write-Host ''
Write-Host " 出来上がり (確かめ用): $preview" -ForegroundColor Green
Write-Host '   Excel で開いて、文が変わっていること・ほかが 元と同じこと を確かめてください (開いたら保存せずに閉じる)。'

# 3. y のときだけテンプレートのフォルダに置く
$ans = Read-Host " 健診ナビのテンプレートのフォルダに $newName を置きますか？ (y で置く)"
if ($ans -ne 'y') { Write-Host ' やめました。テンプレートのフォルダには何も置いていません。'; return }
$dest = Join-Path $Folder $newName
if (Test-Path $dest) { throw "既にあります: $dest" }
Copy-Item $preview $dest
Write-Host " 置きました: $dest" -ForegroundColor Green
Write-Host " 健診ナビの「健康診断結果報告書」の画面を開き直すと、$No が一覧に出るはずです。出なければ -Undo $No で消してください。"
