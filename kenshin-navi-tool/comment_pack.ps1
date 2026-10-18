<#
.SYNOPSIS
  医師指示事項 (総合判定N(編集) 11001〜11020) の途中の空き枠を詰める (comment_pack.ps1)

.DESCRIPTION
  ツールで文を消した枠 (重複を消す・治療継続に揃える) が空のまま残り、その下に文があると、結果票の医師指示事項に空いた行が出る。
  文の並び順はそのままで、空き枠の分だけ上に詰める (文章と番号を一緒に動かす)。総合判定(編集) 11000 は触らない。
  ・プレビュー → y で書く。書く前に本人の T_KENSA を backup\ に保存し、変更を backup\kukuri_comment_pack_<日時>.csv に残す (kukuri_comment_undo.bat で戻せる)。
  例:
    comment_pack.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
#>
[CmdletBinding()]
param(
    [string]$Dates,
    [string]$UkeNo,
    [int]$UkeFrom = 0,
    [int]$UkeTo = 0,
    [switch]$Commit,
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Test-Locked'       'function Commit-Plan')
$BackupDir = Join-Path $PSScriptRoot 'backup'
if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }

function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
function Split-List([string]$s) { return @(($s -split '[,、\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }

$SLOTS = @(1..20 | ForEach-Object { '{0}' -f (11000 + $_) })
function Is-Blank([string]$v) { $t = $v.Trim(); return ($t -eq '' -or $t -eq '#') }

$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
$ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })
if ($dateList.Count -eq 0) { throw '-Dates に受診日を指定してください' }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 医師指示事項の途中の空き枠を詰める (文の順番はそのまま)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  受診日: {0}  受付番号 {1}〜{2}" -f ($dateList -join ', '), $UkeFrom, $UkeTo)

$conn = Open-Db
try {
    $dLit = "'" + ($dateList -join "','") + "'"
    $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit)"
    if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
    if ($UkeFrom -gt 0) { $where += " AND s.UKE_NO_KENSA >= $UkeFrom" }
    if ($UkeTo -gt 0)   { $where += " AND s.UKE_NO_KENSA <= $UkeTo" }
    $slotLit = "'" + ($SLOTS -join "','") + "'"
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, ISNULL(k.KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($slotLit)
WHERE $where
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@).Rows
    $People = [ordered]@{}
    foreach ($r in $rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $People.Contains($pk)) { $People[$pk] = @{ Uke = [string]$r.UKE; Ymd = [string]$r.YMD; Name = [string]$r.NAME; C = @{} } }
        $People[$pk].C[[string]$r.CD] = @{ V = [string]$r.V; VC = [string]$r.VC }
    }

    $plan = @()
    foreach ($pk in $People.Keys) {
        $p = $People[$pk]; $c = $p.C
        $have = @($SLOTS | Where-Object { $c.ContainsKey($_) })          # この人にある枠 (番号順)
        $texts = @($have | Where-Object { -not (Is-Blank $c[$_].V) } | ForEach-Object { $c[$_] })
        $gap = $false
        for ($i = 0; $i -lt $texts.Count; $i++) { if ((Is-Blank $c[$have[$i]].V)) { $gap = $true } }
        if (-not $gap) { continue }
        for ($i = 0; $i -lt $have.Count; $i++) {
            $sc = $have[$i]
            $newV = $(if ($i -lt $texts.Count) { $texts[$i].V } else { '' })
            $newC = $(if ($i -lt $texts.Count) { $texts[$i].VC } else { '' })
            if ($c[$sc].V -eq $newV -and $c[$sc].VC -eq $newC) { continue }
            if ((Is-Blank $c[$sc].V) -and $newV -eq '') { continue }
            $plan += New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = $p.Ymd; 受付番号 = $p.Uke; 氏名 = $p.Name; 枠 = $sc
                前 = $c[$sc].V; 前CD = $c[$sc].VC; 後 = $newV; 後CD = $newC }
        }
    }

    Write-Host ''
    Write-Host '=== 詰める人 ===' -ForegroundColor Cyan
    $plan | Select-Object 受診日, 受付番号, 氏名, 枠, @{n='前';e={ if ($_.前.Length -gt 30) { $_.前.Substring(0,30) + '…' } else { $_.前 } }}, @{n='後';e={ if ($_.後.Length -gt 30) { $_.後.Substring(0,30) + '…' } else { $_.後 } }} | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
    Line
    Write-Host ("詰める人: {0} 人 / 書き換える枠: {1}" -f @($plan | Group-Object PkSeq).Count, $plan.Count) -ForegroundColor Green
    Line
    if ($plan.Count -eq 0) { Write-Host '途中に空き枠がある人はいません。'; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。文は消さず、順番も変えずに上へ詰めるだけです (kukuri_comment_undo.bat で戻せます)。' -ForegroundColor Yellow
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    $undoFile = Join-Path $BackupDir ("kukuri_comment_pack_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($plan | Group-Object PkSeq)) {
        $pk = [decimal]$grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} ({1}) は健診ナビで編集中です ({2})" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $lock) -ForegroundColor Yellow; $skipped += $grp.Group[0].受付番号; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            # 1 人分は全部そろって書けるときだけ書く (途中で変わっていたら、その人はまるごとやめる)
            $changed = $false
            foreach ($t in $grp.Group) {
                $q = Invoke-DbQuery $conn "SELECT ISNULL(KEKKA,'') AS V FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd" @{ p = $pk; cd = $t.枠 } $tran
                if ($q.Rows.Count -ne 1 -or [string]$q.Rows[0].V -ne $t.前) { $changed = $true }
            }
            if ($changed) { $tran.Rollback(); Write-Host ("  [変わっていたので触らず] {0} ({1})" -f $grp.Group[0].氏名, $grp.Group[0].受付番号) -ForegroundColor Yellow; continue }
            foreach ($t in $grp.Group) {
                $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET KEKKA = @v, KEKKA_CD = @c WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ v = $t.後; c = $t.後CD; p = $pk; cd = $t.枠 }
                if ($n -ne 1) { throw ("UPDATE 影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.枠) }
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.枠; 前の文章 = $t.前; 前の番号 = $t.前CD; 後の文章 = $t.後; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [詰めた] {0} ({1}) : {2} 枠" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $grp.Count) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) { $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 前の文章, 前の番号, 後の文章, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} 枠を書き換えました。記録: {1}  (戻すときは kukuri_comment_undo.bat でこの記録を選ぶ)" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0} → 画面を閉じてもう一度実行" -f ($skipped -join ', ')) -ForegroundColor Red }
}
finally { $conn.Close() }
