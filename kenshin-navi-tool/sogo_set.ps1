<#
.SYNOPSIS
  先生の判断で、指定した人の 総合判定(編集) 11000 を決めた文字に直す (sogo_set.ps1)
  -Set "4770=D>B,4778=D>C"  … 受付番号=今の文字>直す文字。今の文字が違う人は触らない。文章は 2026 年の決まった言葉 (B=ほぼ正常(心配なし) など)。
  自動判定の総合 10000 は触らない。プレビュー → y。書く前に本人の T_KENSA を backup\ に保存し、backup\kukuri_fix_sogo_<日時>.csv に残す (kukuri_fix_undo.bat で戻せる)。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Set,
    [Parameter(Mandatory = $true)][string]$Dates,
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) { $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i); return $src.Substring($i, $j - $i) }
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Test-Locked'       'function Commit-Plan')
$BackupDir = Join-Path $PSScriptRoot 'backup'
$CANON = @{ 'A' = '異常なし'; 'B' = 'ほぼ正常(心配なし)'; 'C' = '要注意'; 'D' = '要観察(要経過観察)'; 'E' = '要再検'; 'F' = '要精検'; 'G' = '要治療'; 'J' = '治療中' }
$want = [ordered]@{}
foreach ($x in ($Set -split ',')) { if ($x -match '^\s*(\d+)\s*=\s*([A-J])\s*>\s*([A-J])\s*$') { $want[$Matches[1]] = @{ Old = $Matches[2]; New = $Matches[3] } } else { throw "書き方が違います: $x (例 4770=D>B)" } }
$dl = "'" + ((($Dates -split ',') | ForEach-Object { $_.Trim() }) -join "','") + "'"
$conn = Open-Db
try {
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '11000'
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dl) AND s.UKE_NO_KENSA IN ($(@($want.Keys) -join ','))
"@).Rows
    $plan = @()
    foreach ($r in $rows) {
        $w = $want[[string]$r.UKE]; $h = ([string]$r.H).ToUpper()
        if ($h -ne $w.Old) { Write-Host ("  [とばす] {0}: 今の総合判定は {1} (指定は {2})" -f $r.UKE, $h, $w.Old) -ForegroundColor Yellow; continue }
        $plan += [PSCustomObject]@{ PkSeq = [string]$r.PK_SEQ; 受付番号 = [string]$r.UKE; 氏名 = [string]$r.NAME; 前 = $h; 後 = $w.New; 前の結果 = [string]$r.V; 後の結果 = $CANON[$w.New] }
    }
    Write-Host ''
    Write-Host ' 総合判定(編集) を先生の判断で直す' -ForegroundColor Cyan
    $plan | Select-Object 受付番号, 氏名, 前, 後, 前の結果, 後の結果 | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    if ($plan.Count -eq 0) { Write-Host '直すものがありません。'; return }
    $a = Read-Host '直しますか (y = 直す / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    $log = @()
    foreach ($t in $plan) {
        $pk = [decimal]$t.PkSeq
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} は健診ナビで開かれています" -f $t.受付番号) -ForegroundColor Yellow; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET HANTEI_KIGO = @h, KEKKA = @v WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = '11000' AND UPPER(LTRIM(RTRIM(ISNULL(HANTEI_KIGO,'')))) = @old" @{ h = $t.後; v = $t.後の結果; p = $pk; old = $t.前 }
            if ($n -ne 1) { throw ("{0} の総合判定が変わっていました" -f $t.受付番号) }
            $tran.Commit()
            $log += [PSCustomObject][ordered]@{ PK_SEQ = $t.PkSeq; KOMOKU_CD = '11000'; 欄 = '総合判定(編集)'; 前 = $t.前; 後 = $t.後; 前HL = ''; 後HL = ''; 前の結果 = $t.前の結果; 後の結果 = $t.後の結果; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
            Write-Host ("  [直した] {0} : {1} → {2}" -f $t.受付番号, $t.前, $t.後) -ForegroundColor Green
        } catch { $tran.Rollback(); Write-Host ("  [直さず] {0}: {1}" -f $t.受付番号, $_.Exception.Message) -ForegroundColor Red }
    }
    if ($log.Count -gt 0) { $f = Join-Path $BackupDir ("kukuri_fix_sogo_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')); $log | Export-Csv -Path $f -NoTypeInformation -Encoding UTF8; Write-Host ("[完了] {0} 人。記録: {1} (戻すときは kukuri_fix_undo.bat)" -f $log.Count, $f) -ForegroundColor Green }
    Write-Host '  ★直したら ishiya_snapshot_save.bat でもう一度保存すること。' -ForegroundColor Yellow
}
finally { $conn.Close() }
