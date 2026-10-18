<#
.SYNOPSIS
  先生の判断で、胸部X線を F (要精検) から D (要観察) に下げる (xray_d.ps1)

.DESCRIPTION
  2026/10/18 先生の判断: 指定した人の胸部X線は D 判定 (経過観察)。次の欄を F→D にする。
    ・胸部X線1 077010A の判定 / 【判定】胸部X線 600090 / 【協会判定】胸部X線指導区分 610120 (結果が 1 文字なら結果も)
    ・医師指示事項の「胸部Ｘ線検査にて異常所見を認めます。精密検査の必要があります。」を 経過観察の文に戻す
    ・総合判定(編集) 11000 が F で、胸部X線を D にするとほかに F 以上 (J を含む) のくくりが無い人は、くくりの最重の文字と文章に下げる
  プレビュー → y。書く前に本人の T_KENSA を backup\ に保存。判定は backup\kukuri_fix_xrayd_<日時>.csv (kukuri_fix_undo.bat)、
  文は backup\kukuri_comment_xrayd_<日時>.csv (kukuri_comment_undo.bat) で戻せる。
  例: xray_d.ps1 -UkeNo 4638,4655,4670,4683 -Dates 2026/09/03
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$UkeNo,
    [Parameter(Mandatory = $true)][string]$Dates,
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
function Split-List([string]$s) { return @(($s -split '[,、\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }

$TEXT_F = '胸部Ｘ線検査にて異常所見を認めます。精密検査の必要があります。'
$TEXT_D = '胸部Ｘ線検査にて異常所見を認めます。経過観察の必要があります。自覚症状があれば検査を受けて下さい。'
$XCDS = [ordered]@{ '077010A' = '胸部X線1 の判定'; '600090' = '【判定】胸部X線'; '610120' = '【協会判定】胸部X線指導区分' }
$SLOTS = @(1..20 | ForEach-Object { [string](11000 + $_) })
$RANK = @{ 'A' = 1; 'B' = 2; 'C' = 3; 'D' = 4; 'E' = 5; 'F' = 6; 'G' = 7; 'J' = 8 }
$CANON = @{ 'A' = '異常なし'; 'B' = 'ほぼ正常(心配なし)'; 'C' = '要注意'; 'D' = '要観察(要経過観察)'; 'E' = '要再検'; 'F' = '要精検'; 'G' = '要治療'; 'J' = '治療中' }

$ukeList = @(Split-List $UkeNo | ForEach-Object { [int]$_ })
$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
Write-Host ''
Write-Host ' 胸部X線を F → D に下げる (先生の判断)' -ForegroundColor Cyan
Write-Host ("  受診日 {0} / 受付番号 {1}" -f ($dateList -join ','), ($ukeList -join ','))

$conn = Open-Db
try {
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD,
       LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H, ISNULL(k.KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('$($dateList -join "','")') AND s.UKE_NO_KENSA IN ($($ukeList -join ','))
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','610120','11000') OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600___' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '110__')
"@).Rows
    $People = [ordered]@{}
    foreach ($r in $rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $People.Contains($pk)) { $People[$pk] = @{ Uke = [string]$r.UKE; Name = [string]$r.NAME; C = @{} } }
        $People[$pk].C[[string]$r.CD] = @{ H = ([string]$r.H).ToUpper(); V = [string]$r.V; VC = [string]$r.VC }
    }
    $judge = @(); $com = @()
    foreach ($pk in $People.Keys) {
        $p = $People[$pk]; $c = $p.C
        if (-not $c.ContainsKey('600090') -or $c['600090'].H -ne 'F') { Write-Host ("  [とばす] {0} は【判定】胸部X線が F ではありません ({1})" -f $p.Uke, $c['600090'].H) -ForegroundColor Yellow; continue }
        foreach ($cd in $XCDS.Keys) {
            if (-not $c.ContainsKey($cd) -or $c[$cd].H -ne 'F') { continue }
            $v0 = Normalize-Text $c[$cd].V
            $judge += [PSCustomObject]@{ PkSeq = $pk; 受付番号 = $p.Uke; 氏名 = $p.Name; KOMOKU_CD = $cd; 欄 = $XCDS[$cd]; 前 = 'F'; 後 = 'D'; 前V = $v0; V = $(if ($v0 -match '^[A-J]$') { 'D' } else { $null }) }
        }
        foreach ($sc in $SLOTS) {
            if ($c.ContainsKey($sc) -and (Normalize-Text $c[$sc].V) -eq $TEXT_F) {
                $com += [PSCustomObject]@{ PkSeq = $pk; 受付番号 = $p.Uke; 氏名 = $p.Name; 枠 = $sc; 前 = (Normalize-Text $c[$sc].V); 前CD = $c[$sc].VC; 後 = $TEXT_D }
            }
        }
        # 総合判定(編集): F で、胸部X線を D にしたあと F 以上のくくりが無ければ、くくりの最重に下げる
        if ($c.ContainsKey('11000') -and $c['11000'].H -eq 'F') {
            $max = 0; $maxH = ''
            foreach ($cd in @($c.Keys | Where-Object { $_ -match '^600\d{3}$' })) {
                $h = $(if ($cd -eq '600090') { 'D' } else { $c[$cd].H })
                if ($RANK.ContainsKey($h) -and $RANK[$h] -gt $max) { $max = $RANK[$h]; $maxH = $h }
            }
            if ($max -gt 0 -and $max -lt $RANK['F']) {
                $judge += [PSCustomObject]@{ PkSeq = $pk; 受付番号 = $p.Uke; 氏名 = $p.Name; KOMOKU_CD = '11000'; 欄 = '総合判定(編集)'; 前 = 'F'; 後 = $maxH; 前V = (Normalize-Text $c['11000'].V); V = $CANON[$maxH] }
            }
        }
    }
    Write-Host ''
    Write-Host '=== 判定 ===' -ForegroundColor Cyan
    $judge | Select-Object 受付番号, 氏名, 欄, 前, 後, @{n='結果';e={ if ($null -ne $_.V) { "{0}→{1}" -f $_.前V, $_.V } else { '' } }} | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    Write-Host '=== 医師指示事項 ===' -ForegroundColor Cyan
    $com | Select-Object 受付番号, 氏名, 枠, 前, 後 | Format-Table -AutoSize -Wrap | Out-String -Width 240 | Write-Host
    Write-Host ("判定 {0} か所 / 文 {1} か所 ({2} 人)" -f $judge.Count, $com.Count, @(@($judge) + @($com) | Group-Object PkSeq).Count) -ForegroundColor Green
    if ($judge.Count + $com.Count -eq 0) { Write-Host '直すものがありません。'; return }
    Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
    $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }

    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $undoJ = @(); $undoC = @()
    foreach ($grp in (@(@($judge) + @($com)) | Group-Object PkSeq)) {
        $pk = [decimal]$grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} は健診ナビで開かれています ({1})" -f $grp.Group[0].受付番号, $lock) -ForegroundColor Yellow; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            foreach ($t in @($judge | Where-Object { [string]$_.PkSeq -eq $grp.Name })) {
                $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET HANTEI_KIGO = @h WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd AND UPPER(LTRIM(RTRIM(ISNULL(HANTEI_KIGO,'')))) = @old" @{ h = $t.後; p = $pk; cd = $t.KOMOKU_CD; old = $t.前 }
                if ($n -ne 1) { throw ("{0} {1} が変わっていました。ロールバックします。" -f $t.受付番号, $t.欄) }
                if ($null -ne $t.V) { [void](Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET KEKKA = @v WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ v = $t.V; p = $pk; cd = $t.KOMOKU_CD }) }
                $undoJ += [PSCustomObject][ordered]@{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.KOMOKU_CD; 欄 = $t.欄; 前 = $t.前; 後 = $t.後; 前HL = ''; 後HL = ''
                    前の結果 = $(if ($null -ne $t.V) { $t.前V } else { '' }); 後の結果 = $(if ($null -ne $t.V) { $t.V } else { '' }); 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
            }
            foreach ($t in @($com | Where-Object { [string]$_.PkSeq -eq $grp.Name })) {
                $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET KEKKA = @v, KEKKA_CD = '' WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(KEKKA,''))) = @old" @{ v = $t.後; p = $pk; cd = $t.枠; old = $t.前 }
                if ($n -ne 1) { throw ("{0} {1} の文が変わっていました。ロールバックします。" -f $t.受付番号, $t.枠) }
                $undoC += [PSCustomObject][ordered]@{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.枠; 前の文章 = $t.前; 前の番号 = $t.前CD; 後の文章 = $t.後; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
            }
            $tran.Commit()
            Write-Host ("  [直した] {0}" -f $grp.Group[0].受付番号) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoJ.Count -gt 0) { $undoJ | Export-Csv -Path (Join-Path $BackupDir ("kukuri_fix_xrayd_{0}.csv" -f $ts)) -NoTypeInformation -Encoding UTF8 }
    if ($undoC.Count -gt 0) { $undoC | Export-Csv -Path (Join-Path $BackupDir ("kukuri_comment_xrayd_{0}.csv" -f $ts)) -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] 判定 {0} か所 / 文 {1} か所。戻すときは kukuri_fix_undo.bat と kukuri_comment_undo.bat で xrayd の記録を選ぶ" -f $undoJ.Count, $undoC.Count) -ForegroundColor Green
}
finally { $conn.Close() }
