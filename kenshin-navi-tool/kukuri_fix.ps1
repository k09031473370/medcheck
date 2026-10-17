<#
.SYNOPSIS
  HbA1c・血小板・eGFR の判定を「くくり判定」と総合判定に反映する (kukuri_fix.ps1)

.DESCRIPTION
  健診ナビの自動判定をかけ直さずに、check98 と同じ計算で次の判定欄 (HANTEI_KIGO) を上げる。
    HbA1c 068689 → 【判定】糖代謝 600340
    血小板 069010 → 【判定】血液一般 600310
    eGFR  068676 → 【判定】腎機能 600220
    3 項目の最も重い判定 → 総合判定(編集) 11000
  ・項目の判定の方が重いときだけ、その欄を項目の判定に上げる (軽くする方向には触らない)。
  ・総合判定が付いていない人 (自動判定前) は対象外。
  ・必ずプレビューを出し、y を押したときだけ書く。書く前に本人の T_KENSA を backup\ に保存し、
    直した欄の元の値を backup\kukuri_fix_<日時>.csv に残す。-Undo でその値に戻せる。
  例:
    kukuri_fix.ps1 -Dates 2026/09/03 -UkeFrom 4631 -UkeTo 4831
    kukuri_fix.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
    kukuri_fix.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$Dates,        # 受診日 (カンマ区切り)
    [string]$UkeNo,        # 受付番号 (カンマ区切り・任意)
    [int]$UkeFrom = 0,     # 受付番号の範囲 (任意)
    [int]$UkeTo = 0,
    [switch]$Commit,       # 付けると確認なしで書く。付けなければプレビュー後に y/n を聞く
    [string]$Undo,         # backup\kukuri_fix_*.csv のパス、または ask
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

# form_import.ps1 から接続・バックアップまわりを借りる (同じ作法で扱うため)
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
# 判定の重さ: A=1 B=2 C=3 D=4 E=5 F=6 G=7 J=7 (空は 0)
function Rank([string]$h) {
    switch ((Normalize-Text $h).ToUpper()) { 'A' {1} 'B' {2} 'C' {3} 'D' {4} 'E' {5} 'F' {6} 'G' {7} 'J' {7} default {0} }
}
$ITEMS = @(
    @{ Src = '068689'; SrcName = 'HbA1c';  Dst = '600340'; DstName = '糖代謝' },
    @{ Src = '069010'; SrcName = '血小板'; Dst = '600310'; DstName = '血液一般' },
    @{ Src = '068676'; SrcName = 'eGFR';   Dst = '600220'; DstName = '腎機能' }
)
$SOGO = '11000'

function Write-Hantei($conn, $tran, $pk, [string]$cd, [string]$h) {
    $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET HANTEI_KIGO = @h WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ h = $h; p = $pk; cd = $cd }
    if ($n -ne 1) { throw ("UPDATE 影響行数が {0} でした (PK_SEQ {1} / {2})。ロールバックします。" -f $n, $pk, $cd) }
}

# ============================================================
# 元に戻す (-Undo)
# ============================================================
if ($Undo) {
    if ($Undo -eq 'ask') {
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'kukuri_fix_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\kukuri_fix_*.csv) がありません。'; return }
        for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
        $a = Read-Host '戻す記録の番号 (それ以外 = やめる)'
        if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt $files.Count) { Write-Host 'やめました。'; return }
        $Undo = $files[[int]$a - 1].FullName
    }
    if (-not (Test-Path $Undo)) { throw "記録ファイルがありません: $Undo" }
    $rows = @(Import-Csv -Path $Undo -Encoding UTF8)
    Write-Host ("この記録で直した判定 {0} か所を元の値に戻します (今の値が記録の「後」と違う欄は触りません): {1}" -f $rows.Count, $Undo) -ForegroundColor Yellow
    $a = Read-Host '戻しますか (y = 戻す / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    $conn = Open-Db
    try {
        $back = 0; $kept = 0
        foreach ($grp in ($rows | Group-Object PK_SEQ)) {
            $pk = [decimal]$grp.Name
            $lock = Test-Locked $conn $pk
            if ($lock) { Write-Host ("  [とばしました] PK_SEQ {0} は健診ナビで編集中です ({1})" -f $pk, $lock) -ForegroundColor Yellow; continue }
            [void](Backup-Kensa $conn $pk)
            $tran = $conn.BeginTransaction()
            try {
                Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
                foreach ($r in $grp.Group) {
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET HANTEI_KIGO = @old WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) = @new" `
                        @{ old = [string]$r.前; p = $pk; cd = [string]$r.KOMOKU_CD; new = [string]$r.後 }
                    if ($n -eq 1) { $back++ } else { $kept++ }
                }
                $tran.Commit()
            } catch { $tran.Rollback(); throw }
        }
        Write-Host ("[完了] 戻した欄 {0} / 値が変わっていたので触らなかった欄 {1}" -f $back, $kept) -ForegroundColor Green
    } finally { $conn.Close() }
    return
}

# ============================================================
# 直す
# ============================================================
$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
$ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })
if ($dateList.Count -eq 0) { throw '-Dates に受診日を指定してください (例 2026/09/03)' }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' HbA1c・血小板・eGFR の判定を くくり判定・総合判定 に反映する' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  受診日    : {0}" -f ($dateList -join ', '))
if ($ukeList.Count -gt 0) { Write-Host ("  受付番号  : {0}" -f ($ukeList -join ', ')) }
if ($UkeFrom -gt 0 -or $UkeTo -gt 0) { Write-Host ("  受付番号の範囲: {0}〜{1}" -f $UkeFrom, $UkeTo) }
Write-Host '  糖代謝 600340 ← HbA1c / 血液一般 600310 ← 血小板 / 腎機能 600220 ← eGFR / 総合判定 11000 ← 3 項目の最重' -ForegroundColor DarkGray
Write-Host '  ※ 重くする方向だけ。総合判定が付いていない人 (自動判定前) は対象外。' -ForegroundColor DarkGray
Write-Host ''

$conn = Open-Db
try {
    $dLit = "'" + ($dateList -join "','") + "'"
    $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit)"
    if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
    if ($UkeFrom -gt 0) { $where += " AND s.UKE_NO_KENSA >= $UkeFrom" }
    if ($UkeTo -gt 0)   { $where += " AND s.UKE_NO_KENSA <= $UkeTo" }
    $people = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE $where
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@).Rows
    if ($people.Count -eq 0) { Write-Host '対象の人がいません (受診日・受付番号を見直してください)。' -ForegroundColor Yellow; return }
    Write-Host ("  対象者: {0} 人" -f $people.Count)

    $pkLit = ($people | ForEach-Object { [string]$_.PK_SEQ }) -join ','
    $cdLit = "'" + ((($ITEMS | ForEach-Object { $_.Src }) + ($ITEMS | ForEach-Object { $_.Dst }) + @($SOGO)) -join "','") + "'"
    $cur = @{}   # pk -> cd -> @{ H = 判定; V = 値 }
    foreach ($r in (Invoke-DbQuery $conn "SELECT PK_SEQ, LTRIM(RTRIM(KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(KEKKA,''))) AS V FROM T_KENSA WHERE PK_SEQ IN ($pkLit) AND LTRIM(RTRIM(KOMOKU_CD)) IN ($cdLit)").Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $cur.ContainsKey($pk)) { $cur[$pk] = @{} }
        $cur[$pk][[string]$r.CD] = @{ H = [string]$r.H; V = [string]$r.V }
    }

    # ---- 計画 ----
    $plan = @(); $skipNoSogo = 0
    foreach ($p in $people) {
        $pk = [string]$p.PK_SEQ
        if (-not $cur.ContainsKey($pk) -or -not $cur[$pk].ContainsKey($SOGO) -or (Normalize-Text $cur[$pk][$SOGO].H) -eq '') { $skipNoSogo++; continue }
        $c = $cur[$pk]
        $maxR = 0; $maxH = ''
        foreach ($it in $ITEMS) {
            if (-not $c.ContainsKey($it.Src)) { continue }
            $h = (Normalize-Text $c[$it.Src].H).ToUpper(); $r = Rank $h
            if ($r -gt $maxR) { $maxR = $r; $maxH = $h }
            if ($r -eq 0 -or -not $c.ContainsKey($it.Dst)) { continue }
            $dH = (Normalize-Text $c[$it.Dst].H).ToUpper()
            if ($r -gt (Rank $dH)) {
                $plan += New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$p.YMD; 受付番号 = [string]$p.UKE; 氏名 = [string]$p.NAME
                    欄 = $it.DstName; KOMOKU_CD = $it.Dst; 前 = $dH; 後 = $h; 理由 = ("{0} {1} = {2}" -f $it.SrcName, $c[$it.Src].V, $h) }
            }
        }
        $sH = (Normalize-Text $c[$SOGO].H).ToUpper()
        if ($maxR -gt (Rank $sH)) {
            $plan += New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$p.YMD; 受付番号 = [string]$p.UKE; 氏名 = [string]$p.NAME
                欄 = '総合判定'; KOMOKU_CD = $SOGO; 前 = $sH; 後 = $maxH; 理由 = ("3 項目の最重 = {0}" -f $maxH) }
        }
    }
    if ($skipNoSogo -gt 0) { Write-Host ("  総合判定が付いていないので対象外: {0} 人 (自動判定前)" -f $skipNoSogo) -ForegroundColor DarkGray }

    Write-Host ''
    Write-Host '=== 直す欄 ===' -ForegroundColor Cyan
    $plan | Select-Object 受診日, 受付番号, 氏名, 欄, 前, 後, 理由 | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    Line
    Write-Host ("直す: {0} か所 ({1} 人)" -f $plan.Count, @($plan | Group-Object PkSeq).Count) -ForegroundColor Green
    Line
    if ($plan.Count -eq 0) { Write-Host '直すものがありません。'; return }
    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
        $a = Read-Host '上の欄を直しますか (y = 直す / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }

    # ---- 書く ----
    $undoFile = Join-Path $BackupDir ("kukuri_fix_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($plan | Group-Object PkSeq)) {
        $pk = [decimal]$grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} ({1}) は健診ナビで編集中です ({2})" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $lock) -ForegroundColor Yellow; $skipped += $grp.Group[0].受付番号; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            $lock2 = Test-Locked $conn $pk $tran
            if ($lock2) { throw "書込直前に健診ナビで開かれました ($lock2)" }
            foreach ($t in $grp.Group) {
                # プレビュー後に値が変わっていたら、その欄は触らない
                $now = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS H FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd" @{ p = $pk; cd = $t.KOMOKU_CD } $tran
                if ($now.Rows.Count -ne 1 -or ([string]$now.Rows[0].H).ToUpper() -ne $t.前) { Write-Host ("  [変わっていたので触らず] {0} {1}" -f $t.氏名, $t.欄) -ForegroundColor Yellow; continue }
                Write-Hantei $conn $tran $pk $t.KOMOKU_CD $t.後
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.KOMOKU_CD; 欄 = $t.欄; 前 = $t.前; 後 = $t.後; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [直した] {0} ({1}) : {2}" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, (($grp.Group | ForEach-Object { "{0} {1}→{2}" -f $_.欄, $_.前, $_.後 }) -join ' / ')) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) {
        $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 欄, 前, 後, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8
    }
    Write-Host ''
    Write-Host ("[完了] {0} か所を直しました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて直せなかった受付番号: {0} → 画面を閉じてもう一度実行" -f ($skipped -join ', ')) -ForegroundColor Red }
    Write-Host '   元に戻すときは kukuri_fix_undo.bat。' -ForegroundColor DarkGray
}
finally { $conn.Close() }
