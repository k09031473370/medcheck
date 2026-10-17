<#
.SYNOPSIS
  HbA1c・血小板・eGFR・中性脂肪・眼底 の判定を「くくり判定」と総合判定に反映する (kukuri_fix.ps1)

.DESCRIPTION
  健診ナビの自動判定をかけ直さずに、check98 と同じ計算で次の判定欄 (HANTEI_KIGO) を上げる。
    HbA1c 068689 → 【判定】糖代謝 600340
    血小板 069010 → 【判定】血液一般 600310
    eGFR  068676 → 【判定】腎機能 600220
    中性脂肪(未判別) R06-0002 → 【判定】脂質代謝 600320
    【判定】眼底 600060 (業者の判定を mapping_gantei_hantei_* で入れたもの) → そのまま総合判定の材料に
    5 項目の最も重い判定 → 総合判定(編集) 11000
  ・中性脂肪(未判別) には基準値が無く判定が付かない (自動判定は採血時間を見て 空腹時/随時中性脂肪 に写してから判定する)。
    自動判定のあとに採血時間を入れた 9/9・9/10 は写っていないので、このツールが空腹時の基準で判定と H/L 印を付ける
    (ユーザー決定: 9/9 以降は全員空腹)。空腹時/随時中性脂肪に判定が付いている人 (9/3 など) は触らない。
      29 以下 F(L) / 30〜149 A / 150〜299 B(H) / 300〜499 D(H) / 500〜700 F(H) / 701 以上 G(H)  … 健診ナビの 068702 の基準値と同じ
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
    @{ Src = '068676'; SrcName = 'eGFR';   Dst = '600220'; DstName = '腎機能' },
    @{ Src = 'R06-0002'; SrcName = '中性脂肪'; Dst = '600320'; DstName = '脂質代謝'; Calc = $true; Alt = @('068702','R06-0003') },
    @{ Src = '600060'; SrcName = '眼底';   Dst = '600060'; DstName = '眼底' }   # くくり自身。総合判定にだけ効く (眼底 F → 総合判定 F)
)
$SOGO = '11000'
# 中性脂肪の値 → 判定と H/L (空腹時の基準)
function TgJudge([string]$v) {
    $d = 0.0
    if (-not [double]::TryParse(($v -replace '[^0-9.\-]', ''), [ref]$d)) { return $null }
    if ($d -lt 30)  { return @{ H = 'F'; HL = 'L' } }
    if ($d -lt 150) { return @{ H = 'A'; HL = '' } }
    if ($d -lt 300) { return @{ H = 'B'; HL = 'H' } }
    if ($d -lt 500) { return @{ H = 'D'; HL = 'H' } }
    if ($d -le 700) { return @{ H = 'F'; HL = 'H' } }
    return @{ H = 'G'; HL = 'H' }
}

function Write-Hantei($conn, $tran, $pk, [string]$cd, [string]$h, $hl) {
    if ($null -ne $hl) {
        $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET HANTEI_KIGO = @h, HL = @hl WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ h = $h; hl = [string]$hl; p = $pk; cd = $cd }
    } else {
        $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET HANTEI_KIGO = @h WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ h = $h; p = $pk; cd = $cd }
    }
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
                    $hasHL = ($r.PSObject.Properties.Name -contains '後HL') -and ([string]$r.後HL) -ne ''
                    if ($hasHL) {
                        $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET HANTEI_KIGO = @old, HL = @oldhl WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) = @new" `
                            @{ old = [string]$r.前; oldhl = [string]$r.前HL; p = $pk; cd = [string]$r.KOMOKU_CD; new = [string]$r.後 }
                    } else {
                        $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET HANTEI_KIGO = @old WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) = @new" `
                            @{ old = [string]$r.前; p = $pk; cd = [string]$r.KOMOKU_CD; new = [string]$r.後 }
                    }
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
Write-Host ' HbA1c・血小板・eGFR・中性脂肪 の判定を くくり判定・総合判定 に反映する' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  受診日    : {0}" -f ($dateList -join ', '))
if ($ukeList.Count -gt 0) { Write-Host ("  受付番号  : {0}" -f ($ukeList -join ', ')) }
if ($UkeFrom -gt 0 -or $UkeTo -gt 0) { Write-Host ("  受付番号の範囲: {0}〜{1}" -f $UkeFrom, $UkeTo) }
Write-Host '  糖代謝 600340 ← HbA1c / 血液一般 600310 ← 血小板 / 腎機能 600220 ← eGFR / 脂質代謝 600320 ← 中性脂肪 / 眼底 600060 / 総合判定 11000 ← 5 項目の最重' -ForegroundColor DarkGray
Write-Host '  中性脂肪(未判別) に判定が無い人は、空腹時の基準で判定と H/L 印を付ける (空腹時/随時中性脂肪に判定がある人は触らない)。' -ForegroundColor DarkGray
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
    $allCds = @($ITEMS | ForEach-Object { $_.Src }) + @($ITEMS | ForEach-Object { $_.Dst }) + @($SOGO)
    foreach ($it in $ITEMS) { if ($it.Alt) { $allCds += $it.Alt } }
    $cdLit = "'" + ($allCds -join "','") + "'"
    $cur = @{}   # pk -> cd -> @{ H = 判定; V = 値; HL = 印 }
    foreach ($r in (Invoke-DbQuery $conn "SELECT PK_SEQ, LTRIM(RTRIM(KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(KEKKA,''))) AS V, LTRIM(RTRIM(ISNULL(HL,''))) AS HL FROM T_KENSA WHERE PK_SEQ IN ($pkLit) AND LTRIM(RTRIM(KOMOKU_CD)) IN ($cdLit)").Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $cur.ContainsKey($pk)) { $cur[$pk] = @{} }
        $cur[$pk][[string]$r.CD] = @{ H = [string]$r.H; V = [string]$r.V; HL = [string]$r.HL }
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
            $h = (Normalize-Text $c[$it.Src].H).ToUpper()
            $src = $it.SrcName
            if ($it.Calc -and $h -eq '') {
                # 空腹時/随時中性脂肪に判定が付いている人 (自動判定が振り分け済み) は健診ナビに任せる
                $altDone = $false
                foreach ($a in $it.Alt) { if ($c.ContainsKey($a) -and (Normalize-Text $c[$a].H) -ne '') { $altDone = $true } }
                if ($altDone) { continue }
                $j = TgJudge $c[$it.Src].V
                if ($null -eq $j) { continue }
                $h = $j.H
                $plan += New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$p.YMD; 受付番号 = [string]$p.UKE; 氏名 = [string]$p.NAME
                    欄 = '中性脂肪の判定'; KOMOKU_CD = $it.Src; 前 = ''; 後 = $h; HL = $j.HL; 前HL = (Normalize-Text $c[$it.Src].HL)
                    理由 = ("中性脂肪 {0} を空腹時の基準で判定 (印 {1})" -f $c[$it.Src].V, $(if ($j.HL) { $j.HL } else { 'なし' })) }
                $src = '中性脂肪(空腹時基準)'
            }
            $r = Rank $h
            if ($r -gt $maxR) { $maxR = $r; $maxH = $h }
            if ($r -eq 0 -or -not $c.ContainsKey($it.Dst)) { continue }
            $dH = (Normalize-Text $c[$it.Dst].H).ToUpper()
            if ($r -gt (Rank $dH)) {
                $plan += New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$p.YMD; 受付番号 = [string]$p.UKE; 氏名 = [string]$p.NAME
                    欄 = $it.DstName; KOMOKU_CD = $it.Dst; 前 = $dH; 後 = $h; HL = $null; 前HL = ''; 理由 = ("{0} {1} = {2}" -f $src, $c[$it.Src].V, $h).Replace("  "," ") }
            }
        }
        $sH = (Normalize-Text $c[$SOGO].H).ToUpper()
        if ($maxR -gt (Rank $sH)) {
            $plan += New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$p.YMD; 受付番号 = [string]$p.UKE; 氏名 = [string]$p.NAME
                欄 = '総合判定'; KOMOKU_CD = $SOGO; 前 = $sH; 後 = $maxH; HL = $null; 前HL = ''; 理由 = ("5 項目の最重 = {0}" -f $maxH) }
        }
    }
    if ($skipNoSogo -gt 0) { Write-Host ("  総合判定が付いていないので対象外: {0} 人 (自動判定前)" -f $skipNoSogo) -ForegroundColor DarkGray }

    Write-Host ''
    Write-Host '=== 直す欄 ===' -ForegroundColor Cyan
    $plan | Select-Object 受診日, 受付番号, 氏名, 欄, 前, 後, @{n='印';e={ if ($null -ne $_.HL) { $_.HL } else { '' } }}, 理由 | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
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
                Write-Hantei $conn $tran $pk $t.KOMOKU_CD $t.後 $t.HL
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.KOMOKU_CD; 欄 = $t.欄; 前 = $t.前; 後 = $t.後
                    前HL = $(if ($null -ne $t.HL) { $t.前HL } else { '' }); 後HL = $(if ($null -ne $t.HL) { [string]$t.HL } else { '' }); 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [直した] {0} ({1}) : {2}" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, (($grp.Group | ForEach-Object { "{0} {1}→{2}" -f $_.欄, $_.前, $_.後 }) -join ' / ')) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) {
        $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 欄, 前, 後, 前HL, 後HL, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8
    }
    Write-Host ''
    Write-Host ("[完了] {0} か所を直しました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて直せなかった受付番号: {0} → 画面を閉じてもう一度実行" -f ($skipped -join ', ')) -ForegroundColor Red }
    Write-Host '   元に戻すときは kukuri_fix_undo.bat。' -ForegroundColor DarkGray
}
finally { $conn.Close() }
