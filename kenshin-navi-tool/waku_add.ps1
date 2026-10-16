<#
.SYNOPSIS
  健診ナビの結果の枠 (T_KENSA の行) を足す (waku_add.ps1)

.DESCRIPTION
  コースに無い検査項目の枠を、同じ項目を持っている別の人の行を手本にしてコピーし、
  対象の人に空の枠として追加する。値 (KEKKA) は入れない。あとで form_import の取込で値を入れる。
    - 健診ナビの「検査追加」は T_KENSA と T_RYOUKIN(料金明細) の 2 表に行を作るが、
      このツールは T_KENSA だけ作る (請求は健診ナビから出さない、という前提。2026/10/16 ユーザー判断)。
    - F_KOMOKU は既定で '2' (健診ナビが検査追加で付ける印)。
    - 必ずプレビューを出し、y を押したときだけ書く。書く前に本人の T_KENSA を backup\ に保存し、
      足した行の一覧を backup\waku_add_<日時>.csv に残す。-Undo でその一覧の行を (値が空のままなら) 消せる。
  例:
    waku_add.ps1 -Items 068689,069010 -Dates 2026/09/03,2026/09/09,2026/09/10 -Course KY101,KY102 -UkeFrom 4631 -UkeTo 4831
    waku_add.ps1 -Items OPJ010,069245,069246 -Dates 2026/09/03,2026/09/09 -UkeNo 4634,4635,4658,4693,4702
    waku_add.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$Items,        # 足す項目CD (カンマ区切り)。セットの子を足すときは親も一緒に書く (例 OPJ010,069245,069246)
    [string]$Dates,        # 受診日 (カンマ区切り)
    [string]$UkeNo,        # 受付番号 (カンマ区切り・任意)
    [string]$Course,       # コースCD (カンマ区切り・任意)
    [int]$UkeFrom = 0,     # 受付番号の範囲 (任意)
    [int]$UkeTo = 0,
    [string]$FKomoku = '2',
    [switch]$Commit,       # 付けると確認なしで書く。付けなければプレビュー後に y/n を聞く
    [string]$Undo,         # backup\waku_add_*.csv のパス、または ask
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
function DbVal($v) { if ($null -eq $v -or ($v -is [System.DBNull]) -or ([string]$v) -eq '') { return [DBNull]::Value } return $v }
function Split-List([string]$s) { return @(($s -split '[,、\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }

# ============================================================
# 元に戻す (-Undo)
# ============================================================
if ($Undo) {
    if ($Undo -eq 'ask') {
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'waku_add_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\waku_add_*.csv) がありません。'; return }
        for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
        $a = Read-Host '戻す記録の番号 (それ以外 = やめる)'
        if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt $files.Count) { Write-Host 'やめました。'; return }
        $Undo = $files[[int]$a - 1].FullName
    }
    if (-not (Test-Path $Undo)) { throw "記録ファイルがありません: $Undo" }
    $rows = @(Import-Csv -Path $Undo -Encoding UTF8)
    Write-Host ("この記録で足した枠 {0} 行を消します (値が入っている行は消しません): {1}" -f $rows.Count, $Undo) -ForegroundColor Yellow
    $a = Read-Host '消しますか (y = 消す / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    $conn = Open-Db
    try {
        $deleted = 0; $kept = 0
        foreach ($grp in ($rows | Group-Object PK_SEQ)) {
            $pk = [decimal]$grp.Name
            $lock = Test-Locked $conn $pk
            if ($lock) { Write-Host ("  [とばしました] PK_SEQ {0} は健診ナビで編集中です ({1})" -f $pk, $lock) -ForegroundColor Yellow; continue }
            [void](Backup-Kensa $conn $pk)
            $tran = $conn.BeginTransaction()
            try {
                Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
                foreach ($r in $grp.Group) {
                    $n = Invoke-DbExec $conn $tran "DELETE FROM T_KENSA WHERE PK_SEQ = @p AND KOMOKU_CD = @cd AND LTRIM(RTRIM(ISNULL(KEKKA,''))) = '' AND LTRIM(RTRIM(ISNULL(KEKKA_CD,''))) = ''" @{ p = $pk; cd = [string]$r.KOMOKU_CD }
                    if ($n -eq 1) { $deleted++ } else { $kept++ }
                }
                $tran.Commit()
            } catch { $tran.Rollback(); throw }
        }
        Write-Host ("[完了] 消した枠 {0} 行 / 値が入っていたので残した枠 {1} 行" -f $deleted, $kept) -ForegroundColor Green
    } finally { $conn.Close() }
    return
}

# ============================================================
# 枠を足す
# ============================================================
$itemList = Split-List $Items
$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
$ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })
$crsList  = @(Split-List $Course | ForEach-Object { if ($_ -notmatch '^[A-Za-z0-9_\-]+$') { throw "コースCDがおかしいです: $_" }; $_.ToUpper() })
if ($itemList.Count -eq 0) { throw '-Items に足す項目CDを指定してください (例 068689,069010)' }
if ($dateList.Count -eq 0) { throw '-Dates に受診日を指定してください (例 2026/09/03,2026/09/09)' }
foreach ($c in $itemList) { if ($c -notmatch '^[A-Za-z0-9\-]+$') { throw "項目CDがおかしいです: $c" } }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 結果の枠を足す (T_KENSA に空の行を追加)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  項目      : {0}" -f ($itemList -join ', '))
Write-Host ("  受診日    : {0}" -f ($dateList -join ', '))
if ($ukeList.Count -gt 0) { Write-Host ("  受付番号  : {0}" -f ($ukeList -join ', ')) }
if ($crsList.Count -gt 0) { Write-Host ("  コース    : {0}" -f ($crsList -join ', ')) }
if ($UkeFrom -gt 0 -or $UkeTo -gt 0) { Write-Host ("  受付番号の範囲: {0}〜{1}" -f $UkeFrom, $UkeTo) }
Write-Host ("  F_KOMOKU  : {0} (健診ナビの検査追加と同じ印)" -f $FKomoku)
Write-Host '  ※ 料金明細 (T_RYOUKIN) は作りません。請求は健診ナビから出さない前提です。' -ForegroundColor DarkGray
Write-Host ''

$conn = Open-Db
try {
    # ---- 対象者 ----
    $dLit = "'" + ($dateList -join "','") + "'"
    $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit)"
    if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
    if ($crsList.Count -gt 0) { $where += " AND UPPER(LTRIM(RTRIM(ISNULL(s.COURSE_CD,'')))) IN ('" + ($crsList -join "','") + "')" }
    if ($UkeFrom -gt 0) { $where += " AND s.UKE_NO_KENSA >= $UkeFrom" }
    if ($UkeTo -gt 0)   { $where += " AND s.UKE_NO_KENSA <= $UkeTo" }
    $people = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS COURSE, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS DANTAI
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE $where
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@).Rows
    if ($people.Count -eq 0) { Write-Host '対象の人がいません (受診日・受付番号・コースを見直してください)。' -ForegroundColor Yellow; return }
    Write-Host ("  対象者: {0} 人" -f $people.Count)

    # ---- いまある枠と SEQ1/SEQ2 ----
    $pkLit = ($people | ForEach-Object { [string]$_.PK_SEQ }) -join ','
    $have = @{}; $seq = @{}
    foreach ($r in (Invoke-DbQuery $conn "SELECT PK_SEQ, LTRIM(RTRIM(KOMOKU_CD)) AS CD, SEQ1, SEQ2 FROM T_KENSA WHERE PK_SEQ IN ($pkLit)").Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $have.ContainsKey($pk)) { $have[$pk] = @{} }
        $have[$pk][[string]$r.CD] = 1
        if (-not $seq.ContainsKey($pk)) { $seq[$pk] = @{ S1 = $r.SEQ1; S2 = $r.SEQ2 } }
    }

    # ---- 手本の行 (項目ごと) ----
    $names = @{}
    foreach ($r in (Invoke-DbQuery $conn ("SELECT LTRIM(RTRIM(KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS NM FROM T_KOMOKU WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('" + ($itemList -join "','") + "')")).Rows) { $names[[string]$r.CD] = [string]$r.NM }
    $tCrs = @($people | ForEach-Object { [string]$_.COURSE } | Sort-Object -Unique)
    $cLit = "'" + ($tCrs -join "','") + "'"
    $tmpl = @{}
    foreach ($cd in $itemList) {
        if (-not $names.ContainsKey($cd)) { Write-Host ("  ★項目マスタに無い項目CD: {0}" -f $cd) -ForegroundColor Red; continue }
        $t = (Invoke-DbQuery $conn @"
SELECT TOP 1 k.KOMOKU_CD, k.HL, k.F_KOMOKU, k.F_HANTEI, k.HANTEI_CD, k.OYA_KOMOKU_CD, k.F_GOJITU,
       s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS COURSE
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = @cd AND s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01'
ORDER BY CASE WHEN LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) IN ($cLit) THEN 0 ELSE 1 END,
         CASE WHEN s.D_KENSIN IN ($dLit) THEN 0 ELSE 1 END,
         CASE WHEN LTRIM(RTRIM(ISNULL(k.F_KOMOKU,''))) = '0' THEN 0 ELSE 1 END,
         s.D_KENSIN DESC
"@ @{ cd = $cd }).Rows
        if ($t.Count -eq 0) { Write-Host ("  ★手本になる行がありません (2026 年に誰もこの枠を持っていない): {0} {1}" -f $cd, $names[$cd]) -ForegroundColor Red; continue }
        $tmpl[$cd] = $t[0]
        Write-Host ("  手本 {0,-8} {1,-16} ← 受付番号 {2} ({3} {4})  親={5} F_HANTEI={6} HANTEI_CD={7}" -f $cd, $names[$cd], $t[0].UKE, $t[0].YMD, $t[0].COURSE,
            (Normalize-Text ([string]$t[0].OYA_KOMOKU_CD)), (Normalize-Text ([string]$t[0].F_HANTEI)), (Normalize-Text ([string]$t[0].HANTEI_CD))) -ForegroundColor DarkGray
    }

    # ---- 計画 ----
    $plan = @()
    foreach ($p in $people) {
        $pk = [string]$p.PK_SEQ
        foreach ($cd in $itemList) {
            $rep = New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$p.YMD; 受付番号 = [string]$p.UKE; 氏名 = [string]$p.NAME; コース = [string]$p.COURSE
                                                    項目CD = $cd; 項目名 = $names[$cd]; 状態 = '' }
            if (-not $tmpl.ContainsKey($cd))        { $rep.状態 = '手本なし'; $plan += $rep; continue }
            if ($have.ContainsKey($pk) -and $have[$pk].ContainsKey($cd)) { $rep.状態 = '既にある'; $plan += $rep; continue }
            if (-not $seq.ContainsKey($pk))        { $rep.状態 = 'この人に枠が1つも無い(SEQ不明)'; $plan += $rep; continue }
            $oya = Normalize-Text ([string]$tmpl[$cd].OYA_KOMOKU_CD)
            if ($oya -ne '' -and -not ($have[$pk].ContainsKey($oya) -or ($itemList -contains $oya))) { $rep.状態 = "親の枠なし($oya)"; $plan += $rep; continue }
            $rep.状態 = 'OK'
            $plan += $rep
        }
    }

    Write-Host ''
    Write-Host '=== 足す枠 ===' -ForegroundColor Cyan
    $plan | Where-Object { $_.状態 -ne '既にある' } | Select-Object 受診日, 受付番号, 氏名, コース, 項目CD, 項目名, 状態 |
        Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    $ok   = @($plan | Where-Object { $_.状態 -eq 'OK' })
    $has  = @($plan | Where-Object { $_.状態 -eq '既にある' })
    $err  = @($plan | Where-Object { $_.状態 -ne 'OK' -and $_.状態 -ne '既にある' })
    Line
    Write-Host ("足す: {0} 行 ({1} 人) / 既にある: {2} 行 / 要確認: {3} 行" -f $ok.Count, @($ok | Group-Object PkSeq).Count, $has.Count, $err.Count) `
        -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })
    Line
    if ($ok.Count -eq 0) { Write-Host '足すものがありません。'; return }
    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
        $a = Read-Host '上の枠を足しますか (y = 足す / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }

    # ---- 書く ----
    $undoFile = Join-Path $BackupDir ("waku_add_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $undoRows = @()
    $done = 0
    foreach ($grp in ($ok | Group-Object PkSeq)) {
        $pk = [decimal]$grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} は健診ナビで編集中です ({1})" -f $grp.Group[0].氏名, $lock) -ForegroundColor Yellow; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            $lock2 = Test-Locked $conn $pk $tran
            if ($lock2) { throw "書込直前に健診ナビで開かれました ($lock2)" }
            foreach ($t in $grp.Group) {
                $m = $tmpl[$t.項目CD]
                # 二重に足さない (プレビュー後に誰かが足した場合)
                $dup = Invoke-DbQuery $conn 'SELECT COUNT(*) AS N FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ p = $pk; cd = $t.項目CD } $tran
                if ([int]$dup.Rows[0].N -gt 0) { Write-Host ("  [既にある] {0} {1}" -f $t.氏名, $t.項目CD) -ForegroundColor DarkGray; continue }
                $n = Invoke-DbExec $conn $tran @'
INSERT INTO T_KENSA (PK_SEQ, SEQ1, SEQ2, KOMOKU_CD, KEKKA, KEKKA_CD, HANTEI_KIGO, HL, F_KOMOKU, F_HANTEI, HANTEI_CD, OYA_KOMOKU_CD, KEKKAE, F_GOJITU)
VALUES (@p, @s1, @s2, @cd, NULL, NULL, NULL, @hl, @fk, @fh, @hc, @oya, NULL, @fg)
'@ @{ p = $pk; s1 = (DbVal $seq[[string]$pk].S1); s2 = (DbVal $seq[[string]$pk].S2); cd = [string]$m.KOMOKU_CD
       hl = (DbVal $m.HL); fk = $FKomoku; fh = (DbVal $m.F_HANTEI); hc = (DbVal $m.HANTEI_CD); oya = (DbVal $m.OYA_KOMOKU_CD); fg = (DbVal $m.F_GOJITU) }
                if ($n -ne 1) { throw ("INSERT影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.項目CD) }
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = [string]$m.KOMOKU_CD; 受付番号 = $t.受付番号; 氏名 = $t.氏名; 項目名 = $t.項目名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [追加] {0} ({1}) : {2} 行" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $grp.Count) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) {
        $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 受付番号, 氏名, 項目名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8
    }
    Write-Host ''
    Write-Host ("[完了] {0} 行の枠を足しました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    Write-Host '   元に戻すときは waku_add_undo.bat (値が入る前なら消せます)。' -ForegroundColor DarkGray
    Write-Host '   このあと form_import (start.bat) の取込で値を入れてください。' -ForegroundColor Yellow
}
finally { $conn.Close() }
