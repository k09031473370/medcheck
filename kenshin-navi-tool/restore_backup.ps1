<#
.SYNOPSIS
  書込前バックアップから T_KENSA を戻す (restore_backup.ps1)

.DESCRIPTION
  form_import.ps1 が書込前に backup\ へ保存した CSV を読み、
  PK_SEQ + KOMOKU_CD 単位で 結果 / 結果CD / 判定記号 を元の値に戻す。

  既定はプレビュー。実際に戻すには -Commit を付ける。
  戻す前の状態も backup\ に保存するので、やり直しがきく。

.EXAMPLE
  # 何が戻るかを見る
  powershell -ExecutionPolicy Bypass -File restore_backup.ps1 -File backup\T_KENSA_2004717_20260731_094500.csv

  # 実際に戻す
  powershell -ExecutionPolicy Bypass -File restore_backup.ps1 -File backup\T_KENSA_2004717_20260731_094500.csv -Commit

  # 直近のバックアップを一覧する
  powershell -ExecutionPolicy Bypass -File restore_backup.ps1 -List
#>
[CmdletBinding()]
param(
    [string]$File,                # 戻す元のバックアップCSV
    [switch]$Commit,              # 付けると実際に戻す。付けなければ表示のみ
    [switch]$List,                # backup\ の一覧を表示して終了
    [string]$Only,                # 特定の項目コードだけ戻す (カンマ区切り)
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)

$ErrorActionPreference = 'Stop'
$BackupDir = Join-Path $PSScriptRoot 'backup'

# form_import.ps1 から接続まわりを借りる (同じ接続先を使うため)
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
# Invoke-Expression は呼んだ場所のスコープに定義されるので、必ずスクリプト直下で実行する
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Normalize-KenNo')
Invoke-Expression (Get-Part 'function Resolve-ConnectionString' 'function Get-CurrentKensa')
Invoke-Expression (Get-Part 'function Test-Locked' 'function Commit-Plan')

# ---- 一覧モード ----
if ($List -or -not $File) {
    if (-not (Test-Path $BackupDir)) { throw "backup フォルダがありません: $BackupDir" }
    Write-Host "=== backup フォルダの中身 (新しい順) ===" -ForegroundColor Cyan
    Get-ChildItem $BackupDir -File |
        Where-Object { $_.Name -like 'T_KENSA_*.csv' -or $_.Name -like 'T_KOJIN1_*.csv' -or $_.Name -like 'UKE_NO_*.csv' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 40 @{n='保存日時';e={$_.LastWriteTime}}, @{n='ファイル';e={$_.Name}}, @{n='サイズ';e={$_.Length}} |
        Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    if (-not $File) {
        Write-Host '戻すには -File でファイルを指定してください。' -ForegroundColor Yellow
    }
    if ($List) { return }
}

if (-not (Test-Path $File)) { throw "バックアップファイルがありません: $File" }
$rows = @(Import-Csv -Path $File -Encoding UTF8)
if ($rows.Count -eq 0) { throw "バックアップが空です: $File" }
$cols = $rows[0].PSObject.Properties.Name

# ---- 個人マスタ(社員番号)のバックアップなら、そちらの手順で戻して終わり ----
# このツールが個人マスタに書くのは社員番号(KOJIN_NO)だけなので、戻すのもそこだけ。
# 氏名や住所まで巻き戻すと、取込と関係のない手直しまで消えてしまう。
if ($cols -contains 'KOJIN_ID' -and $cols -notcontains 'KOMOKU_CD') {
    if ($rows.Count -ne 1) { throw "1人分のバックアップだけを指定してください (行数 $($rows.Count))" }
    $kojinId = Normalize-Text $rows[0].KOJIN_ID
    $bNo = Normalize-Text $rows[0].KOJIN_NO
    Write-Host ("バックアップ: {0}" -f $File) -ForegroundColor Cyan
    Write-Host ("対象 KOJIN_ID : {0} (個人マスタの社員番号)" -f $kojinId) -ForegroundColor Cyan

    $conn = Open-Db
    try {
        $dt = Invoke-DbQuery $conn 'SELECT KOJIN_ID, KANJI_SIMEI, KOJIN_NO FROM T_KOJIN1 WHERE KOJIN_ID = @k' @{ k = $kojinId }
        if ($dt.Rows.Count -eq 0) { throw "この KOJIN_ID が健診ナビにありません: $kojinId" }
        $nNo = Normalize-Text ([string]$dt.Rows[0].KOJIN_NO)
        Write-Host ("氏名: {0}" -f $dt.Rows[0].KANJI_SIMEI)
        if ($bNo -eq $nNo) {
            Write-Host 'バックアップの内容と現在の値は同じです。戻すものはありません。' -ForegroundColor Green
            return
        }
        Write-Host ''
        Write-Host ("社員番号  今: 「{0}」  →  戻す: 「{1}」" -f $nNo, $bNo)
        if (-not $Commit) {
            Write-Host ''
            Write-Host '※ 表示のみです。実際に戻すには -Commit を付けて再実行してください。' -ForegroundColor Yellow
            return
        }
        $safety = Join-Path $BackupDir ("T_KOJIN1_{0}_{1}_before_restore.csv" -f $kojinId, (Get-Date -Format 'yyyyMMdd_HHmmss'))
        (Invoke-DbQuery $conn 'SELECT * FROM T_KOJIN1 WHERE KOJIN_ID = @k' @{ k = $kojinId }) |
            Export-Csv -Path $safety -NoTypeInformation -Encoding UTF8
        Write-Host "[バックアップ] 戻す前の状態: $safety" -ForegroundColor DarkGray

        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KOJIN_IMPORT:" + $kojinId)
            $n = Invoke-DbExec $conn $tran 'UPDATE T_KOJIN1 SET KOJIN_NO = @v WHERE KOJIN_ID = @k' `
                 @{ v = $bNo; k = $kojinId }
            if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした。ロールバックします。" -f $n) }
            $tran.Commit()
            Write-Host ("[復元完了] 社員番号を元に戻しました (KOJIN_ID={0})" -f $kojinId) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }
    finally { $conn.Close() }
    return
}

# ---- 受付番号(UKE_NO_*.csv)のバックアップなら、そちらの手順で戻して終わり ----
# 「受付番号を設定」で書いた T_KENSIN.UKE_NO_KENSA を、設定前の値に戻す。
# 設定前が空欄だった人は空欄に戻す。
if ($cols -contains 'PkSeq' -and $cols -contains '現在の番号') {
    $conn = Open-Db
    try {
        $plan = @()
        foreach ($r in $rows) {
            $pk = Normalize-Text $r.PkSeq
            if ($pk -eq '') { continue }
            $before = Normalize-KenNo ([string]$r.'現在の番号')      # 設定前の値
            $dt = Invoke-DbQuery $conn 'SELECT UKE_NO_KENSA, D_KENSIN FROM T_KENSIN WHERE PK_SEQ = @p' @{ p = $pk }
            if ($dt.Rows.Count -eq 0) { Write-Warning "受診が見つかりません (PK_SEQ=$pk)"; continue }
            $now = Normalize-KenNo ([string]$dt.Rows[0].UKE_NO_KENSA)
            if ($now -eq $before) { continue }                        # すでに元どおり
            $plan += [pscustomobject]@{
                PkSeq = $pk; 受診日 = [string]$dt.Rows[0].D_KENSIN
                氏名 = [string]$r.氏名; いまの番号 = $now
                戻す番号 = $(if ($before -eq '') { '(空欄)' } else { $before })
                Before = $before
            }
        }
        if ($plan.Count -eq 0) {
            Write-Host 'すでに元の状態です。戻すものはありません。' -ForegroundColor Green
            return
        }
        Write-Host ''
        Write-Host '=== 受付番号を元に戻す ===' -ForegroundColor Cyan
        $plan | Select-Object 受診日, 氏名, いまの番号, 戻す番号, PkSeq |
            Format-Table -AutoSize | Out-String -Width 200 | Write-Host
        if (-not $Commit) {
            Write-Host '※ 表示のみです。実際に戻すには -Commit を付けて再実行してください。' -ForegroundColor Yellow
            return
        }
        $tran = $conn.BeginTransaction()
        try {
            foreach ($t in $plan) {
                $v = $(if ($t.Before -eq '') { [DBNull]::Value } else { $t.Before })
                $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSIN SET UKE_NO_KENSA = @no WHERE PK_SEQ = @p' `
                     @{ no = $v; p = $t.PkSeq }
                if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした (PK_SEQ={1})。ロールバックします。" -f $n, $t.PkSeq) }
            }
            $tran.Commit()
            Write-Host ("[復元完了] {0} 人の受付番号を元に戻しました。" -f $plan.Count) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }
    finally { $conn.Close() }
    return
}

foreach ($need in @('PK_SEQ','KOMOKU_CD','KEKKA')) {
    if ($cols -notcontains $need) {
        throw "このCSVは T_KENSA のバックアップではないようです ($need の列がありません): $File"
    }
}

$pkList = @($rows | ForEach-Object { Normalize-Text $_.PK_SEQ } | Sort-Object -Unique)
if ($pkList.Count -ne 1) { throw "1人分のバックアップだけを指定してください (PK_SEQ が {0} 種類あります)" -f $pkList.Count }
$pkSeq = $pkList[0]

$wants = @()
if ($Only) { $wants = @($Only -split '[,、]' | ForEach-Object { Normalize-Text $_ } | Where-Object { $_ -ne '' }) }

Write-Host ("バックアップ: {0}" -f $File) -ForegroundColor Cyan
Write-Host ("対象 PK_SEQ : {0} / {1} 行" -f $pkSeq, $rows.Count) -ForegroundColor Cyan

$conn = Open-Db
try {
    $cur = @{}
    $dt = Invoke-DbQuery $conn 'SELECT KOMOKU_CD, KEKKA, KEKKA_CD, HANTEI_KIGO FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pkSeq }
    foreach ($r in $dt.Rows) { $cur[(Normalize-Text $r.KOMOKU_CD)] = $r }
    if ($cur.Count -eq 0) { throw "この PK_SEQ の検査行が健診ナビにありません: $pkSeq" }

    $diff = @()
    foreach ($b in $rows) {
        $k = Normalize-Text $b.KOMOKU_CD
        if ($k -eq '') { continue }
        if ($wants.Count -gt 0 -and $wants -notcontains $k) { continue }
        if (-not $cur.ContainsKey($k)) { continue }
        $now = $cur[$k]
        $bK  = Normalize-Text $b.KEKKA
        $bC  = Normalize-Text $b.KEKKA_CD
        $bH  = Normalize-Text $b.HANTEI_KIGO
        $nK  = Normalize-Text ([string]$now.KEKKA)
        $nC  = Normalize-Text ([string]$now.KEKKA_CD)
        $nH  = Normalize-Text ([string]$now.HANTEI_KIGO)
        if ($bK -eq $nK -and $bC -eq $nC -and $bH -eq $nH) { continue }
        $diff += New-Object PSObject -Property @{
            KOMOKU_CD = $k
            今の値 = $nK; 戻す値 = $bK
            今のCD = $nC; 戻すCD = $bC
            今の判定 = $nH; 戻す判定 = $bH
        }
    }

    if ($diff.Count -eq 0) {
        Write-Host 'バックアップの内容と現在の値は同じです。戻すものはありません。' -ForegroundColor Green
        return
    }

    Write-Host ''
    $diff | Select-Object KOMOKU_CD, 今の値, 戻す値, 今のCD, 戻すCD, 今の判定, 戻す判定 |
        Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    Write-Host ("戻す項目: {0} 件" -f $diff.Count) -ForegroundColor Cyan

    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ 表示のみです。実際に戻すには -Commit を付けて再実行してください。' -ForegroundColor Yellow
        return
    }

    # 先に編集中かどうかを見る。中止するのに控えだけ残ると、
    # 次に戻すファイルを選ぶときの一覧が紛らわしくなるため。
    $lock = Test-Locked $conn $pkSeq
    if ($lock) { throw "この受診者は健診ナビの結果入力画面で編集中です ($lock)。画面を閉じてから再実行してください。" }

    # 戻す前の状態も保存しておく (やり直せるように)
    $safety = Join-Path $BackupDir ("T_KENSA_{0}_{1}_before_restore.csv" -f $pkSeq, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    (Invoke-DbQuery $conn 'SELECT * FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pkSeq }) |
        Export-Csv -Path $safety -NoTypeInformation -Encoding UTF8
    Write-Host "[バックアップ] 戻す前の状態: $safety" -ForegroundColor DarkGray

    $tran = $conn.BeginTransaction()
    try {
        Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pkSeq)
        $lock2 = Test-Locked $conn $pkSeq $tran
        if ($lock2) { throw "この受診者は健診ナビの結果入力画面で編集中です ($lock2)。" }
        $done = 0
        foreach ($d in $diff) {
            $n = Invoke-DbExec $conn $tran `
                'UPDATE T_KENSA SET KEKKA = @k, KEKKA_CD = @kc, HANTEI_KIGO = @h WHERE PK_SEQ = @p AND KOMOKU_CD = @cd' `
                @{ k = $d.戻す値; kc = $d.戻すCD; h = $d.戻す判定; p = $pkSeq; cd = $d.KOMOKU_CD }
            if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした (KOMOKU_CD={1})。ロールバックします。" -f $n, $d.KOMOKU_CD) }
            $done++
        }
        $tran.Commit()
        Write-Host ("[復元完了] {0} 項目を元に戻しました (PK_SEQ={1})" -f $done, $pkSeq) -ForegroundColor Green
        Write-Host '※ 健診ナビで対象者を開き「自動判定」を実行し直してください。'
    }
    catch { $tran.Rollback(); throw }
}
finally { $conn.Close() }
