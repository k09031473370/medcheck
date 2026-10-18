<#
.SYNOPSIS
  石屋 3 日分 (199 人) の検査結果を丸ごと保存し、あとで丸ごと戻す (ishiya_snapshot.ps1)

.DESCRIPTION
  誤って自動判定をかけたときに、今日の最終の状態へ戻すための道具。
    -Save    : T_KENSA の 199 人分の行を、そのまま (空白も含めて) backup\ishiya_snapshot_<日時>.csv に保存する。DB は読むだけ。
    -Restore : 保存ファイルと今の DB を 1 行ずつ比べ、違う欄 (結果・結果CD・判定・印・所見 など) を保存したときの値に戻す。
               プレビュー → y。書く前に 1 人ずつ T_KENSA を backup\ に保存し、戻した欄は backup\ishiya_restore_<日時>.csv に残す。
               保存ファイルに無い行 (自動判定で新しくできた行) は触らずに件数だけ出す。健診ナビで開いている人はとばす。
  例:
    ishiya_snapshot.ps1 -Save
    ishiya_snapshot.ps1 -Restore                     (いちばん新しい保存ファイルを使う)
    ishiya_snapshot.ps1 -Restore -File backup\ishiya_snapshot_20261018_143000.csv
#>
[CmdletBinding()]
param(
    [switch]$Save,
    [switch]$Restore,
    [string]$File,
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

$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$SQL = "SELECT k.* FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ WHERE $ISHIYA ORDER BY k.PK_SEQ, k.KOMOKU_CD"
# 戻す欄: 結果・判定・印・所見など (更新日時や担当者の欄は戻さない)
$COLPAT = '^(KEKKA|HANTEI|HL$|SHOKEN|SYOKEN)'
$SKIPPAT = '(DATE|YMD|TIME|DT$|USER|TANTO|UPD|INS)'
$SQLK = "SELECT s.* FROM T_KENSIN s WHERE $ISHIYA ORDER BY s.PK_SEQ"
function Key($pk, $cd) { return ([string]$pk).Trim() + '|' + ([string]$cd).Trim() }
function Str($v) { if ($null -eq $v -or $v -is [DBNull]) { return '' } else { return [string]$v } }

if (-not $Save -and -not $Restore) { throw '-Save か -Restore を指定してください。' }

$conn = Open-Db
try {
    if ($Save) {
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $dt = Invoke-DbQuery $conn $SQL
        $f = Join-Path $BackupDir ("ishiya_snapshot_{0}.csv" -f $ts)
        $dt | Export-Csv -Path $f -NoTypeInformation -Encoding UTF8
        # 受診 (T_KENSIN) も保存する。戻しはしないが、自動判定のあとに変わった欄を一覧にするため
        $dk = Invoke-DbQuery $conn $SQLK
        $fk = Join-Path $BackupDir ("ishiya_kensin_{0}.csv" -f $ts)
        $dk | Export-Csv -Path $fk -NoTypeInformation -Encoding UTF8
        $np = @($dt.Rows | ForEach-Object { [string]$_.PK_SEQ } | Sort-Object -Unique).Count
        Write-Host ''
        Write-Host ("[保存しました] {0} 人 / 検査 {1} 行 / 受診 {2} 行" -f $np, $dt.Rows.Count, $dk.Rows.Count) -ForegroundColor Green
        Write-Host ("  {0}" -f $f)
        Write-Host ("  {0}" -f $fk)
        Write-Host '  誤って自動判定をかけたら ishiya_snapshot_restore.bat で、この時点に戻せます。' -ForegroundColor Yellow
        Write-Host '  ★このあと画面で直したら、もう一度保存し直すこと (古い保存で戻すと、その直しが消える)。' -ForegroundColor Yellow
        return
    }

    # ---- 戻す ----
    if (-not $File) {
        $last = @(Get-ChildItem -Path $BackupDir -Filter 'ishiya_snapshot_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($last.Count -eq 0) { throw '保存ファイル (backup\ishiya_snapshot_*.csv) がありません。' }
        for ($i = 0; $i -lt [Math]::Min(5, $last.Count); $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $last[$i].Name, $last[$i].LastWriteTime) }
        $a = Read-Host '使う保存ファイルの番号 (Enter = 1 / それ以外 = やめる)'
        if ([string]::IsNullOrWhiteSpace($a)) { $a = '1' }
        if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt [Math]::Min(5, $last.Count)) { Write-Host 'やめました。'; return }
        $File = $last[[int]$a - 1].FullName
    }
    if (-not (Test-Path $File)) { throw "保存ファイルがありません: $File" }
    $snap = @(Import-Csv -Path $File -Encoding UTF8)
    if ($snap.Count -eq 0) { throw "保存ファイルが空です: $File" }
    $cur = Invoke-DbQuery $conn $SQL
    $allCols = @($snap[0].PSObject.Properties.Name | Where-Object { $cur.Columns.Contains($_) -and $_ -notin @('PK_SEQ','KOMOKU_CD') })
    $cols = @($allCols | Where-Object { $_ -match $COLPAT -and $_ -notmatch $SKIPPAT })
    $other = @($allCols | Where-Object { $cols -notcontains $_ })
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    Write-Host ''
    Write-Host (' 石屋 3 日分を保存した時点に戻す: {0}' -f $File) -ForegroundColor Cyan
    Write-Host ('  戻す欄: {0}' -f ($cols -join ', ')) -ForegroundColor DarkGray

    $S = @{}; foreach ($r in $snap) { $S[(Key $r.PK_SEQ $r.KOMOKU_CD)] = $r }
    $NOW = @{}; foreach ($r in $cur.Rows) { $NOW[(Key $r.PK_SEQ $r.KOMOKU_CD)] = $r }
    $plan = @(); $newRows = @(); $otherDiff = @{}
    foreach ($r in $cur.Rows) {
        $k = Key $r.PK_SEQ $r.KOMOKU_CD
        if (-not $S.ContainsKey($k)) { $newRows += [PSCustomObject]@{ PK_SEQ = ([string]$r.PK_SEQ).Trim(); KOMOKU_CD = ([string]$r.KOMOKU_CD).Trim(); 結果 = (Str $r['KEKKA']); 判定 = (Str $r['HANTEI_KIGO']) }; continue }
        $b = $S[$k]; $chg = [ordered]@{}; $was = [ordered]@{}
        foreach ($c in $cols) { $bv = Str $b.$c; $nv = Str $r[$c]; if ($bv -ne $nv) { $chg[$c] = $bv; $was[$c] = $nv } }
        foreach ($c in $other) { if ((Str $b.$c) -ne (Str $r[$c])) { $otherDiff[$c] = 1 + [int]$otherDiff[$c] } }
        if ($chg.Count -gt 0) {
            $plan += [PSCustomObject]@{ PkSeq = ([string]$r.PK_SEQ).Trim(); KOMOKU_CD = ([string]$r.KOMOKU_CD).Trim(); Chg = $chg; Was = $was
                今 = ($(foreach ($c in $chg.Keys) { "{0}={1}" -f $c, $was[$c] }) -join ' / '); 戻す = ($(foreach ($c in $chg.Keys) { "{0}={1}" -f $c, $chg[$c] }) -join ' / ') }
        }
    }
    $missing = @($S.Keys | Where-Object { -not $NOW.ContainsKey($_) } | ForEach-Object { $x = $S[$_]; [PSCustomObject]@{ PK_SEQ = $x.PK_SEQ; KOMOKU_CD = $x.KOMOKU_CD; 結果 = $x.KEKKA; 判定 = $x.HANTEI_KIGO } })
    Write-Host ''
    Write-Host ("  戻す行: {0} 行 ({1} 人)" -f $plan.Count, @($plan | Group-Object PkSeq).Count)
    Write-Host ("  保存後に新しくできた行 (戻さない・一覧に出す): {0} 行 / 保存にあって今は無い行 (戻せない・一覧に出す): {1} 行" -f $newRows.Count, $missing.Count)
    if ($newRows.Count -gt 0) { $fn = Join-Path $BackupDir ("ishiya_restore_new_rows_{0}.csv" -f $stamp); $newRows | Export-Csv -Path $fn -NoTypeInformation -Encoding UTF8; Write-Host ("    ★新しい行の一覧: {0}" -f $fn) -ForegroundColor Yellow }
    if ($missing.Count -gt 0) { $fm = Join-Path $BackupDir ("ishiya_restore_missing_rows_{0}.csv" -f $stamp); $missing | Export-Csv -Path $fm -NoTypeInformation -Encoding UTF8; Write-Host ("    ★無くなった行の一覧: {0}" -f $fm) -ForegroundColor Yellow }
    if ($otherDiff.Count -gt 0) {
        Write-Host '  ★戻さない欄にも違いがあります (知らせてください):' -ForegroundColor Yellow
        foreach ($c in $otherDiff.Keys) { Write-Host ("    {0}: {1} 行" -f $c, $otherDiff[$c]) -ForegroundColor Yellow }
    }
    # 受診 (T_KENSIN) の違い (読むだけ)
    $kf = $File -replace 'ishiya_snapshot_', 'ishiya_kensin_'
    if ((Test-Path $kf) -and ($kf -ne $File)) {
        $ks = @(Import-Csv -Path $kf -Encoding UTF8); $kc = Invoke-DbQuery $conn $SQLK; $KSmap = @{}; foreach ($x in $ks) { $KSmap[([string]$x.PK_SEQ).Trim()] = $x }
        $kd = @{}
        foreach ($r in $kc.Rows) { $x = $KSmap[([string]$r.PK_SEQ).Trim()]; if (-not $x) { continue }
            foreach ($c in @($kc.Columns | ForEach-Object { $_.ColumnName })) { if ($x.PSObject.Properties.Name -contains $c -and (Str $x.$c) -ne (Str $r[$c])) { $kd[$c] = 1 + [int]$kd[$c] } } }
        if ($kd.Count -gt 0) { Write-Host '  ★受診 (T_KENSIN) の欄にも違いがあります (このツールでは戻さない。知らせてください):' -ForegroundColor Yellow; foreach ($c in $kd.Keys) { Write-Host ("    {0}: {1} 人" -f $c, $kd[$c]) -ForegroundColor Yellow } }
        else { Write-Host '  受診 (T_KENSIN) は保存した時点と同じです。' -ForegroundColor DarkGray }
    }
    if ($plan.Count -eq 0) { Write-Host '  検査の結果・判定は保存した時点と同じです。戻すものはありません。' -ForegroundColor Green; return }
    Write-Host '  項目ごとの件数:'
    $plan | Group-Object KOMOKU_CD | Sort-Object Count -Descending | Select-Object -First 25 | ForEach-Object { Write-Host ("    {0,-10} {1} 行" -f $_.Name, $_.Count) }
    Write-Host '  例 (はじめの 15 行):'
    $plan | Select-Object -First 15 PkSeq, KOMOKU_CD, 今, 戻す | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
    Write-Host '※ まだ何も書き込んでいません。健診ナビの画面は閉じてから実行してください (開いている人はとばします)。' -ForegroundColor Yellow
    Write-Host '※ 保存したあとに画面で正しく直したものがあれば、それも保存時点に戻ります。心当たりがあれば y を押さないでください。' -ForegroundColor Yellow
    $a = Read-Host '保存した時点に戻しますか (y = 戻す / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }

    $types = @{}; foreach ($c in $cols) { $types[$c] = $cur.Columns[$c].DataType }
    $log = @(); $done = 0; $skipped = @(); $failed = @()
    foreach ($grp in ($plan | Group-Object PkSeq)) {
        $pk = [decimal]$grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] PK_SEQ {0} は健診ナビで開かれています ({1})" -f $pk, $lock) -ForegroundColor Yellow; $skipped += $grp.Name; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        $mine = @()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            foreach ($t in $grp.Group) {
                # プレビューのあとで値が変わっていたら、この人は戻さない (ロールバック)
                $q = Invoke-DbQuery $conn ("SELECT " + (($t.Chg.Keys | ForEach-Object { "[$_]" }) -join ', ') + " FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd") @{ p = $pk; cd = $t.KOMOKU_CD } $tran
                if ($q.Rows.Count -ne 1) { throw ("行が {0} 件 ({1})" -f $q.Rows.Count, $t.KOMOKU_CD) }
                foreach ($c in $t.Chg.Keys) { if ((Str $q.Rows[0][$c]) -ne $t.Was[$c]) { throw ("プレビューのあとで {0} の {1} が変わっていた" -f $t.KOMOKU_CD, $c) } }
                $sets = @(); $prm = @{ p = $pk; cd = $t.KOMOKU_CD }; $i = 0
                foreach ($c in $t.Chg.Keys) {
                    $i++; $sets += ("[{0}] = @v{1}" -f $c, $i)
                    $v = $t.Chg[$c]
                    $prm["v$i"] = $(if ($v -eq '' -and $types[$c] -ne [string]) { [DBNull]::Value } else { $v })
                }
                $n = Invoke-DbExec $conn $tran ("UPDATE T_KENSA SET " + ($sets -join ', ') + " WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd") $prm
                if ($n -ne 1) { throw ("UPDATE 影響行数が {0} ({1})" -f $n, $t.KOMOKU_CD) }
                $mine += [PSCustomObject]@{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.KOMOKU_CD; 戻す前 = $t.今; 戻した値 = $t.戻す }
            }
            $tran.Commit()
            $log += $mine; $done += $mine.Count
        } catch {
            $tran.Rollback()
            Write-Host ("  [戻さず・元のまま] PK_SEQ {0}: {1}" -f $pk, $_.Exception.Message) -ForegroundColor Red
            $failed += $grp.Name
        }
    }
    $lf = Join-Path $BackupDir ("ishiya_restore_{0}.csv" -f $stamp)
    if ($log.Count -gt 0) { $log | Export-Csv -Path $lf -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} 行を保存した時点に戻しました ({1} 人)。記録: {2}" -f $done, @($log | Group-Object PK_SEQ).Count, $lf) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて戻せなかった PK_SEQ: {0} → 画面を閉じてもう一度実行" -f ($skipped -join ', ')) -ForegroundColor Red }
    if ($failed.Count -gt 0) { Write-Host ("  ★途中で値が変わっていた・行が合わないので戻さなかった PK_SEQ: {0}" -f ($failed -join ', ')) -ForegroundColor Red }
}
finally { $conn.Close() }
