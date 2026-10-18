<#
.SYNOPSIS
  自動判定のかけ直しの前に 医師指示事項 (コメント) と全項目の判定を保存し、かけ直しのあとに消えたコメントだけ戻す (comment_keep.ps1)

.DESCRIPTION
  自動判定は医師指示事項 11001〜11009 を作り直すので、画面で手入力したコメントやツールで足したコメントが消える。
  どれが手入力かを見分けなくても守れるように、かけ直しの前に全員分を保存し、あとで「保存した文が今どの枠にも無い」ものだけを空いている枠に戻す。
    -Save    … 保存。backup\comment_keep_<日時>.csv (コメント) と backup\kensa_keep_<日時>.csv (全項目の値・判定・印) を作る。DB は変えない。
    -Restore … 保存ファイルを選び、(1) 判定が変わった・消えた項目の一覧 (読むだけ) と (2) 戻すコメントのプレビューを出す。y で書く。
  ・同じ文が既にどこかの枠にあれば戻さない (「既にあり」)。
  ・同じ種類の文 (先頭の【○○】が同じ、または 腎機能 / 眼底 / 糖尿病 を含む) が既にあれば、勝手に戻さず「似た文あり」で一覧に出すだけ。
  ・戻した枠の元の文章は backup\kukuri_comment_restore_<日時>.csv に残す。kukuri_comment_undo.bat で戻せる。
  順番 (9/9 の例):
    1. comment_keep_99_save.bat (保存)   2. 健診ナビで自動判定   3. 判定取込・眼底取込・kukuri_fix_99・kukuri_comment_99
    4. comment_keep_restore.bat (消えたコメントを戻す)
  例:
    comment_keep.ps1 -Save -Dates 2026/09/09 -UkeFrom 4631 -UkeTo 4831
    comment_keep.ps1 -Restore ask
#>
[CmdletBinding()]
param(
    [string]$Dates,
    [string]$UkeNo,
    [int]$UkeFrom = 0,
    [int]$UkeTo = 0,
    [switch]$Save,
    [string]$Restore,
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
$SLOTS = @('11001','11002','11003','11004','11005','11006','11007','11008','11009')
$SLOT_LIT = "'" + ($SLOTS -join "','") + "'"

# 文の種類 (似た文の見分け): 先頭の【○○】、無ければ決まった言葉
function Kind-Of([string]$t) {
    if ($t -match '^\s*【([^】]+)】') { return '【' + $Matches[1] + '】' }
    foreach ($w in @('腎機能','眼底','糖尿病','中性脂肪','血圧','肝機能','貧血','心電図','胸部','胃')) { if ($t -like "*$w*") { return $w } }
    return ''
}

if (-not $Save -and -not $Restore) { throw '-Save か -Restore を指定してください。' }

# ============================================================
# 保存 (-Save)
# ============================================================
if ($Save) {
    $dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
    $ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })
    if ($dateList.Count -eq 0) { throw '-Dates に受診日を指定してください (例 2026/09/09)' }
    $conn = Open-Db
    try {
        $dLit = "'" + ($dateList -join "','") + "'"
        $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit)"
        if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
        if ($UkeFrom -gt 0) { $where += " AND s.UKE_NO_KENSA >= $UkeFrom" }
        if ($UkeTo -gt 0)   { $where += " AND s.UKE_NO_KENSA <= $UkeTo" }
        $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, ISNULL(k.KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC,
       LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(k.HL,''))) AS HL
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $where
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA, k.KOMOKU_CD
"@).Rows
        if ($rows.Count -eq 0) { Write-Host '対象の人がいません。'; return }
        $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
        $fC = Join-Path $BackupDir ("comment_keep_{0}.csv" -f $ts)
        $fK = Join-Path $BackupDir ("kensa_keep_{0}.csv" -f $ts)
        $all = @($rows | ForEach-Object { New-Object PSObject -Property ([ordered]@{ PK_SEQ = [string]$_.PK_SEQ; 受付番号 = [string]$_.UKE; 受診日 = [string]$_.YMD; 氏名 = [string]$_.NAME
                    KOMOKU_CD = [string]$_.CD; 値 = (Normalize-Text ([string]$_.V)); 結果CD = [string]$_.VC; 判定 = [string]$_.H; 印 = [string]$_.HL }) })
        $cm = @($all | Where-Object { $SLOTS -contains $_.KOMOKU_CD -and $_.値 -ne '' -and $_.値 -ne '#' })
        $all | Export-Csv -Path $fK -NoTypeInformation -Encoding UTF8
        $cm  | Export-Csv -Path $fC -NoTypeInformation -Encoding UTF8
        $np = @($all | Group-Object PK_SEQ).Count
        Write-Host ''
        Write-Host ("[保存しました] {0} 人 / コメント {1} 件 / 全項目 {2} 行" -f $np, $cm.Count, $all.Count) -ForegroundColor Green
        Write-Host ("  コメント : {0}" -f $fC)
        Write-Host ("  全項目   : {0}" -f $fK)
        Write-Host '  このあと健診ナビで自動判定 → 判定取込・眼底取込・kukuri_fix・kukuri_comment → comment_keep_restore.bat の順に進めてください。' -ForegroundColor Yellow
    } finally { $conn.Close() }
    return
}

# ============================================================
# 戻す (-Restore)
# ============================================================
if ($Restore -eq 'ask') {
    $files = @(Get-ChildItem -Path $BackupDir -Filter 'comment_keep_*.csv' -File | Sort-Object LastWriteTime -Descending)
    if ($files.Count -eq 0) { Write-Host '保存ファイル (backup\comment_keep_*.csv) がありません。先に保存してください。'; return }
    for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
    $a = Read-Host '使う保存ファイルの番号 (それ以外 = やめる)'
    if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt $files.Count) { Write-Host 'やめました。'; return }
    $Restore = $files[[int]$a - 1].FullName
}
if (-not (Test-Path $Restore)) { throw "保存ファイルがありません: $Restore" }
$saved = @(Import-Csv -Path $Restore -Encoding UTF8)
$kFile = $Restore -replace 'comment_keep_', 'kensa_keep_'
$savedK = if (Test-Path $kFile) { @(Import-Csv -Path $kFile -Encoding UTF8) } else { @() }
$pks = @(@($saved | ForEach-Object { $_.PK_SEQ }) + @($savedK | ForEach-Object { $_.PK_SEQ }) | Sort-Object -Unique)
if ($pks.Count -eq 0) { Write-Host '保存ファイルに人がいません。'; return }

$conn = Open-Db
try {
    $pkLit = $pks -join ','
    $now = @{}   # pk -> cd -> @{ V; VC; H; HL }
    foreach ($r in (Invoke-DbQuery $conn "SELECT PK_SEQ, LTRIM(RTRIM(KOMOKU_CD)) AS CD, ISNULL(KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(KEKKA_CD,''))) AS VC, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(HL,''))) AS HL FROM T_KENSA WHERE PK_SEQ IN ($pkLit)").Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $now.ContainsKey($pk)) { $now[$pk] = @{} }
        $now[$pk][[string]$r.CD] = @{ V = (Normalize-Text ([string]$r.V)); VC = [string]$r.VC; H = [string]$r.H; HL = [string]$r.HL }
    }
    $names = @{}
    foreach ($r in (Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(KOMOKU_CD)) AS CD, MAX(MEISYO1) AS NM FROM T_KOMOKU GROUP BY LTRIM(RTRIM(KOMOKU_CD))").Rows) { $names[[string]$r.CD] = [string]$r.NM }

    # (1) 判定・値が変わった項目 (読むだけ)
    if ($savedK.Count -gt 0) {
        $diff = @()
        foreach ($s in $savedK) {
            if ($SLOTS -contains $s.KOMOKU_CD) { continue }
            $n = if ($now.ContainsKey($s.PK_SEQ) -and $now[$s.PK_SEQ].ContainsKey($s.KOMOKU_CD)) { $now[$s.PK_SEQ][$s.KOMOKU_CD] } else { $null }
            $nh = if ($n) { $n.H } else { '(枠なし)' }
            $nv = if ($n) { $n.V } else { '' }
            if ($s.判定 -ne $nh -or ($s.値 -ne '' -and $s.値 -ne $nv)) {
                $diff += New-Object PSObject -Property @{ 受付番号 = $s.受付番号; 氏名 = $s.氏名; 項目CD = $s.KOMOKU_CD; 項目名 = $names[$s.KOMOKU_CD]
                    前の判定 = $s.判定; 今の判定 = $nh; 前の値 = $(if ($s.値.Length -gt 20) { $s.値.Substring(0,20) } else { $s.値 }); 今の値 = $(if ($nv.Length -gt 20) { $nv.Substring(0,20) } else { $nv }) }
            }
        }
        Write-Host ''
        Write-Host ("=== (1) 保存したときから 判定・値が変わった項目: {0} か所 (読むだけ。消えた判定は取込ツール・kukuri_fix で入れ直す) ===" -f $diff.Count) -ForegroundColor Cyan
        if ($diff.Count -gt 0) {
            Write-Host '  項目ごとの件数:'
            $diff | Group-Object 項目CD | Sort-Object Count -Descending | ForEach-Object { Write-Host ("    {0,-9} {1,-20} {2} か所" -f $_.Name, $names[$_.Name], $_.Count) }
            $diff | Sort-Object { [int]$_.受付番号 }, 項目CD | Select-Object 受付番号, 氏名, 項目CD, 項目名, 前の判定, 今の判定, 前の値, 今の値 | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
        }
    }

    # (2) 戻すコメント
    $plan = @(); $used = @{}
    foreach ($grp in ($saved | Group-Object PK_SEQ)) {
        $pk = $grp.Name
        if (-not $now.ContainsKey($pk)) { continue }
        $cur = $now[$pk]
        if (-not $used.ContainsKey($pk)) { $used[$pk] = @{} }
        $curTexts = @($SLOTS | Where-Object { $cur.ContainsKey($_) } | ForEach-Object { $cur[$_].V } | Where-Object { $_ -ne '' -and $_ -ne '#' })
        foreach ($s in ($grp.Group | Sort-Object KOMOKU_CD)) {
            $t = $s.値
            $rep = New-Object PSObject -Property @{ PkSeq = $pk; 受付番号 = $s.受付番号; 氏名 = $s.氏名; 文章 = $t; 番号 = $s.結果CD; 枠 = ''; 状態 = ''; 今の文章 = ''; 今の番号 = ''; 似た文 = '' }
            if ($curTexts -contains $t) { $rep.状態 = '既にあり'; $plan += $rep; continue }
            $k = Kind-Of $t
            if ($k -ne '') {
                $sim = @($curTexts | Where-Object { (Kind-Of $_) -eq $k })
                if ($sim.Count -gt 0) { $rep.状態 = '似た文あり(戻さない)'; $rep.似た文 = $sim[0]; $plan += $rep; continue }
            }
            $slot = $null
            foreach ($sc in $SLOTS) {
                if (-not $cur.ContainsKey($sc) -or $used[$pk].ContainsKey($sc)) { continue }
                if ($cur[$sc].V -eq '' -or $cur[$sc].V -eq '#') { $slot = $sc; break }
            }
            if (-not $slot) { $rep.状態 = '空いている枠がありません'; $plan += $rep; continue }
            $used[$pk][$slot] = 1
            $rep.枠 = $slot; $rep.今の文章 = $cur[$slot].V; $rep.今の番号 = $cur[$slot].VC; $rep.状態 = '戻す'
            $plan += $rep
        }
    }
    $back = @($plan | Where-Object { $_.状態 -eq '戻す' })
    $sim  = @($plan | Where-Object { $_.状態 -like '似た文あり*' })
    $full = @($plan | Where-Object { $_.状態 -eq '空いている枠がありません' })
    $same = @($plan | Where-Object { $_.状態 -eq '既にあり' })
    Write-Host ''
    Write-Host '=== (2) 消えていたコメント ===' -ForegroundColor Cyan
    @($plan | Where-Object { $_.状態 -ne '既にあり' }) | Sort-Object { [int]$_.受付番号 } | Select-Object 受付番号, 氏名, 状態, 枠, 文章, 似た文 | Format-Table -AutoSize -Wrap | Out-String -Width 240 | Write-Host
    Line
    Write-Host ("戻す: {0} 件 / 似た文あり (戻さない・確認): {1} 件 / 枠が足りない: {2} 件 / 既にあり: {3} 件" -f $back.Count, $sim.Count, $full.Count, $same.Count) -ForegroundColor $(if ($sim.Count + $full.Count -gt 0) { 'Yellow' } else { 'Green' })
    Line
    if ($back.Count -eq 0) { Write-Host '戻すコメントはありません。'; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。「似た文あり」「枠が足りない」は画面で見てください。' -ForegroundColor Yellow
        $a = Read-Host '「戻す」の分を書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    $undoFile = Join-Path $BackupDir ("kukuri_comment_restore_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($back | Group-Object PkSeq)) {
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
                $q = Invoke-DbQuery $conn "SELECT ISNULL(KEKKA,'') AS V FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd" @{ p = $pk; cd = $t.枠 } $tran
                if ($q.Rows.Count -ne 1 -or (Normalize-Text ([string]$q.Rows[0].V)) -ne $t.今の文章) { Write-Host ("  [変わっていたので触らず] {0} {1}" -f $t.氏名, $t.枠) -ForegroundColor Yellow; continue }
                $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET KEKKA = @v, KEKKA_CD = @c WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ v = $t.文章; c = $t.番号; p = $pk; cd = $t.枠 }
                if ($n -ne 1) { throw ("UPDATE 影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.枠) }
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.枠; 前の文章 = $t.今の文章; 前の番号 = $t.今の番号; 後の文章 = $t.文章; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [戻した] {0} ({1}) : {2} 件" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $grp.Count) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) { $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 前の文章, 前の番号, 後の文章, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} 件を戻しました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0} → 画面を閉じてもう一度実行" -f ($skipped -join ', ')) -ForegroundColor Red }
}
finally { $conn.Close() }
