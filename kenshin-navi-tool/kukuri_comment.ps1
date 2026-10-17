<#
.SYNOPSIS
  HbA1c・eGFR・中性脂肪 の判定 (D/F) に応じた定型文を医師指示事項に足す (kukuri_comment.ps1)

.DESCRIPTION
  後から枠を足した HbA1c・eGFR は自動判定のときに無かったので、判定コメント (医師指示事項 11001〜11009) が付いていない。
  自動判定が同じ場面で入れている文章を、空いている枠にだけ足す。kukuri_fix でくくり判定を直したあとに使う。
    HbA1c 068689  D → 【糖代謝】　高値です。経過観察の必要があります。
                  F → 【糖代謝】　糖尿病が疑われますので精密検査の必要があります。
    eGFR  068676  D → 腎機能検査に軽度異常があります。経過観察し、定期的に検査して下さい。 (番号 159)
                  F → 腎機能検査に異常があります。精査を受けて下さい。 (番号 161)
    中性脂肪 R06-0002 (空腹時 068702 / 随時 R06-0003 に判定がある人はそちらを見る)
                  D → 【中性脂肪】　高値です。飲酒に注意しバランスの良い食事と適度な運動を心がけ、経過観察の必要があります。
                  F → 【中性脂肪】　高値です。精密検査の必要があります。
  ・判定が A/B の人には何もしない (自動判定もコメントを付けない)。
  ・既に同じ種類のコメント (糖 / 腎機能) がある人:
      重さが合っていれば「既にあり」で何もしない。軽い文 (経過観察) しか無いのに判定が F なら「文章を直す」でその枠を差し替える。
    番号 (KEKKA_CD) は自動判定のかけ直しで文章とずれていることがあるので、照合は文章で行う。
  ・総合判定が付いていない人 (自動判定前) は対象外。
  ・必ずプレビューを出し、y を押したときだけ書く。書く前に本人の T_KENSA を backup\ に保存し、
    書いた枠の元の文章を backup\kukuri_comment_<日時>.csv に残す。-Undo でその文章に戻せる。
  例:
    kukuri_comment.ps1 -Dates 2026/09/03 -UkeFrom 4631 -UkeTo 4831
    kukuri_comment.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
    kukuri_comment.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$Dates,
    [string]$UkeNo,
    [int]$UkeFrom = 0,
    [int]$UkeTo = 0,
    [switch]$Commit,
    [string]$Undo,
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

# ---- 入れる文章 ----
#   Item … 判定を見る項目 / Kind … 同じ種類のコメントを見分ける文字 / Heavy … F のときの文に含まれる言葉 (重さの判定用)
$RULES = @(
    @{ Name = 'HbA1c'; Item = '068689'; Kind = '糖';    KindPat = '糖代謝|糖尿病';   # 「糖」だけだと【尿糖】の文に当たって「既にあり」になる (2026/10/17 に発覚)
       D = @{ Cd = '';    Text = '【糖代謝】　高値です。経過観察の必要があります。' }
       F = @{ Cd = '';    Text = '【糖代謝】　糖尿病が疑われますので精密検査の必要があります。' } }
    @{ Name = 'eGFR';  Item = '068676'; Kind = '腎機能'; KindPat = '腎機能';
       D = @{ Cd = '159'; Text = '腎機能検査に軽度異常があります。経過観察し、定期的に検査して下さい。' }
       F = @{ Cd = '161'; Text = '腎機能検査に異常があります。精査を受けて下さい。' } }
    @{ Name = '中性脂肪'; Item = 'R06-0002'; Alt = @('068702','R06-0003'); Kind = '中性脂肪'; KindPat = '中性脂肪';
       D = @{ Cd = '';    Text = '【中性脂肪】　高値です。飲酒に注意しバランスの良い食事と適度な運動を心がけ、経過観察の必要があります。' }
       F = @{ Cd = '';    Text = '【中性脂肪】　高値です。精密検査の必要があります。' } }
    # 眼底: 【判定】眼底 600060 (業者の判定を mapping_gantei_hantei_* で入れたもの) が D/F の人。文章は 2026 年に使われていた定型文 (番号 235/236)。
    #       先生の言い回しに変えるときはこの 2 行の Text を直す (Cd は番号。分からなければ空でよい)。
    @{ Name = '眼底';  Item = '600060'; Kind = '眼底'; KindPat = '眼底';
       D = @{ Cd = '235'; Text = '眼底に所見を認めますので、経過観察が必要です。症状があれば眼科を受診して下さい。' }
       F = @{ Cd = '236'; Text = '眼底に異常があります。眼科受診し、精密検査を受けて下さい。' } }
)
# 項目CD → 規則 (Alt の項目も同じ規則に向ける)
$RULE_OF = @{}
foreach ($r in $RULES) { $RULE_OF[$r.Item] = $r; if ($r.Alt) { foreach ($a in $r.Alt) { $RULE_OF[$a] = $r } } }
$HEAVY = '精密|精査|治療'      # これを含む文は F 相当
$SLOTS = @('11001','11002','11003','11004','11005','11006','11007','11008','11009')
$SOGO  = '11000'

# ============================================================
# 元に戻す (-Undo)
# ============================================================
if ($Undo) {
    if ($Undo -eq 'ask') {
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'kukuri_comment_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\kukuri_comment_*.csv) がありません。'; return }
        for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
        $a = Read-Host '戻す記録の番号 (それ以外 = やめる)'
        if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt $files.Count) { Write-Host 'やめました。'; return }
        $Undo = $files[[int]$a - 1].FullName
    }
    if (-not (Test-Path $Undo)) { throw "記録ファイルがありません: $Undo" }
    $rows = @(Import-Csv -Path $Undo -Encoding UTF8)
    Write-Host ("この記録で書いた文章 {0} か所を元に戻します (今の文章が記録の「後」と違う枠は触りません): {1}" -f $rows.Count, $Undo) -ForegroundColor Yellow
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
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KENSA SET KEKKA = @oldv, KEKKA_CD = @oldc WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(KEKKA,''))) = @newv" `
                        @{ oldv = [string]$r.前の文章; oldc = [string]$r.前の番号; p = $pk; cd = [string]$r.KOMOKU_CD; newv = [string]$r.後の文章 }
                    if ($n -eq 1) { $back++ } else { $kept++ }
                }
                $tran.Commit()
            } catch { $tran.Rollback(); throw }
        }
        Write-Host ("[完了] 戻した枠 {0} / 文章が変わっていたので触らなかった枠 {1}" -f $back, $kept) -ForegroundColor Green
    } finally { $conn.Close() }
    return
}

# ============================================================
# 足す
# ============================================================
$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
$ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })
if ($dateList.Count -eq 0) { throw '-Dates に受診日を指定してください (例 2026/09/03)' }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' HbA1c・eGFR・中性脂肪 の判定に応じた定型文を 医師指示事項 に足す' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  受診日    : {0}" -f ($dateList -join ', '))
if ($ukeList.Count -gt 0) { Write-Host ("  受付番号  : {0}" -f ($ukeList -join ', ')) }
if ($UkeFrom -gt 0 -or $UkeTo -gt 0) { Write-Host ("  受付番号の範囲: {0}〜{1}" -f $UkeFrom, $UkeTo) }
Write-Host '  入れる先: 総合判定N(編集) 11001〜11009 の空いている枠。総合判定が付いていない人は対象外。' -ForegroundColor DarkGray
foreach ($r in $RULES) {
    Write-Host ('    {0,-6} D: {1}' -f $r.Name, $r.D.Text)
    Write-Host ('    {0,-6} F: {1}' -f '',      $r.F.Text)
}
Write-Host ''

$conn = Open-Db
try {
    $dLit = "'" + ($dateList -join "','") + "'"
    $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit)"
    if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
    if ($UkeFrom -gt 0) { $where += " AND s.UKE_NO_KENSA >= $UkeFrom" }
    if ($UkeTo -gt 0)   { $where += " AND s.UKE_NO_KENSA <= $UkeTo" }
    $itemLit = "'" + (@($RULE_OF.Keys) -join "','") + "'"
    $slotLit = "'" + ($SLOTS -join "','") + "'"

    # 対象者 (総合判定あり) と項目の判定
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($itemLit)
WHERE $where
  AND ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(x.HANTEI_KIGO,''))) FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = '$SOGO'),'') <> ''
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('D','F','G')
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA, k.KOMOKU_CD
"@).Rows
    if ($rows.Count -eq 0) { Write-Host '  判定が D/F の人がいません。入れるものがありません。' -ForegroundColor Yellow; return }
    Write-Host ("  判定 D/F の行: {0} ({1} 人)" -f $rows.Count, @($rows | Group-Object PK_SEQ).Count)

    # いまの医師指示事項
    $pkLit = (($rows | ForEach-Object { [string]$_.PK_SEQ }) | Sort-Object -Unique) -join ','
    $cur = @{}
    foreach ($r in (Invoke-DbQuery $conn "SELECT PK_SEQ, LTRIM(RTRIM(KOMOKU_CD)) AS CD, ISNULL(KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(KEKKA_CD,''))) AS VC FROM T_KENSA WHERE PK_SEQ IN ($pkLit) AND LTRIM(RTRIM(KOMOKU_CD)) IN ($slotLit)").Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $cur.ContainsKey($pk)) { $cur[$pk] = @{} }
        $cur[$pk][[string]$r.CD] = @{ Kekka = (Normalize-Text ([string]$r.V)); KekkaCd = [string]$r.VC }
    }

    # 同じ人・同じ規則 (例: 中性脂肪の未判別と空腹時) が 2 行あれば重い方だけ残す
    $best = @{}
    foreach ($r in $rows) {
        $key = [string]$r.PK_SEQ + '|' + $RULE_OF[[string]$r.CD].Name
        $rk = switch (([string]$r.H).ToUpper()) { 'D' {4} 'F' {6} 'G' {7} default {0} }
        if (-not $best.ContainsKey($key) -or $rk -gt $best[$key].RK) { $best[$key] = @{ RK = $rk; Row = $r } }
    }
    $rows = @($best.Values | ForEach-Object { $_.Row } | Sort-Object { [string]$_.YMD }, { [int]$_.UKE }, { [string]$_.CD })

    # 計画
    $plan = @(); $used = @{}
    foreach ($r in $rows) {
        $pk = [string]$r.PK_SEQ; $cd = [string]$r.CD; $h = ([string]$r.H).ToUpper()
        $rule = $RULE_OF[$cd]
        $t = if ($h -eq 'D') { $rule.D } else { $rule.F }
        $rep = New-Object PSObject -Property @{ PkSeq = $pk; 受診日 = [string]$r.YMD; 受付番号 = [string]$r.UKE; 氏名 = [string]$r.NAME
            検査 = $rule.Name; 値 = [string]$r.V; 判定 = $h; 枠 = ''; 状態 = ''; 文章 = $t.Text; 番号CD = $t.Cd; 今の文章 = ''; 今の番号 = '' }
        if (-not $cur.ContainsKey($pk)) { $rep.状態 = '医師指示事項の枠がありません'; $plan += $rep; continue }
        if (-not $used.ContainsKey($pk)) { $used[$pk] = @{} }

        # 同じ種類 (糖 / 腎機能) のコメントが既にあるか (文章で見る)
        $kindSlot = $null
        foreach ($sc in $SLOTS) {
            if (-not $cur[$pk].ContainsKey($sc)) { continue }
            $v = $cur[$pk][$sc].Kekka
            if ($v -ne '' -and $v -match $rule.KindPat) { $kindSlot = $sc; break }
        }
        if ($kindSlot) {
            $v = $cur[$pk][$kindSlot].Kekka
            $isHeavy = ($v -match $HEAVY)
            if ($v -eq $t.Text -or ($h -eq 'D') -or ($h -ne 'D' -and $isHeavy)) {
                # D なら経過観察でも精検でも既にある扱い / F で既に精検・治療の文があれば既にある扱い
                $rep.枠 = $kindSlot; $rep.今の文章 = $v; $rep.状態 = '既にあり'; $plan += $rep; continue
            }
            # F なのに軽い文しか無い → その枠を差し替える
            $rep.枠 = $kindSlot; $rep.今の文章 = $v; $rep.今の番号 = $cur[$pk][$kindSlot].KekkaCd; $rep.状態 = '文章を直す'
            $used[$pk][$kindSlot] = 1
            $plan += $rep; continue
        }
        # 空いている枠
        $slot = $null
        foreach ($sc in $SLOTS) {
            if (-not $cur[$pk].ContainsKey($sc)) { continue }
            if ($used[$pk].ContainsKey($sc)) { continue }
            $v = $cur[$pk][$sc].Kekka
            if ($v -eq '' -or $v -eq '#') { $slot = $sc; break }
        }
        if (-not $slot) { $rep.状態 = '空いている枠がありません'; $plan += $rep; continue }
        $used[$pk][$slot] = 1
        $rep.枠 = $slot; $rep.今の文章 = $cur[$pk][$slot].Kekka; $rep.今の番号 = $cur[$pk][$slot].KekkaCd; $rep.状態 = 'OK'
        $plan += $rep
    }

    Write-Host ''
    Write-Host '=== 追加する内容 ===' -ForegroundColor Cyan
    $plan | Select-Object 受診日, 受付番号, 氏名, 検査, 値, 判定, 枠, 状態, 今の文章 | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
    $ok   = @($plan | Where-Object { $_.状態 -eq 'OK' -or $_.状態 -eq '文章を直す' })
    $add  = @($plan | Where-Object { $_.状態 -eq 'OK' })
    $fix  = @($plan | Where-Object { $_.状態 -eq '文章を直す' })
    $skip = @($plan | Where-Object { $_.状態 -eq '既にあり' })
    $err  = @($plan | Where-Object { $_.状態 -notin @('OK','文章を直す','既にあり') })
    Write-Host '=== 入れる文章 ===' -ForegroundColor Cyan
    foreach ($t in ($ok | Sort-Object 文章 -Unique)) {
        $n = @($ok | Where-Object { $_.文章 -eq $t.文章 }).Count
        Write-Host ("  ({0}件)  {1}" -f $n, $t.文章)
    }
    Line
    Write-Host ("追加: {0} 件 / 文章を直す: {1} 件 / 既にあり: {2} 件 / 要確認: {3} 件" -f $add.Count, $fix.Count, $skip.Count, $err.Count) -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })
    Line
    if ($ok.Count -eq 0) { Write-Host '追加するものがありません。'; return }
    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ まだ何も書き込んでいません。「文章を直す」は今の文章を消して置き換えます。' -ForegroundColor Yellow
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }

    # 書く
    $undoFile = Join-Path $BackupDir ("kukuri_comment_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($ok | Group-Object PkSeq)) {
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
                $now = Invoke-DbQuery $conn "SELECT ISNULL(KEKKA,'') AS V FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd" @{ p = $pk; cd = $t.枠 } $tran
                if ($now.Rows.Count -ne 1 -or (Normalize-Text ([string]$now.Rows[0].V)) -ne $t.今の文章) { Write-Host ("  [変わっていたので触らず] {0} {1}" -f $t.氏名, $t.枠) -ForegroundColor Yellow; continue }
                $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET KEKKA = @v, KEKKA_CD = @c WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ v = $t.文章; c = $t.番号CD; p = $pk; cd = $t.枠 }
                if ($n -ne 1) { throw ("UPDATE 影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.枠) }
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.枠; 前の文章 = $t.今の文章; 前の番号 = $t.今の番号; 後の文章 = $t.文章; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [書込] {0} ({1}) : {2} 件" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $grp.Count) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) { $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 前の文章, 前の番号, 後の文章, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} 件を医師指示事項に入れました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0} → 画面を閉じてもう一度実行" -f ($skipped -join ', ')) -ForegroundColor Red }
    Write-Host '   元に戻すときは kukuri_comment_undo.bat。このあと自動判定をかけ直すと文章が消える可能性があります。' -ForegroundColor Yellow
}
finally { $conn.Close() }
