<#
.SYNOPSIS
  問診の服薬歴 (はい) から 高血圧・糖尿病・高脂血症 (薬剤治療中) を既往歴 (T_KIOU) に入れる (kiou_med.ps1)

.DESCRIPTION
  健診ナビは既往歴 (T_KIOU) と「病名 → 判定」の対応表 (T_KIOU_HANTEI) を見て、自動判定のときに
  高血圧 → 血圧、糖尿病 → 糖代謝、高脂血症 → 脂質代謝 を J (治療中) などにする。石屋は既往歴が入っていない
  (198 人が「特記事項なし」だけ) ので、服薬中でも J が付かなかった (check106〜109)。
  問診 084001 服薬歴(血圧) / 084002 (血糖) / 084003 (脂質) が「はい」の人に、
  高血圧 (病名CD 19) / 糖尿病 (48) / 高脂血症 (47) を 治療 5 = 薬剤治療中 で入れる (2026/10/18 ユーザー決定。業者ファイルの既往歴は
  2 枠しかなく糖尿病・高脂血症が漏れるため。業者ファイルで既往歴に載っている人は全員この形だった)。
  ・本人の既往歴として入れる。最初の 1 件は「特記事項なし」の行を置き換え、2 件目は行を足す。
  ・同じ病名が既にあれば入れない。本人の既往歴は合計 -Max 件 (既定 4) まで。
  ・病名の入った行は消さない・書き換えない。
  ・プレビュー → y で書く。書く前に本人の T_KIOU を backup\ に保存し、変更を backup\kiou_med_<日時>.csv に残す。-Undo で戻せる。
  ・入れたあと自動判定をかけると J などが付く (かけなければ判定は変わらず、結果票の既往歴欄に出るだけ)。
  例:
    kiou_med.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10
    kiou_med.ps1 -Dates 2026/09/09 -UkeNo 4691
    kiou_med.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$Dates = '2026/09/03,2026/09/09,2026/09/10',
    [string]$UkeNo,
    [int]$UkeFrom = 4631,
    [int]$UkeTo = 4831,
    [int]$Max = 4,
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
Invoke-Expression (Get-Part 'function Normalize-Text'     'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile'  'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Test-Locked'        'function Commit-Plan')
$BackupDir = Join-Path $PSScriptRoot 'backup'
if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
function Split-List([string]$s) { return @(($s -split '[,、\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }
$PLACEHOLDER = @('', '特記事項なし', '特記事項無し', '特になし', '特に無し', 'なし', '無し')
function Test-Placeholder([string]$s) { return ($PLACEHOLDER -contains ((Normalize-Text $s) -replace '[\s　]', '')) }
function DbV($v) { if ($null -eq $v -or [string]$v -eq '') { return [DBNull]::Value } else { return $v } }

# ---- コード表 ----
$codeFile = Join-Path $PSScriptRoot 'form\kiou_codes.csv'
if (-not (Test-Path $codeFile)) { throw "form\kiou_codes.csv がありません。" }
$BY = @{}; $CH = @{}
foreach ($l in (Get-Content $codeFile -Encoding UTF8)) {
    if ($l -match '^\s*#' -or $l -match '^Kind,') { continue }
    $f = $l -split ','
    if ($f.Count -lt 3) { continue }
    if ($f[0] -eq 'BYOMEI') { $BY[[int]$f[1]] = $f[2] } elseif ($f[0] -eq 'CHIRYO') { $CH[[int]$f[1]] = $f[2] }
}

# ============================================================
# 元に戻す
# ============================================================
if ($Undo) {
    if ($Undo -eq 'ask') {
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'kiou_med_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\kiou_med_*.csv) がありません。'; return }
        for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
        $a = Read-Host '戻す記録の番号 (それ以外 = やめる)'
        if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt $files.Count) { Write-Host 'やめました。'; return }
        $Undo = $files[[int]$a - 1].FullName
    }
    $rows = @(Import-Csv -Path $Undo -Encoding UTF8)
    Write-Host ("この記録で入れた既往歴 {0} 行を元に戻します (追加した行は消し、置き換えた行は元の内容に戻す): {1}" -f $rows.Count, $Undo) -ForegroundColor Yellow
    $a = Read-Host '戻しますか (y = 戻す / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。'; return }
    $conn = Open-Db
    try {
        $back = 0; $kept = 0
        foreach ($r in $rows) {
            $tran = $conn.BeginTransaction()
            try {
                Acquire-AppLock $conn $tran ("KIOU_IMPORT:" + $r.KOJIN_ID)
                if ($r.操作 -eq '追加') {
                    $n = Invoke-DbExec $conn $tran "DELETE FROM T_KIOU WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @b" @{ k = $r.KOJIN_ID; r = $r.RENBAN; b = $r.病名 }
                } else {
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KIOU SET BYOMEI = @ob, BYOMEI_CD = @oc, CHIRYO = @och, CHIRYOMEI = @ochm, AGE = @oa WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @b" `
                        @{ ob = $r.前の病名; oc = $r.前の病名CD; och = $r.前の治療; ochm = $r.前の治療名; oa = (DbV $r.前の年齢); k = $r.KOJIN_ID; r = $r.RENBAN; b = $r.病名 }
                }
                $tran.Commit()
                if ($n -eq 1) { $back++ } else { $kept++ }
            } catch { $tran.Rollback(); throw }
        }
        Write-Host ("[完了] 戻した行 {0} / 中身が変わっていたので触らなかった行 {1}" -f $back, $kept) -ForegroundColor Green
    } finally { $conn.Close() }
    return
}

# ============================================================
# 入れる
# ============================================================
$MED = @(
    @{ Q = '084001'; Cd = '19'; Name = '高血圧' },
    @{ Q = '084002'; Cd = '48'; Name = '糖尿病' },
    @{ Q = '084003'; Cd = '47'; Name = '高脂血症' }
)
foreach ($m in $MED) { if (-not $BY.ContainsKey([int]$m.Cd) -or $BY[[int]$m.Cd] -ne $m.Name) { throw ("form\kiou_codes.csv の病名コード {0} が {1} ではありません。" -f $m.Cd, $m.Name) } }
$CHIRYO = '5'; $CHIRYOMEI = $CH[5]
$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
$ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 服薬歴 → 既往歴 (高血圧・糖尿病・高脂血症 / 薬剤治療中、本人)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  受診日: {0}  受付番号 {1}〜{2}{3}  1 人 {4} 件まで" -f ($dateList -join ', '), $UkeFrom, $UkeTo, $(if ($ukeList.Count -gt 0) { ' / ' + ($ukeList -join ',') } else { '' }), $Max)

$conn = Open-Db
try {
    $dLit = "'" + ($dateList -join "','") + "'"
    $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit) AND s.UKE_NO_KENSA BETWEEN $UkeFrom AND $UkeTo"
    if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
    $qLit = "'" + (($MED | ForEach-Object { $_.Q }) -join "','") + "'"
    $ans = Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, s.KOJIN_ID, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS Q, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($qLit)
WHERE $where
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@
    $selfMei = '本人'
    $z = Invoke-DbQuery $conn "SELECT TOP 1 LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) AS M, COUNT(*) AS C FROM T_KIOU WHERE LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) = '0' GROUP BY LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) ORDER BY 2 DESC"
    if ($z.Rows.Count -gt 0 -and [string]$z.Rows[0].M -ne '') { $selfMei = [string]$z.Rows[0].M }

    $who = @{}
    foreach ($r in $ans.Rows) { $who[[string]$r.KOJIN_ID] = @{ Pk = [string]$r.PK_SEQ; Uke = [string]$r.UKE; Ymd = [string]$r.YMD; Name = [string]$r.NAME } }
    $exist = @{}
    if ($who.Count -gt 0) {
        $kLit = ($who.Keys | Sort-Object -Unique) -join ','
        foreach ($r in (Invoke-DbQuery $conn "SELECT KOJIN_ID, RENBAN, LTRIM(RTRIM(ISNULL(BYOMEI,''))) AS B, LTRIM(RTRIM(ISNULL(BYOMEI_CD,''))) AS BC, LTRIM(RTRIM(ISNULL(CHIRYO,''))) AS C, LTRIM(RTRIM(ISNULL(CHIRYOMEI,''))) AS CM, AGE, LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) AS ZK FROM T_KIOU WHERE KOJIN_ID IN ($kLit) ORDER BY KOJIN_ID, RENBAN").Rows) {
            $kj = [string]$r.KOJIN_ID
            if (-not $exist.ContainsKey($kj)) { $exist[$kj] = @{ MaxR = 0; Self = @{}; Count = 0; Holder = $null } }
            $rn = [int]$r.RENBAN; if ($rn -gt $exist[$kj].MaxR) { $exist[$kj].MaxR = $rn }
            if ([string]$r.ZK -eq '' -or [string]$r.ZK -eq '0') {
                if (Test-Placeholder ([string]$r.B)) {
                    if ($null -eq $exist[$kj].Holder) { $exist[$kj].Holder = @{ R = $rn; B = [string]$r.B; BC = [string]$r.BC; C = [string]$r.C; CM = [string]$r.CM; A = $(if ($r.AGE -eq [DBNull]::Value) { '' } else { [string]$r.AGE }) } }
                } else { $exist[$kj].Self[[string]$r.B] = 1; $exist[$kj].Count++ }
            }
        }
    }

    $plan = @()
    foreach ($r in $ans.Rows) {
        $yes = ([string]$r.V -eq 'はい' -or [string]$r.VC -eq '1')
        if (-not $yes) { continue }
        $m = @($MED | Where-Object { $_.Q -eq [string]$r.Q })[0]
        $kj = [string]$r.KOJIN_ID; $p = $who[$kj]
        if (-not $exist.ContainsKey($kj)) { $exist[$kj] = @{ MaxR = 0; Self = @{}; Count = 0; Holder = $null } }
        $e = $exist[$kj]
        $rep = New-Object PSObject -Property @{ 受診日 = $p.Ymd; 受付番号 = $p.Uke; 氏名 = $p.Name; PkSeq = $p.Pk; KOJIN_ID = $kj
            病名CD = $m.Cd; 病名 = $m.Name; 年齢 = ''; 治療 = $CHIRYO; 治療名 = $CHIRYOMEI; 操作 = ''; 連番 = ''; 前 = $null; 状態 = '' }
        if ($e.Self.ContainsKey($m.Name)) { $rep.状態 = '既に入っています'; $plan += $rep; continue }
        if ($e.Count -ge $Max) { $rep.状態 = ("本人の既往歴が既に {0} 件" -f $Max); $plan += $rep; continue }
        if ($null -ne $e.Holder) { $rep.操作 = '置換'; $rep.連番 = [string]$e.Holder.R; $rep.前 = $e.Holder; $e.Holder = $null }
        else { $e.MaxR++; $rep.操作 = '追加'; $rep.連番 = [string]$e.MaxR }
        $e.Self[$m.Name] = 1; $e.Count++
        $rep.状態 = 'OK'; $plan += $rep
    }

    Write-Host ''
    Write-Host '=== 入れる既往歴 ===' -ForegroundColor Cyan
    $plan | Sort-Object 受診日, { [int]$_.受付番号 } | Select-Object 受診日, 受付番号, 氏名, 病名, 治療名, 操作, 連番, 状態 | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    $ok = @($plan | Where-Object { $_.状態 -eq 'OK' })
    Write-Host '=== 受診日 × 病名 の件数 ===' -ForegroundColor Cyan
    foreach ($g in ($ok | Group-Object 受診日, 病名 | Sort-Object Name)) { Write-Host ("  {0}  {1,3} 件" -f $g.Name, $g.Count) }
    $hit = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(h.BYOMEI)) AS B, LTRIM(RTRIM(h.KOMOKU_CD)) AS CD, MAX(m.MEISYO1) AS NM, LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) AS H FROM T_KIOU_HANTEI h LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(h.KOMOKU_CD)) WHERE LTRIM(RTRIM(h.BYOMEI)) IN (N'高血圧', N'糖尿病', N'高脂血症') GROUP BY LTRIM(RTRIM(h.BYOMEI)), LTRIM(RTRIM(h.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) ORDER BY 1, 2"
    Write-Host ''
    Write-Host '=== 自動判定をかけたときに判定に効くもの (健診ナビの対応表 T_KIOU_HANTEI) ===' -ForegroundColor Yellow
    foreach ($r in $hit.Rows) { Write-Host ("  {0,-6} → {1} {2} が {3}" -f $r.B, $r.CD, $r.NM, $r.H) }
    Line
    Write-Host ("入れる: {0} 件 ({1} 人) / 既にあり・要確認: {2} 件" -f $ok.Count, @($ok | Group-Object KOJIN_ID).Count, ($plan.Count - $ok.Count)) -ForegroundColor Green
    Line
    if ($ok.Count -eq 0) { Write-Host '入れるものがありません。'; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $undoFile = Join-Path $BackupDir ("kiou_med_{0}.csv" -f $stamp)
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($ok | Group-Object KOJIN_ID)) {
        $kj = $grp.Name; $pk = $grp.Group[0].PkSeq
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} ({1}) は健診ナビで編集中です ({2})" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $lock) -ForegroundColor Yellow; $skipped += $grp.Group[0].受付番号; continue }
        (Invoke-DbQuery $conn 'SELECT * FROM T_KIOU WHERE KOJIN_ID = @k' @{ k = $kj }) | Export-Csv -Path (Join-Path $BackupDir ("T_KIOU_{0}_{1}.csv" -f $kj, $stamp)) -NoTypeInformation -Encoding UTF8
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KIOU_IMPORT:" + $kj)
            foreach ($t in $grp.Group) {
                $age = $(if ($t.年齢 -match '^\d+$') { [int]$t.年齢 } else { [DBNull]::Value })
                if ($t.操作 -eq '置換') {
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KIOU SET BYOMEI = @b, BYOMEI_CD = @bc, CHIRYO = @c, CHIRYOMEI = @cm, AGE = @a WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @ob" `
                        @{ b = $t.病名; bc = $t.病名CD; c = $t.治療; cm = $t.治療名; a = $age; k = $kj; r = [int]$t.連番; ob = $t.前.B }
                } else {
                    $n = Invoke-DbExec $conn $tran "INSERT INTO T_KIOU (KOJIN_ID, RENBAN, BYOMEI, BYOMEI_CD, CHIRYO, CHIRYOMEI, AGE, ZOKUGARA, ZOKUGARAMEI, BIKOU) VALUES (@k, @r, @b, @bc, @c, @cm, @a, '0', @zm, '')" `
                        @{ k = $kj; r = [int]$t.連番; b = $t.病名; bc = $t.病名CD; c = $t.治療; cm = $t.治療名; a = $age; zm = $selfMei }
                }
                if ($n -ne 1) { throw ("書込の影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.病名) }
                $undoRows += New-Object PSObject -Property @{ KOJIN_ID = $kj; RENBAN = $t.連番; 操作 = $t.操作; 病名 = $t.病名; 受付番号 = $t.受付番号; 氏名 = $t.氏名
                    前の病名 = $(if ($t.前) { $t.前.B } else { '' }); 前の病名CD = $(if ($t.前) { $t.前.BC } else { '' }); 前の治療 = $(if ($t.前) { $t.前.C } else { '' })
                    前の治療名 = $(if ($t.前) { $t.前.CM } else { '' }); 前の年齢 = $(if ($t.前) { $t.前.A } else { '' }) }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [書込] {0} ({1}) : {2}" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, (($grp.Group | ForEach-Object { "{0}({1})" -f $_.病名, $_.治療名 }) -join ' / ')) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) { $undoRows | Select-Object KOJIN_ID, RENBAN, 操作, 病名, 前の病名, 前の病名CD, 前の治療, 前の治療名, 前の年齢, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} 件の既往歴を入れました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0}" -f ($skipped -join ', ')) -ForegroundColor Red }
    Write-Host '   このあと自動判定をかけると、高血圧・糖尿病・高脂血症 (治療中) の人の判定が J などになります。元に戻すときは kiou_med_undo.bat。' -ForegroundColor Yellow
}
finally { $conn.Close() }
