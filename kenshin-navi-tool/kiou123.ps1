<#
.SYNOPSIS
  石屋 9/9・9/10 の既往歴を「入力画面バックアップ」(123 列) から健診ナビの既往歴 (T_KIOU) に入れる (kiou123.ps1)

.DESCRIPTION
  健診ナビは既往歴 (T_KIOU) と「病名 → 判定」の対応表 (T_KIOU_HANTEI) を見て、自動判定のときに
  高血圧 → 血圧、糖尿病 → 糖代謝、高脂血症 → 脂質代謝 を J (治療中) などにする。石屋は既往歴が入っていない
  (198 人が「特記事項なし」だけ) ので、服薬中でも J が付かなかった (check106〜109)。
  123 列ファイルの 60〜65 列 = 病名1・年齢1・転帰1・病名2・年齢2・転帰2 は健診ナビと同じコード
  (病名 19 = 高血圧、転帰 05 = 薬剤治療中)。form\kiou_codes.csv で名前に直して入れる。
  ・本人の既往歴として入れる。最初の 1 件は「特記事項なし」の行を置き換え、2 件目は行を足す。
  ・同じ病名が既にあれば入れない。本人の既往歴は合計 -Max 件 (既定 4) まで。
  ・病名の入った行は消さない・書き換えない。
  ・プレビュー → y で書く。書く前に本人の T_KIOU を backup\ に保存し、変更を backup\kiou123_<日時>.csv に残す。-Undo で戻せる。
  ・入れたあと自動判定をかけると J などが付く (かけなければ判定は変わらず、結果票の既往歴欄に出るだけ)。
  例:
    kiou123.ps1 -File "...\910 入力＿受診票_問診票_BKUP ...xlsx"
    kiou123.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$File,
    [string]$Dates = '2026/09/09,2026/09/10',
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
Invoke-Expression (Get-Part 'function Normalize-Text'     '$XlsxLib = Join-Path')
. (Join-Path $PSScriptRoot 'xlsx_read.ps1')
Invoke-Expression (Get-Part 'function Test-XlsxEncrypted' 'function Detect-MappingPath')
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
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'kiou123_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\kiou123_*.csv) がありません。'; return }
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
if (-not $File) { $File = Read-Host '123 列ファイル (入力＿受診票_問診票_BKUP) をドラッグ＆ドロップして Enter' }
$File = ($File -replace '^"|"$', '').Trim()
if (-not (Test-Path $File)) { throw "ファイルが見つかりません: $File" }
$rows = Read-FormCsv $File 'SJIS'
$head = $rows[0]
if ($head.Count -lt 71 -or (Normalize-Text (Get-Field $head 60)) -notlike '*病名*' -or (Normalize-Text (Get-Field $head 62)) -notlike '*転帰*') {
    throw ("123 列の入力画面バックアップではないようです (60 列目「{0}」 / 62 列目「{1}」)。" -f (Get-Field $head 60), (Get-Field $head 62))
}
$dateList = @(Split-List $Dates | ForEach-Object { Normalize-Ymd $_ })

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 既往歴の取込 (123 列 → T_KIOU、本人)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($File))
Write-Host ("  受診日: {0}  受付番号 {1}〜{2}  1 人 {3} 件まで" -f ($dateList -join ', '), $UkeFrom, $UkeTo, $Max)

$conn = Open-Db
try {
    $dLit = "'" + ($dateList -join "','") + "'"
    $people = @{}
    foreach ($r in (Invoke-DbQuery $conn "SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, s.KOJIN_ID, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit) AND s.UKE_NO_KENSA BETWEEN $UkeFrom AND $UkeTo").Rows) {
        $people[[string]$r.UKE] = @{ Pk = [string]$r.PK_SEQ; Kojin = [string]$r.KOJIN_ID; Name = [string]$r.NAME }
    }
    # 本人の続柄名 (実データでいちばん多いもの)
    $selfMei = '本人'
    $z = Invoke-DbQuery $conn "SELECT TOP 1 LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) AS M, COUNT(*) AS C FROM T_KIOU WHERE LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) = '0' GROUP BY LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) ORDER BY 2 DESC"
    if ($z.Rows.Count -gt 0 -and [string]$z.Rows[0].M -ne '') { $selfMei = [string]$z.Rows[0].M }

    # 今の既往歴
    $exist = @{}
    if ($people.Count -gt 0) {
        $kLit = (@($people.Values | ForEach-Object { $_.Kojin }) | Sort-Object -Unique) -join ','
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

    # 計画
    $plan = @(); $noHist = 0; $notFound = @()
    foreach ($f in $rows[1..($rows.Count - 1)]) {
        $uke = Normalize-KenNo (Get-Field $f 2)
        if ($uke -notmatch '^\d+$') { continue }
        $items = @()
        foreach ($c in @(60, 63)) {
            $bc = Normalize-Text (Get-Field $f $c)
            if ($bc -eq '') { continue }
            $items += @{ BC = $bc; Age = (Normalize-Text (Get-Field $f ($c + 1))); Ch = (Normalize-Text (Get-Field $f ($c + 2))) }
        }
        if ($items.Count -eq 0) { $noHist++; continue }
        if (-not $people.ContainsKey($uke)) { $notFound += $uke; continue }
        $p = $people[$uke]; $kj = $p.Kojin
        if (-not $exist.ContainsKey($kj)) { $exist[$kj] = @{ MaxR = 0; Self = @{}; Count = 0; Holder = $null } }
        $e = $exist[$kj]
        foreach ($it in $items) {
            $bn = 0; [void][int]::TryParse(($it.BC -replace '^0+(?=\d)', ''), [ref]$bn)
            $cn = -1; if ($it.Ch -ne '') { [void][int]::TryParse(($it.Ch -replace '^0+(?=\d)', ''), [ref]$cn) }
            $rep = New-Object PSObject -Property @{ 受付番号 = $uke; 氏名 = $p.Name; PkSeq = $p.Pk; KOJIN_ID = $kj
                病名CD = [string]$bn; 病名 = ''; 年齢 = $it.Age; 治療 = $(if ($cn -ge 0) { [string]$cn } else { '' }); 治療名 = ''
                操作 = ''; 連番 = ''; 前 = $null; 状態 = '' }
            if (-not $BY.ContainsKey($bn)) { $rep.状態 = ("病名コード {0} が表にありません" -f $it.BC); $plan += $rep; continue }
            $rep.病名 = $BY[$bn]
            if ($cn -ge 0) { if ($CH.ContainsKey($cn)) { $rep.治療名 = $CH[$cn] } else { $rep.状態 = ("転帰コード {0} が表にありません" -f $it.Ch); $plan += $rep; continue } }
            if ($e.Self.ContainsKey($rep.病名)) { $rep.状態 = '既に入っています'; $plan += $rep; continue }
            if ($e.Count -ge $Max) { $rep.状態 = ("本人の既往歴が既に {0} 件" -f $Max); $plan += $rep; continue }
            if ($null -ne $e.Holder) { $rep.操作 = '置換'; $rep.連番 = [string]$e.Holder.R; $rep.前 = $e.Holder; $e.Holder = $null }
            else { $e.MaxR++; $rep.操作 = '追加'; $rep.連番 = [string]$e.MaxR }
            $e.Self[$rep.病名] = 1; $e.Count++
            $rep.状態 = 'OK'; $plan += $rep
        }
    }

    Write-Host ''
    Write-Host '=== 入れる既往歴 ===' -ForegroundColor Cyan
    $plan | Sort-Object { [int]$_.受付番号 } | Select-Object 受付番号, 氏名, 病名CD, 病名, 年齢, 治療名, 操作, 連番, 状態 | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    $ok = @($plan | Where-Object { $_.状態 -eq 'OK' })
    Write-Host '=== 病名ごとの件数 ===' -ForegroundColor Cyan
    foreach ($g in ($ok | Group-Object 病名 | Sort-Object Count -Descending)) { Write-Host ("  {0,3} 件  {1}" -f $g.Count, $g.Name) }
    # 自動判定で変わるもの
    $names = @($ok | ForEach-Object { $_.病名 } | Sort-Object -Unique)
    if ($names.Count -gt 0) {
        $inList = "N'" + (($names | ForEach-Object { $_ -replace "'", "''" }) -join "',N'") + "'"
        $hit = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(h.BYOMEI)) AS B, LTRIM(RTRIM(h.KOMOKU_CD)) AS CD, MAX(m.MEISYO1) AS NM, LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) AS H FROM T_KIOU_HANTEI h LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(h.KOMOKU_CD)) WHERE LTRIM(RTRIM(h.BYOMEI)) IN ($inList) GROUP BY LTRIM(RTRIM(h.BYOMEI)), LTRIM(RTRIM(h.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) ORDER BY 1, 2"
        Write-Host ''
        Write-Host '=== 自動判定をかけたときに判定に効く病名 (健診ナビの対応表 T_KIOU_HANTEI) ===' -ForegroundColor Yellow
        foreach ($r in $hit.Rows) {
            $n = @($ok | Where-Object { $_.病名 -eq [string]$r.B }).Count
            Write-Host ("  {0,-10} {1,3} 人 → {2} {3} が {4}" -f $r.B, $n, $r.CD, $r.NM, $r.H)
        }
    }
    Line
    Write-Host ("入れる: {0} 件 ({1} 人) / 既往歴なし: {2} 人 / 要確認: {3} 件 / 健診ナビに見つからない受付番号: {4}" -f $ok.Count, @($ok | Group-Object KOJIN_ID).Count, $noHist,
        @($plan | Where-Object { $_.状態 -ne 'OK' -and $_.状態 -ne '既に入っています' }).Count, $(if ($notFound.Count -gt 0) { $notFound -join ',' } else { 'なし' })) -ForegroundColor Green
    Line
    if ($ok.Count -eq 0) { Write-Host '入れるものがありません。'; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $undoFile = Join-Path $BackupDir ("kiou123_{0}.csv" -f $stamp)
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
    Write-Host '   このあと自動判定をかけると、高血圧・糖尿病・高脂血症 (治療中) の人の判定が J などになります。元に戻すときは kiou123_undo.bat。' -ForegroundColor Yellow
}
finally { $conn.Close() }
