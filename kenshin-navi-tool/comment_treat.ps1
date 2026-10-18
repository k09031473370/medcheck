<#
.SYNOPSIS
  くくり判定が J (治療中) の人の医師指示事項を「治療を継続してください」に揃える (comment_treat.ps1)

.DESCRIPTION
  服薬歴から既往歴 (高血圧・糖尿病・高脂血症 / 薬剤治療中) を入れて自動判定をかけると、血圧 600150・糖代謝 600340・脂質代謝 600320 が J になる。
  ところが医師指示事項 (総合判定N(編集) 11001〜11009) には、それより前に付けた「精密検査の必要があります」などの文が残る。
  ユーザー決定 (2026/10/18): J になった病気の文は「治療を継続してください」に変える。
    ・J のくくりと同じ種類の文 (血圧 / 糖 / 脂質) を探す。
      - 治療継続の文が既にあれば、ほかの同じ種類の文を消す。
      - 無ければ、最初の 1 つを治療継続の文に差し替え、残りの同じ種類の文を消す。
      - 同じ種類の文が 1 つも無ければ、空いている枠に治療継続の文を足す。
    ・治療継続の文は、2026 年に J が付いた人のくくりの結果に入っている文 (健診ナビの自動判定が使う文) から拾う。
  ・プレビュー → y で書く。書く前に本人の T_KENSA を backup\ に保存し、変更を backup\kukuri_comment_treat_<日時>.csv に残す (kukuri_comment_undo.bat で戻せる)。
  例:
    comment_treat.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
#>
[CmdletBinding()]
param(
    [string]$Dates,
    [string]$UkeNo,
    [int]$UkeFrom = 0,
    [int]$UkeTo = 0,
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
$SOGO  = '11000'
# J になるくくりと、同じ種類の文の見分け方
$KINDS = @(
    @{ Cd = '600150'; Name = '血圧'; Pat = '血圧'; Not = ''; Like = @('血圧'); Fallback = '【血圧】　治療を継続してください。' },
    @{ Cd = '600340'; Name = '糖';   Pat = '糖代謝|糖尿病|血糖|ＨｂＡ１ｃ|HbA1c|ヘモグロビンA1c'; Not = '^\s*【尿糖】'; Like = @('糖代謝','糖尿病'); Fallback = '【糖代謝】　治療を継続してください。' },
    @{ Cd = '600320'; Name = '脂質'; Pat = '脂質|中性脂肪|コレステロール|ｺﾚｽﾃﾛｰﾙ|ＬＤＬ|LDL|ＨＤＬ|HDL'; Not = ''; Like = @('脂質','コレステロール','中性脂肪'); Fallback = '【脂質代謝】　治療を継続してください。' }
)
function Is-Kind($k, [string]$t) { if ($t -eq '' -or $t -eq '#') { return $false }; if ($k.Not -ne '' -and $t -match $k.Not) { return $false }; return ($t -match $k.Pat) }

$dateList = @(Split-List $Dates | ForEach-Object { $y = Normalize-Ymd $_; if (-not $y) { throw "受診日が読めません: $_" }; $y })
$ukeList  = @(Split-List $UkeNo | ForEach-Object { if ($_ -notmatch '^\d+$') { throw "受付番号が数字ではありません: $_" }; [int]$_ })
if ($dateList.Count -eq 0) { throw '-Dates に受診日を指定してください' }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' くくり判定が J (治療中) の人の医師指示事項を「治療を継続してください」に揃える' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  受診日: {0}  受付番号 {1}〜{2}" -f ($dateList -join ', '), $UkeFrom, $UkeTo)

$conn = Open-Db
try {
    # 治療継続の文 (2026 年に J の人のくくりの結果でいちばん多い文)
    $TEXT = @{}
    foreach ($k in $KINDS) {
        $r = Invoke-DbQuery $conn "SELECT TOP 1 LTRIM(RTRIM(k.KEKKA)) AS V, COUNT(*) AS N FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND LTRIM(RTRIM(k.KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'J' AND k.KEKKA LIKE N'%治療%' GROUP BY LTRIM(RTRIM(k.KEKKA)) ORDER BY 2 DESC" @{ cd = $k.Cd }
        if ($r.Rows.Count -eq 0) {
            # くくりの結果に文が無ければ (脂質代謝は結果に「J」しか入らないことがある)、J の人の医師指示事項から拾う
            $likes = ($k.Like | ForEach-Object { "c.KEKKA LIKE N'%" + $_ + "%'" }) -join ' OR '
            $r = Invoke-DbQuery $conn "SELECT TOP 1 LTRIM(RTRIM(c.KEKKA)) AS V, COUNT(*) AS N FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = @cd AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'J' JOIN T_KENSA c ON c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN ('11001','11002','11003','11004','11005','11006','11007','11008','11009') WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND c.KEKKA LIKE N'%治療%' AND ($likes) GROUP BY LTRIM(RTRIM(c.KEKKA)) ORDER BY 2 DESC" @{ cd = $k.Cd }
        }
        if ($r.Rows.Count -eq 0) { $TEXT[$k.Cd] = $k.Fallback; Write-Host ("  {0}: DB に見本が無いので「{1}」を使います" -f $k.Name, $k.Fallback) -ForegroundColor Yellow; continue }
        $TEXT[$k.Cd] = [string]$r.Rows[0].V
        Write-Host ("  {0}: {1}  ({2} 件で使用)" -f $k.Name, $TEXT[$k.Cd], $r.Rows[0].N) -ForegroundColor DarkGray
    }

    $dLit = "'" + ($dateList -join "','") + "'"
    $where = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit)"
    if ($ukeList.Count -gt 0) { $where += " AND s.UKE_NO_KENSA IN (" + ($ukeList -join ',') + ")" }
    if ($UkeFrom -gt 0) { $where += " AND s.UKE_NO_KENSA >= $UkeFrom" }
    if ($UkeTo -gt 0)   { $where += " AND s.UKE_NO_KENSA <= $UkeTo" }
    $cds = "'" + ((@($KINDS | ForEach-Object { $_.Cd }) + $SLOTS) -join "','") + "'"
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($cds)
WHERE $where
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@).Rows
    $P = [ordered]@{}
    foreach ($r in $rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $P.Contains($pk)) { $P[$pk] = @{ Uke = [string]$r.UKE; Ymd = [string]$r.YMD; Name = [string]$r.NAME; C = @{} } }
        $P[$pk].C[[string]$r.CD] = @{ H = ([string]$r.H).ToUpper(); V = [string]$r.V; VC = [string]$r.VC }
    }

    $plan = @()
    foreach ($pk in $P.Keys) {
        $p = $P[$pk]; $c = $p.C; $used = @{}
        foreach ($k in $KINDS) {
            if (-not $c.ContainsKey($k.Cd) -or $c[$k.Cd].H -ne 'J') { continue }
            $want = $TEXT[$k.Cd]
            $same = @($SLOTS | Where-Object { $c.ContainsKey($_) -and (Is-Kind $k $c[$_].V) })
            $has  = @($same | Where-Object { $c[$_].V -eq $want })
            $base = @{ PkSeq = $pk; 受診日 = $p.Ymd; 受付番号 = $p.Uke; 氏名 = $p.Name; 種類 = $k.Name }
            if ($same.Count -eq 0) {
                $slot = @($SLOTS | Where-Object { $c.ContainsKey($_) -and -not $used.ContainsKey($_) -and ($c[$_].V -eq '' -or $c[$_].V -eq '#') }) | Select-Object -First 1
                if (-not $slot) { $plan += New-Object PSObject -Property ($base + @{ 枠 = ''; 操作 = '枠が足りない'; 前 = ''; 前CD = ''; 後 = $want }); continue }
                $used[$slot] = 1
                $plan += New-Object PSObject -Property ($base + @{ 枠 = $slot; 操作 = '追加'; 前 = $c[$slot].V; 前CD = $c[$slot].VC; 後 = $want })
                continue
            }
            $keep = $(if ($has.Count -gt 0) { $has[0] } else { $same[0] })
            if ($has.Count -eq 0) { $plan += New-Object PSObject -Property ($base + @{ 枠 = $keep; 操作 = '差し替え'; 前 = $c[$keep].V; 前CD = $c[$keep].VC; 後 = $want }) }
            foreach ($sc in $same) {
                if ($sc -eq $keep) { continue }
                $plan += New-Object PSObject -Property ($base + @{ 枠 = $sc; 操作 = '消す'; 前 = $c[$sc].V; 前CD = $c[$sc].VC; 後 = '' })
            }
        }
    }

    Write-Host ''
    Write-Host '=== 直すコメント ===' -ForegroundColor Cyan
    $plan | Select-Object 受診日, 受付番号, 氏名, 種類, 操作, 枠, 前, 後 | Format-Table -AutoSize -Wrap | Out-String -Width 240 | Write-Host
    $ok = @($plan | Where-Object { $_.操作 -in @('追加','差し替え','消す') })
    Line
    Write-Host ("差し替え: {0} / 追加: {1} / 消す: {2} / 枠が足りない: {3}  (J の欄がある人: {4} 人)" -f @($ok | Where-Object { $_.操作 -eq '差し替え' }).Count, @($ok | Where-Object { $_.操作 -eq '追加' }).Count,
        @($ok | Where-Object { $_.操作 -eq '消す' }).Count, @($plan | Where-Object { $_.操作 -eq '枠が足りない' }).Count,
        @($P.Keys | Where-Object { $pp = $P[$_]; @($KINDS | Where-Object { $pp.C.ContainsKey($_.Cd) -and $pp.C[$_.Cd].H -eq 'J' }).Count -gt 0 }).Count) -ForegroundColor Green
    Line
    if ($ok.Count -eq 0) { Write-Host '直すものがありません。'; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。「消す」は文を空にします (kukuri_comment_undo.bat で戻せます)。' -ForegroundColor Yellow
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    $undoFile = Join-Path $BackupDir ("kukuri_comment_treat_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($ok | Group-Object PkSeq)) {
        $pk = [decimal]$grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} ({1}) は健診ナビで編集中です ({2})" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $lock) -ForegroundColor Yellow; $skipped += $grp.Group[0].受付番号; continue }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            foreach ($t in $grp.Group) {
                $q = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(ISNULL(KEKKA,''))) AS V FROM T_KENSA WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd" @{ p = $pk; cd = $t.枠 } $tran
                if ($q.Rows.Count -ne 1 -or [string]$q.Rows[0].V -ne $t.前) { Write-Host ("  [変わっていたので触らず] {0} {1}" -f $t.氏名, $t.枠) -ForegroundColor Yellow; continue }
                $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET KEKKA = @v, KEKKA_CD = @c WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' @{ v = $t.後; c = ''; p = $pk; cd = $t.枠 }
                if ($n -ne 1) { throw ("UPDATE 影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.枠) }
                $undoRows += New-Object PSObject -Property @{ PK_SEQ = [string]$pk; KOMOKU_CD = $t.枠; 前の文章 = $t.前; 前の番号 = $t.前CD; 後の文章 = $t.後; 受付番号 = $t.受付番号; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [直した] {0} ({1}) : {2} 件" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $grp.Count) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) { $undoRows | Select-Object PK_SEQ, KOMOKU_CD, 前の文章, 前の番号, 後の文章, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} か所を直しました。記録: {1}  (戻すときは kukuri_comment_undo.bat でこの記録を選ぶ)" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0}" -f ($skipped -join ', ')) -ForegroundColor Red }
}
finally { $conn.Close() }
