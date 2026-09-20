<#
.SYNOPSIS
  雛形のコースを複製して、団体に新しいコースを作る (コース作成.ps1)

.DESCRIPTION
  健診ナビのコースマスタ (T_COURSE1〜4 ほか) に、既にあるコースを雛形として
  新しいコースを追加する。何を作るかは form\course_new.csv に書く。

  やること
    1. COURSE_CD と DANTAI_CD1 を持つテーブルを自動でさがす (実績系は除く)
    2. 雛形コースの行を読む
    3. 団体CD と コースCD を差し替えて、必要なら 名称・料金・項目 を変える
    4. プレビューを出す → Y → 書き込む

  安全のため
    ・先に全部プレビューして Y を押すまで何も書かない
    ・書き込む前に、対象テーブルの「その団体の行」を backup に控える
    ・取り消し用の DELETE 文も backup に書き出す
    ・1つのトランザクションでまとめて書く。途中で失敗したら全部戻す
    ・applock を取るので、他の人の操作と重ならない
    ・作ろうとしたコースCDが既にあれば、そのコースは飛ばす (二重作成しない)
    ・書いたあと読み直して、狙った行が入ったことを確かめる

  form\course_new.csv の書き方
    CourseCd       新しいコースCD (例 MMD1)
    Meisyo         コース名
    MeisyoRyaku    略称 (空なら Meisyo と同じ)
    BaseDantai     雛形の団体CD (例 9999999997)
    BaseCourseCd   雛形のコースCD (例 SHTD1)
    DelKomoku      雛形から外す項目CD。; か空白で区切る
    AddKomoku      雛形に足す項目CD。; か空白で区切る
    KenpoRyoukin   健保料金   (空なら雛形のまま)
    DantaiRyoukin  団体料金   (空なら雛形のまま)
    KojinRyoukin   個人料金   (空なら雛形のまま)
    JikoFutanBun   自己負担の文言 (空なら雛形のまま。"(空)" と書くと空にする)
    Note           覚書。読むだけで使わない

.EXAMPLE
  コース作成.bat に 団体CD を聞かれるので入れる (例 0000000475)

  powershell -ExecutionPolicy Bypass -File コース作成.ps1 -Dantai 0000000475
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Dantai,              # 作る先の団体CD
    [string]$Csv,                 # 省略時は form\course_new.csv
    [switch]$Commit,              # 付けると確認なしで書き込む
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

# form_import.ps1 から接続・ロックまわりを借りる
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text'     '$XlsxLib = Join-Path')
Invoke-Expression (Get-Part 'function Get-LocalConnFile'  'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Test-Locked'        'function Commit-Plan')

$BackupDir = Join-Path $PSScriptRoot 'backup'
if (-not $Csv) { $Csv = Join-Path $PSScriptRoot 'form\course_new.csv' }

# 受診実績・請求など、コースの定義ではないテーブル。複製しない。
$SKIP_TABLES = @('T_KENSIN', 'T_KENSIN_LOG', 'T_RYOUKIN', 'T_SEIKYU3')
# 必ず先に作るテーブル (コース本体)
$FIRST_TABLE = 'T_COURSE1'

function Split-Codes([string]$s) {
    if ($null -eq $s) { return @() }
    return @(($s -split '[;,\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
}

# ============================================================================
# 入力
# ============================================================================
if (-not $Dantai) { $Dantai = Read-Host '作る先の団体CDを入れて Enter (例 0000000475)' }
$Dantai = $Dantai.Trim()
if ($Dantai -eq '') { throw '団体CDが空です。' }
if (-not (Test-Path $Csv)) { throw "コースの定義CSVがありません: $Csv" }

$defs = @(Import-Csv -Path $Csv -Encoding UTF8)
if ($defs.Count -eq 0) { throw "定義CSVに行がありません: $Csv" }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 雛形のコースを複製して、新しいコースを作る' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  作る先の団体CD : {0}" -f $Dantai)
Write-Host ("  定義           : {0}  ({1} コース)" -f $Csv, $defs.Count)
Write-Host ''

$conn = Open-Db
try {
    # ---- 団体の確認 ----
    $d = @(Invoke-DbQuery $conn 'SELECT DANTAI_CD1, MEISYO1 FROM T_DANTAI1 WHERE DANTAI_CD1 = @d' @{ d = $Dantai })
    if ($d.Count -eq 0) { throw "団体CD $Dantai が団体マスタ (T_DANTAI1) にありません。先に健診ナビで事業所を登録してください。" }
    Write-Host ("  団体 : {0}  {1}" -f $Dantai, $d[0].MEISYO1) -ForegroundColor Green
    Write-Host ''

    # ---- 複製するテーブルをさがす ----
    $tblSql = @"
SELECT o.name AS TBL
FROM sys.objects o
WHERE o.type = 'U'
  AND o.object_id IN (SELECT object_id FROM sys.columns WHERE name = 'COURSE_CD')
  AND o.object_id IN (SELECT object_id FROM sys.columns WHERE name = 'DANTAI_CD1')
ORDER BY o.name
"@
    $tables = @((Invoke-DbQuery $conn $tblSql @{}) | ForEach-Object { [string]$_.TBL } |
               Where-Object { $SKIP_TABLES -notcontains $_ })
    if ($tables -notcontains $FIRST_TABLE) { throw "$FIRST_TABLE が見つかりません。" }
    # T_COURSE1 を先頭にする
    $tables = @($FIRST_TABLE) + @($tables | Where-Object { $_ -ne $FIRST_TABLE })
    Write-Host ("  複製するテーブル: {0}" -f ($tables -join ', ')) -ForegroundColor DarkGray
    Write-Host ("  複製しないテーブル (受診実績・請求): {0}" -f ($SKIP_TABLES -join ', ')) -ForegroundColor DarkGray
    Write-Host ''

    # ---- 各テーブルの書ける列 (IDENTITY・計算列は除く) ----
    $colsOf = @{}
    foreach ($t in $tables) {
        $cs = @(Invoke-DbQuery $conn @"
SELECT c.name AS COL, ty.name AS TY, c.is_nullable AS NUL
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID(@t) AND c.is_identity = 0 AND c.is_computed = 0
ORDER BY c.column_id
"@ @{ t = $t })
        $colsOf[$t] = $cs
    }

    # ========================================================================
    # 何を書くか組み立てる (まだ書かない)
    # ========================================================================
    $plan = @()      # @{ Table; Cols; Values; Course; Label }
    $skip = @()
    $warn = @()

    foreach ($df in $defs) {
        $cc     = ([string]$df.CourseCd).Trim()
        $bd     = ([string]$df.BaseDantai).Trim()
        $bc     = ([string]$df.BaseCourseCd).Trim()
        if ($cc -eq '' -or $bd -eq '' -or $bc -eq '') { $warn += "定義CSVに空の行があります。飛ばします。"; continue }

        # 既にあれば作らない
        $ex = @(Invoke-DbQuery $conn 'SELECT COURSE_CD FROM T_COURSE1 WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c' @{ d = $Dantai; c = $cc })
        if ($ex.Count -gt 0) { $skip += ("{0} は既にこの団体にあります" -f $cc); continue }
        # 雛形があるか
        $bs = @(Invoke-DbQuery $conn 'SELECT COURSE_CD FROM T_COURSE1 WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c' @{ d = $bd; c = $bc })
        if ($bs.Count -eq 0) { $warn += ("雛形 {0}/{1} が見つかりません。{2} は作りません" -f $bd, $bc, $cc); continue }

        $del = Split-Codes $df.DelKomoku
        $add = Split-Codes $df.AddKomoku

        foreach ($t in $tables) {
            $rows = @(Invoke-DbQuery $conn ("SELECT * FROM [$t] WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c") @{ d = $bd; c = $bc })
            if ($rows.Count -eq 0) { continue }
            foreach ($r in $rows) {
                # 外す項目
                if ($t -eq 'T_COURSE2') {
                    $k = ([string]$r.KOMOKU_CD).Trim()
                    if ($del -contains $k) { continue }
                }
                $vals = @{}
                foreach ($c in $colsOf[$t]) {
                    $cn = [string]$c.COL
                    $v  = $null
                    if ($r.PSObject.Properties.Name -contains $cn) { $v = $r.$cn }
                    if ($v -is [System.DBNull]) { $v = $null }
                    $vals[$cn] = $v
                }
                $vals['DANTAI_CD1'] = $Dantai
                $vals['COURSE_CD']  = $cc
                if ($t -eq 'T_COURSE1') {
                    if (([string]$df.Meisyo).Trim() -ne '')      { $vals['MEISYO'] = ([string]$df.Meisyo).Trim() }
                    $ry = ([string]$df.MeisyoRyaku).Trim()
                    if ($ry -eq '') { $ry = ([string]$df.Meisyo).Trim() }
                    if ($ry -ne '' -and $vals.ContainsKey('MEISYO_RYAKUSYO')) { $vals['MEISYO_RYAKUSYO'] = $ry }
                    $jf = [string]$df.JikoFutanBun
                    if ($jf -eq '(空)') { $vals['JIKO_FUTAN_BUN'] = '' }
                    elseif ($jf.Trim() -ne '') { $vals['JIKO_FUTAN_BUN'] = $jf.Trim() }
                }
                if ($t -eq 'T_COURSE3') {
                    foreach ($pair in @(@('KenpoRyoukin','KENPO_RYOUKIN'), @('DantaiRyoukin','DANTAI_RYOUKIN'), @('KojinRyoukin','KOJIN_RYOUKIN'))) {
                        $sv = ([string]$df.($pair[0])).Trim()
                        if ($sv -ne '' -and $vals.ContainsKey($pair[1])) { $vals[$pair[1]] = [decimal]$sv }
                    }
                }
                $plan += New-Object PSObject -Property @{ Table = $t; Vals = $vals; Course = $cc; Kind = '複製' }
            }
            # 足す項目
            if ($t -eq 'T_COURSE2' -and $add.Count -gt 0) {
                foreach ($k in $add) {
                    $km = @(Invoke-DbQuery $conn 'SELECT KOMOKU_CD, MEISYO1 FROM T_KOMOKU WHERE LTRIM(RTRIM(KOMOKU_CD)) = @k' @{ k = $k })
                    if ($km.Count -eq 0) { $warn += ("項目CD {0} が項目マスタにありません。{1} には足しません" -f $k, $cc); continue }
                    $dup = @($plan | Where-Object { $_.Table -eq 'T_COURSE2' -and $_.Course -eq $cc -and ([string]$_.Vals['KOMOKU_CD']).Trim() -eq $k })
                    if ($dup.Count -gt 0) { $warn += ("項目CD {0} は雛形に既にあります。{1} には足しません" -f $k, $cc); continue }
                    $vals = @{}
                    foreach ($c in $colsOf['T_COURSE2']) { $vals[[string]$c.COL] = $null }
                    $vals['DANTAI_CD1'] = $Dantai
                    $vals['COURSE_CD']  = $cc
                    $vals['KOMOKU_CD']  = $k
                    $plan += New-Object PSObject -Property @{ Table = 'T_COURSE2'; Vals = $vals; Course = $cc; Kind = ('追加 ' + $km[0].MEISYO1) }
                }
            }
        }
        # 外そうとした項目が雛形になかったら教える
        foreach ($k in $del) {
            $hit = @($plan | Where-Object { $_.Table -eq 'T_COURSE2' -and $_.Course -eq $cc -and ([string]$_.Vals['KOMOKU_CD']).Trim() -eq $k })
            $wasThere = @(Invoke-DbQuery $conn 'SELECT KOMOKU_CD FROM T_COURSE2 WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c AND LTRIM(RTRIM(KOMOKU_CD)) = @k' @{ d = $bd; c = $bc; k = $k })
            if ($wasThere.Count -eq 0) { $warn += ("項目CD {0} は雛形 {1} に元々ありません ({2} の DelKomoku)" -f $k, $bc, $cc) }
        }
    }

    # ========================================================================
    # プレビュー
    # ========================================================================
    if ($skip.Count -gt 0) { Write-Host '--- 作らないもの ---' -ForegroundColor Yellow; $skip | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow }; Write-Host '' }
    if ($warn.Count -gt 0) { Write-Host '--- 注意 ---' -ForegroundColor Yellow; $warn | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow }; Write-Host '' }
    if ($plan.Count -eq 0) { Write-Host '作るものがありません。' -ForegroundColor Yellow; return }

    Write-Host '--- 作るコース ---' -ForegroundColor Cyan
    foreach ($g in ($plan | Group-Object Course)) {
        $df = $defs | Where-Object { ([string]$_.CourseCd).Trim() -eq $g.Name } | Select-Object -First 1
        $n2 = @($g.Group | Where-Object { $_.Table -eq 'T_COURSE2' }).Count
        Write-Host ("  [{0}] {1}   雛形: {2}/{3}   検査項目 {4} 個" -f $g.Name, $df.Meisyo, $df.BaseDantai, $df.BaseCourseCd, $n2) -ForegroundColor Green
        foreach ($tg in ($g.Group | Group-Object Table | Sort-Object Name)) {
            Write-Host ("        {0,-16} {1,4} 行" -f $tg.Name, $tg.Count) -ForegroundColor DarkGray
        }
        $c3 = @($g.Group | Where-Object { $_.Table -eq 'T_COURSE3' })
        if ($c3.Count -gt 0) {
            $v = $c3[0].Vals
            Write-Host ("        料金  健保 {0} / 団体 {1} / 個人 {2}" -f $v['KENPO_RYOUKIN'], $v['DANTAI_RYOUKIN'], $v['KOJIN_RYOUKIN']) -ForegroundColor DarkGray
        }
        $adds = @($g.Group | Where-Object { $_.Kind -like '追加*' })
        if ($adds.Count -gt 0) { $adds | ForEach-Object { Write-Host ('        + ' + $_.Kind) -ForegroundColor DarkGray } }
        $dels = Split-Codes ($defs | Where-Object { ([string]$_.CourseCd).Trim() -eq $g.Name } | Select-Object -First 1).DelKomoku
        if ($dels.Count -gt 0) { Write-Host ('        - 外した項目CD: ' + ($dels -join ' ')) -ForegroundColor DarkGray }
    }
    Write-Host ''
    Write-Host ("合計 {0} 行を {1} テーブルに入れます。" -f $plan.Count, @($plan | Group-Object Table).Count) -ForegroundColor Cyan
    Write-Host ''

    if (-not $Commit) {
        $ans = Read-Host '上の内容で健診ナビにコースを作ります。よければ Y'
        if ($ans -notmatch '^[Yy]') { Write-Host '中止しました。何も書いていません。' -ForegroundColor Yellow; return }
    }

    # ========================================================================
    # 書き込む
    # ========================================================================
    if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bkDir = Join-Path $BackupDir ("コース作成_{0}_{1}" -f $Dantai, $stamp)
    [void](New-Item -ItemType Directory -Path $bkDir -Force)

    # 書く前のその団体の行を控える
    foreach ($t in $tables) {
        $cur = @(Invoke-DbQuery $conn ("SELECT * FROM [$t] WHERE DANTAI_CD1 = @d") @{ d = $Dantai })
        $cur | Export-Csv -Path (Join-Path $bkDir ("{0}_{1}_書込前.csv" -f $t, $Dantai)) -NoTypeInformation -Encoding UTF8
    }
    # 取り消し用の DELETE
    $undoSql = Join-Path $bkDir '取り消し.sql'
    $ccs = @($plan | ForEach-Object { $_.Course } | Select-Object -Unique | ForEach-Object { "'" + $_ + "'" })
    $lines = @("-- $stamp に作ったコースを取り消す。健診ナビを閉じてから実行すること。",
               "-- 受診予約が既に入っている場合は消さないこと (T_KENSIN を先に確認)。", "BEGIN TRAN")
    foreach ($t in ($tables | Sort-Object -Descending)) {
        $lines += ("DELETE FROM [{0}] WHERE DANTAI_CD1 = '{1}' AND LTRIM(RTRIM(COURSE_CD)) IN ({2});" -f $t, $Dantai, ($ccs -join ','))
    }
    $lines += @("-- 確認してから COMMIT", "-- COMMIT", "-- ROLLBACK")
    $lines | Out-File -FilePath $undoSql -Encoding Default

    $done = @{}
    $tran = $conn.BeginTransaction()
    try {
        Acquire-AppLock $conn $tran ("COURSE_CREATE:" + $Dantai)
        foreach ($p in $plan) {
            $t = $p.Table
            $cols = @($colsOf[$t] | ForEach-Object { [string]$_.COL })
            $sql = ('INSERT INTO [{0}] ({1}) VALUES ({2})' -f $t,
                    (($cols | ForEach-Object { '[' + $_ + ']' }) -join ', '),
                    (($cols | ForEach-Object { '@' + $_ }) -join ', '))
            $pr = @{}
            foreach ($c in $cols) { $pr[$c] = $p.Vals[$c] }
            $n = Invoke-DbExec $conn $tran $sql $pr
            if ($n -ne 1) { throw ("INSERT影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t, $p.Course) }
            if (-not $done.ContainsKey($t)) { $done[$t] = 0 }
            $done[$t]++
        }
        $tran.Commit()
    }
    catch { $tran.Rollback(); throw }

    Write-Host ''
    foreach ($t in ($done.Keys | Sort-Object)) { Write-Host ("  [書込] {0,-16} {1,4} 行" -f $t, $done[$t]) -ForegroundColor Green }

    # ========================================================================
    # 読み直して確かめる
    # ========================================================================
    Write-Host ''
    Write-Host '--- 読み直して確認 ---' -ForegroundColor Cyan
    $bad = @()
    foreach ($g in ($plan | Group-Object Course)) {
        $cc = $g.Name
        $c1 = @(Invoke-DbQuery $conn 'SELECT MEISYO, KIJUN_CD, TYOHYO, F_TOSINKYO FROM T_COURSE1 WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c' @{ d = $Dantai; c = $cc })
        if ($c1.Count -ne 1) { $bad += ("{0} : T_COURSE1 が {1} 行" -f $cc, $c1.Count); continue }
        $want2 = @($g.Group | Where-Object { $_.Table -eq 'T_COURSE2' }).Count
        $got2  = @(Invoke-DbQuery $conn 'SELECT COUNT(*) AS N FROM T_COURSE2 WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c' @{ d = $Dantai; c = $cc })[0].N
        if ([int]$got2 -ne $want2) { $bad += ("{0} : 検査項目が {1} 個 (予定 {2} 個)" -f $cc, $got2, $want2) }
        $c3 = @(Invoke-DbQuery $conn 'SELECT KENPO_RYOUKIN, DANTAI_RYOUKIN, KOJIN_RYOUKIN FROM T_COURSE3 WHERE DANTAI_CD1 = @d AND LTRIM(RTRIM(COURSE_CD)) = @c' @{ d = $Dantai; c = $cc })
        $ry = $(if ($c3.Count -gt 0) { "健保 {0} / 団体 {1} / 個人 {2}" -f $c3[0].KENPO_RYOUKIN, $c3[0].DANTAI_RYOUKIN, $c3[0].KOJIN_RYOUKIN } else { '料金なし' })
        Write-Host ("  {0}  {1}  基準値{2} 帳票{3} 東振協{4}  項目 {5} 個  {6}" -f `
            $cc, $c1[0].MEISYO, $c1[0].KIJUN_CD, $c1[0].TYOHYO, $c1[0].F_TOSINKYO, $got2, $ry) -ForegroundColor Green
    }
    if ($bad.Count -gt 0) {
        Write-Host ''
        Write-Host '--- 確認NG ---' -ForegroundColor Yellow
        $bad | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow }
    }

    Write-Host ''
    Write-Host ("[控え] 書込前の状態  : {0}" -f $bkDir) -ForegroundColor DarkGray
    Write-Host ("[控え] 取り消しSQL   : {0}" -f $undoSql) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '※ 健診ナビの画面を開き直して、コースが選べることを確認してください。' -ForegroundColor Yellow
    Write-Host '※ 料金は指示書の単価を団体料金に入れています。健保負担と分ける場合は健診ナビの画面で直してください。' -ForegroundColor Yellow
}
finally {
    if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
}
