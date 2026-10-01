<#
.SYNOPSIS
  会場の受付アプリのJSONを読んで、健診ナビの予約に差分を反映する (受付結果取込.ps1)

.DESCRIPTION
  福生商工会の受付アプリ (バージョン9以降) が出す
      GET /api/navi.json  →  fussa_reception_YYYYMMDD_HHMMSS.json
  を読んで、健診ナビの予約レコードを「違うところだけ」更新する。

  反映するもの (第1段階)
    ・受付番号      actual.reception_number → T_KENSIN.UKE_NO_KENSA
    ・カナ氏名      identity.kana           → T_KOJIN1.KANA_SIMEI
    ・生年月日      identity.birth_date     → T_KOJIN1 の生年月日の列 (列名は自動でさがす)
    ・コース変更    actual.course (A/B/C)   → T_KENSIN.COURSE_CD (FA/FB/FC)

  反映しないもの (一覧に出すだけ)
    ・受診日が予定と違う人 (date_changed)   … 日付を動かすと帳票・請求に響くので人が判断する
    ・キャンセル (is_cancelled)             … F_TORIKESI は触らない
    ・オプション・便本数・採血/胃/尿の変更  … 健診ナビ側の受け皿が未確定
    ・コースが未知の人 (warnings)           … 保留

  突き合わせ
    identity.name (空白を詰めたもの) で、健診ナビの受診者と照合する。
    対象は JSONの予定日に受診する「団体名に福生が入る」人。
    1件に決まらない人は書かずに一覧へ。

  安全のため
    ・先に全部プレビューして Y を押すまで何も書かない
    ・前回取り込んだファイルより古ければ止める (exported_at と revision で判断)
    ・書く前に対象行を backup に控える。取り消し用のUPDATE文も残す
    ・1つのトランザクションでまとめて書く。失敗したら全部戻す
    ・applock を取るので他の人の操作と重ならない
    ・書いたあと読み直して確かめる
    ・受診者ID ↔ 健診ナビのPK_SEQ の対応表を残す (2回目以降はこれで確実に当たる)

.EXAMPLE
  受付結果取込.bat に navi.json をドラッグ＆ドロップ

  powershell -ExecutionPolicy Bypass -File 受付結果取込.ps1 -Json C:\...\fussa_reception_20261002_1830.json
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Json,                      # 受付アプリが出した navi.json
    [string]$DantaiLike = '福生',       # 対象にする団体名のキーワード
    [switch]$OverwriteKana,             # 付けると、既に入っているカナも上書きする
    [switch]$Force,                     # 付けると、古いファイルでも続行する
    [switch]$Commit,                    # 付けると確認なしで書き込む
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
Invoke-Expression (Get-Part 'function Normalize-Text'    '$XlsxLib = Join-Path')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Test-Locked'       'function Commit-Plan')

$BackupDir = Join-Path $PSScriptRoot 'backup'
$StateFile = Join-Path $BackupDir '受付結果取込_前回.txt'   # 前回取り込んだ exported_at / revision

# 受付アプリのコース記号 → 健診ナビのコースCD
$COURSE_MAP = @{ 'A' = 'FA'; 'B' = 'FB'; 'C' = 'FC' }

function NoSpace([string]$s) {
    if ($null -eq $s) { return '' }
    return (($s -replace '[\s　]', '').Trim())
}
function ToWideKana([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $w = [Microsoft.VisualBasic.Strings]::StrConv($s, [Microsoft.VisualBasic.VbStrConv]::Wide, 1041)
    return ($w -replace '\s+', ' ').Trim()
}

# ============================================================================
# JSONを読む
# ============================================================================
if (-not $Json) { $Json = Read-Host '受付アプリのJSONをドラッグ＆ドロップして Enter' }
$Json = ($Json -replace '^"|"$', '').Trim()
if (-not (Test-Path $Json)) { throw "見つかりません: $Json" }

$doc = Get-Content -LiteralPath $Json -Raw -Encoding UTF8 | ConvertFrom-Json
foreach ($k in @('schema_version', 'series_id', 'exported_at', 'people')) {
    if (-not ($doc.PSObject.Properties.Name -contains $k)) { throw "このJSONには $k がありません。受付アプリの『健診ナビ連携JSONを出力』で出したファイルを渡してください。" }
}
if ($doc.schema_version -ne '1.0') { Write-Host ("※ 想定と違うスキーマ版です: {0}" -f $doc.schema_version) -ForegroundColor Yellow }

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ' 受付アプリのJSON → 健診ナビ (差分だけ反映)' -ForegroundColor Cyan
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  ファイル  : {0}" -f (Split-Path $Json -Leaf))
Write-Host ("  シリーズ  : {0}   スキーマ {1}   受付アプリ版 {2}   revision {3}" -f $doc.series_id, $doc.schema_version, $doc.reception_version, $doc.revision)
Write-Host ("  出力日時  : {0}" -f $doc.exported_at)
if ($doc.counts) {
    Write-Host ("  件数      : 全{0}人  受付済み{1}  未受付{2}  キャンセル{3}" -f $doc.counts.total, $doc.counts.received, $doc.counts.pending, $doc.counts.cancelled)
}
Write-Host ''

# ---- 古いファイルでないか ----
if (Test-Path $StateFile) {
    $prev = Get-Content $StateFile -Raw | ConvertFrom-Json
    if ($prev.series_id -eq $doc.series_id) {
        $pOld = [datetime]::Parse($prev.exported_at); $pNew = [datetime]::Parse($doc.exported_at)
        if ($pNew -lt $pOld -or ($pNew -eq $pOld -and [int]$doc.revision -lt [int]$prev.revision)) {
            Write-Host ("★ 前回取り込んだファイル ({0} / revision {1}) より古いファイルです。" -f $prev.exported_at, $prev.revision) -ForegroundColor Yellow
            if (-not $Force) { throw '古いデータで新しいデータを戻さないように止めました。本当に取り込むなら -Force を付けてください。' }
            Write-Host '  -Force が付いているので続行します。' -ForegroundColor Yellow
        }
        elseif ($pNew -eq $pOld -and [int]$doc.revision -eq [int]$prev.revision) {
            Write-Host '※ 前回と同じファイルのようです。差分が0件になるはずです。' -ForegroundColor DarkGray
        }
    }
}

# ---- 未知コースの警告 ----
$holdIds = @{}
if ($doc.warnings) {
    foreach ($w in $doc.warnings) { $holdIds[[string]$w.id] = $w }
    if ($holdIds.Count -gt 0) {
        Write-Host ('--- 受付アプリからの警告 ({0}件) ---' -f $holdIds.Count) -ForegroundColor Yellow
        foreach ($w in $doc.warnings) { Write-Host ("  {0}  {1}  {2} = [{3}]  → 保留します" -f $w.id, $w.code, $w.field, $w.value) -ForegroundColor Yellow }
        Write-Host ''
    }
}

$conn = Open-Db
try {
    # ---- 生年月日の列名をさがす ----
    $birthCol = $null
    $cols = @(Invoke-DbQuery $conn @"
SELECT c.name AS COL, ty.name AS TY
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KOJIN1')
"@ @{})
    foreach ($c in $cols) {
        $n = ([string]$c.COL).ToUpper()
        if ($n -match 'BIRTH' -or $n -match 'SEINEN' -or $n -match 'TANJO' -or $n -eq 'D_SEI') { $birthCol = [string]$c.COL; break }
    }
    if ($birthCol) { Write-Host ("  生年月日の列: T_KOJIN1.{0}" -f $birthCol) -ForegroundColor DarkGray }
    else           { Write-Host '  ※ T_KOJIN1 に生年月日らしい列が見つからないので、生年月日は更新しません。' -ForegroundColor Yellow }

    # ---- 受付番号の桁数 ----
    $ukeMax = 9999
    $u = @(Invoke-DbQuery $conn @"
SELECT c.precision AS P, c.scale AS S FROM sys.columns c
WHERE c.object_id = OBJECT_ID('T_KENSIN') AND c.name = 'UKE_NO_KENSA'
"@ @{})
    if ($u.Count -eq 1) {
        $p = [int]$u[0].P; $s = [int]$u[0].S
        $ukeMax = [math]::Pow(10, [math]::Max(1, $p - $s)) - 1
        Write-Host ("  受付番号の列: numeric({0},{1})  入れられる最大 {2}" -f $p, $s, $ukeMax) -ForegroundColor DarkGray
    }
    Write-Host ''

    # ---- 健診ナビ側の対象者を読む ----
    $dates = @($doc.people | ForEach-Object { [string]$_.planned.date } | Where-Object { $_ } | Select-Object -Unique)
    if ($dates.Count -eq 0) { throw 'JSONに予定日がありません。' }
    $inList = ($dates | ForEach-Object { "'" + ($_ -replace '-', '/') + "'" }) -join ','
    Write-Host ("  健診ナビ側の対象: 受診日 {0} / 団体名に「{1}」" -f ($dates -join ' '), $DantaiLike) -ForegroundColor DarkGray

    $naviSql = @"
SELECT s.PK_SEQ, s.KOJIN_ID, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS COURSE_CD,
       s.UKE_NO_KENSA AS UKE,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS KANJI,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS KANA,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS DANTAI,
       LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS DANTAI_CD
       $(if ($birthCol) { ", CONVERT(varchar(10), g.[$birthCol], 111) AS BIRTH" } else { ", '' AS BIRTH" })
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ($inList) AND d.MEISYO1 LIKE N'%$DantaiLike%'
"@
    $navi = @(Invoke-DbQuery $conn $naviSql @{})
    Write-Host ("  健診ナビ側 {0} 人 / JSON {1} 人" -f $navi.Count, $doc.people.Count) -ForegroundColor DarkGray
    Write-Host ''

    $byName = @{}
    foreach ($r in $navi) {
        $k = NoSpace $r.KANJI
        if ($k -eq '') { continue }
        if (-not $byName.ContainsKey($k)) { $byName[$k] = @() }
        $byName[$k] += $r
    }

    # ========================================================================
    # 差分を組み立てる
    # ========================================================================
    $plan = @(); $skip = @(); $note = @(); $map = @()
    foreach ($p in $doc.people) {
        $pid  = [string]$p.id
        $name = NoSpace $p.identity.name
        if ($name -eq '') { $name = NoSpace $p.planned.name }
        $label = "$pid $($p.identity.name)"

        if ($holdIds.ContainsKey($pid)) { $skip += "$label : 受付アプリの警告があるので保留"; continue }
        if (-not $byName.ContainsKey($name)) { $skip += "$label : 健診ナビに見つかりません ($($p.planned.company) / $($p.planned.date))"; continue }
        $hits = @($byName[$name])
        if ($hits.Count -gt 1) { $skip += "$label : 健診ナビに $($hits.Count) 件あって決められません"; continue }
        $n = $hits[0]
        $map += New-Object PSObject -Property @{ series_id = $doc.series_id; 受診者ID = $pid; PK_SEQ = $n.PK_SEQ; KOJIN_ID = $n.KOJIN_ID; 氏名 = $p.identity.name; 事業所 = $n.DANTAI }

        $sets = @()
        # --- 受付番号 ---
        $rn = $null
        if ($p.actual -and $null -ne $p.actual.reception_number) { $rn = [int]$p.actual.reception_number }
        if ($null -ne $rn) {
            if ($rn -gt $ukeMax) { $note += "$label : 受付番号 $rn は健診ナビの桁数($ukeMax)を超えるので入れません" }
            elseif ($null -eq $n.UKE -or $n.UKE -is [System.DBNull]) { $sets += @{ T='KENSIN'; Col='UKE_NO_KENSA'; New=$rn; Old=''; What='受付番号' } }
            elseif ([int]$n.UKE -ne $rn) { $sets += @{ T='KENSIN'; Col='UKE_NO_KENSA'; New=$rn; Old=[int]$n.UKE; What='受付番号' } }
        }
        # --- カナ ---
        $kana = ToWideKana ([string]$p.identity.kana)
        if ($kana -ne '') {
            $cur = [string]$n.KANA
            if ($cur -eq '') { $sets += @{ T='KOJIN'; Col='KANA_SIMEI'; New=$kana; Old=''; What='カナ' } }
            elseif ((NoSpace $cur) -ne (NoSpace $kana) -and $OverwriteKana) { $sets += @{ T='KOJIN'; Col='KANA_SIMEI'; New=$kana; Old=$cur; What='カナ(上書き)' } }
            elseif ((NoSpace $cur) -ne (NoSpace $kana)) { $note += "$label : カナが違います 健診ナビ[$cur] 受付[$kana] → -OverwriteKana で上書きできます" }
        }
        # --- 生年月日 ---
        $bd = [string]$p.identity.birth_date
        if ($birthCol -and $bd -ne '') {
            $bdN = $bd -replace '-', '/'
            if ([string]$n.BIRTH -eq '') { $sets += @{ T='KOJIN'; Col=$birthCol; New=$bdN; Old=''; What='生年月日'; IsDate=$true } }
            elseif ([string]$n.BIRTH -ne $bdN) { $note += "$label : 生年月日が違います 健診ナビ[$($n.BIRTH)] 受付[$bdN] → 手で確認してください" }
        }
        # --- コース変更 ---
        if ($p.actual -and $p.derived.course_changed) {
            $ac = [string]$p.actual.course
            if ($COURSE_MAP.ContainsKey($ac)) {
                $want = $COURSE_MAP[$ac]
                if ([string]$n.COURSE_CD -ne $want) { $sets += @{ T='KENSIN'; Col='COURSE_CD'; New=$want; Old=[string]$n.COURSE_CD; What="コース($ac)" } }
            } else {
                $note += "$label : 当日コースが [$ac] なので健診ナビのコースに変換できません → 手で確認してください"
            }
        }
        # --- 書かないが知らせること ---
        if ($p.derived.date_changed) { $note += "$label : 予定 $($p.planned.date) → 実際 $($p.actual.checked_in_date) に受診。受診日は自動では動かしません" }
        if ($p.derived.is_cancelled) { $note += "$label : 受付アプリ側でキャンセル。健診ナビの予約は取り消していません" }
        if ($p.actual -and (@($p.derived.options_added) + @($p.derived.options_removed)).Count -gt 0) {
            $note += ("$label : オプション変更 追加[{0}] 中止[{1}] → 第2段階 (今回は反映しません)" -f (($p.derived.options_added) -join '.'), (($p.derived.options_removed) -join '.'))
        }
        if ($p.actual -and $null -ne $p.actual.stool_count -and [int]$p.actual.stool_count -ne 2) {
            $note += "$label : 便 $($p.actual.stool_count) 本 → 第2段階 (今回は反映しません)"
        }

        if ($sets.Count -gt 0) {
            $plan += New-Object PSObject -Property @{
                ID = $pid; 氏名 = $p.identity.name; 事業所 = $n.DANTAI; 受診日 = $n.YMD
                PK_SEQ = $n.PK_SEQ; KOJIN_ID = $n.KOJIN_ID; Sets = $sets
                内容 = (($sets | ForEach-Object { "{0} [{1}]→[{2}]" -f $_.What, $_.Old, $_.New }) -join ' / ')
            }
        }
    }

    # ========================================================================
    # プレビュー
    # ========================================================================
    Write-Host ('--- 更新する人 ({0} 人 / {1} か所) ---' -f $plan.Count, (@($plan | ForEach-Object { $_.Sets.Count }) | Measure-Object -Sum).Sum) -ForegroundColor Cyan
    if ($plan.Count -gt 0) {
        $plan | Select-Object ID, 氏名, 事業所, 受診日, 内容 | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    }
    if ($skip.Count -gt 0) {
        Write-Host ('--- 更新できなかった人 ({0} 人) ---' -f $skip.Count) -ForegroundColor Yellow
        $skip | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow }; Write-Host ''
    }
    if ($note.Count -gt 0) {
        Write-Host ('--- お知らせ ({0} 件。書き込みはしません) ---' -f $note.Count) -ForegroundColor DarkYellow
        $note | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor DarkYellow }; Write-Host ''
    }
    if ($plan.Count -eq 0) { Write-Host '更新するものがありません。' -ForegroundColor Green; return }

    if (-not $Commit) {
        $ans = Read-Host ("上の {0} 人を健診ナビに反映します。よければ Y" -f $plan.Count)
        if ($ans -notmatch '^[Yy]') { Write-Host '中止しました。何も書いていません。' -ForegroundColor Yellow; return }
    }

    # ========================================================================
    # 書き込む
    # ========================================================================
    if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bkDir = Join-Path $BackupDir ("受付結果取込_{0}" -f $stamp)
    [void](New-Item -ItemType Directory -Path $bkDir -Force)
    $pks = ($plan | ForEach-Object { $_.PK_SEQ }) -join ','
    (Invoke-DbQuery $conn "SELECT * FROM T_KENSIN WHERE PK_SEQ IN ($pks)" @{}) | Export-Csv (Join-Path $bkDir 'T_KENSIN_書込前.csv') -NoTypeInformation -Encoding UTF8
    $kjs = ($plan | ForEach-Object { "'" + $_.KOJIN_ID + "'" }) -join ','
    (Invoke-DbQuery $conn "SELECT * FROM T_KOJIN1 WHERE KOJIN_ID IN ($kjs)" @{}) | Export-Csv (Join-Path $bkDir 'T_KOJIN1_書込前.csv') -NoTypeInformation -Encoding UTF8
    # 取り消し用のUPDATE
    $undo = @("-- $stamp の取込を元に戻す。健診ナビを閉じてから、内容を確認して実行すること。", 'BEGIN TRAN')
    foreach ($t in $plan) {
        foreach ($s in $t.Sets) {
            $oldv = $(if ($s.Old -eq '' -or $null -eq $s.Old) { 'NULL' } elseif ($s.Col -eq 'UKE_NO_KENSA') { [string]$s.Old } else { "'" + ([string]$s.Old -replace "'", "''") + "'" })
            if ($s.T -eq 'KENSIN') { $undo += ("UPDATE T_KENSIN SET [{0}] = {1} WHERE PK_SEQ = {2};" -f $s.Col, $oldv, $t.PK_SEQ) }
            else                   { $undo += ("UPDATE T_KOJIN1 SET [{0}] = {1} WHERE KOJIN_ID = '{2}';" -f $s.Col, $oldv, $t.KOJIN_ID) }
        }
    }
    $undo += @('-- 確認してから COMMIT', '-- COMMIT', '-- ROLLBACK')
    $undo | Out-File (Join-Path $bkDir '取り消し.sql') -Encoding Default

    $done = 0; $cells = 0
    $tran = $conn.BeginTransaction()
    try {
        Acquire-AppLock $conn $tran ('RECEPTION_IMPORT:' + $doc.series_id)
        foreach ($t in $plan) {
            foreach ($s in $t.Sets) {
                if ($s.T -eq 'KENSIN') {
                    $sql = "UPDATE T_KENSIN SET [$($s.Col)] = @v WHERE PK_SEQ = @pk"
                    $n = Invoke-DbExec $conn $tran $sql @{ v = $s.New; pk = $t.PK_SEQ }
                } else {
                    $sql = "UPDATE T_KOJIN1 SET [$($s.Col)] = @v WHERE KOJIN_ID = @kj"
                    $n = Invoke-DbExec $conn $tran $sql @{ v = $s.New; kj = $t.KOJIN_ID }
                }
                if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $s.What) }
                $cells++
            }
            $done++
        }
        $tran.Commit()
    }
    catch { $tran.Rollback(); throw }

    Write-Host ''
    Write-Host ("[書込] {0} 人 / {1} か所" -f $done, $cells) -ForegroundColor Green

    # ---- 読み直して確認 ----
    $bad = @()
    $chk = @(Invoke-DbQuery $conn $naviSql @{})
    $byPk = @{}; foreach ($r in $chk) { $byPk[[string]$r.PK_SEQ] = $r }
    foreach ($t in $plan) {
        $r = $byPk[[string]$t.PK_SEQ]
        if (-not $r) { $bad += "$($t.氏名) : 読み直せません"; continue }
        foreach ($s in $t.Sets) {
            $got = switch ($s.Col) {
                'UKE_NO_KENSA' { $(if ($r.UKE -is [System.DBNull]) { '' } else { [string][int]$r.UKE }) }
                'COURSE_CD'    { [string]$r.COURSE_CD }
                'KANA_SIMEI'   { [string]$r.KANA }
                default        { [string]$r.BIRTH }
            }
            if ((NoSpace $got) -ne (NoSpace ([string]$s.New))) { $bad += "$($t.氏名) : $($s.What) が [$got] (期待 [$($s.New)])" }
        }
    }
    if ($bad.Count -gt 0) { Write-Host ''; Write-Host '--- 確認NG ---' -ForegroundColor Yellow; $bad | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Yellow } }
    else { Write-Host '  読み直し確認: 全部そろっています' -ForegroundColor Green }

    # ---- 対応表と、今回取り込んだ印を残す ----
    $mapFile = Join-Path $BackupDir ("対応表_{0}.csv" -f $doc.series_id)
    $map | Select-Object series_id, 受診者ID, PK_SEQ, KOJIN_ID, 氏名, 事業所 | Export-Csv $mapFile -NoTypeInformation -Encoding UTF8
    @{ series_id = $doc.series_id; exported_at = $doc.exported_at; revision = $doc.revision; 取込日時 = $stamp } |
        ConvertTo-Json | Out-File $StateFile -Encoding UTF8

    Write-Host ''
    Write-Host ("[控え] 書込前の状態 : {0}" -f $bkDir) -ForegroundColor DarkGray
    Write-Host ("[控え] 取り消しSQL  : {0}" -f (Join-Path $bkDir '取り消し.sql')) -ForegroundColor DarkGray
    Write-Host ("[控え] 対応表       : {0}" -f $mapFile) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '※ オプション・便本数・採血/胃/尿の変更は、まだ反映していません (第2段階)。' -ForegroundColor Yellow
    Write-Host '※ 受診日が変わった人・キャンセルの人は、お知らせに出した内容を見て手で直してください。' -ForegroundColor Yellow
}
finally {
    if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
}
