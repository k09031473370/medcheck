<#
.SYNOPSIS
  東振協CSVの既往歴を健診ナビに取り込む (既往歴取込.ps1)

.DESCRIPTION
  東振協のCSVには既往歴が「病名ごとに1列」で入っている。
    既往歴(1) … 153〜172列
    既往歴(2) … 173〜192列
  どちらも中身は同じ20の病名で、該当する人はその列に印が立つ。
  どちらも「本人」の既往歴。

  健診ナビの既往歴は検査結果(T_KENSA)ではなく、人ごとの T_KIOU という
  別の表に入っている。KOJIN_ID + RENBAN で1行ずつ持つ形。
  結果報告書には **既往歴_状況1_1 〜 4_1 の4枠がある。

  CSVの文言をそのまま入れる。健診ナビの病名マスタに読み替えない。
    ・「胃・十二指腸の病気」のようなまとめ言葉も、そのまま残せる
    ・「脂質異常症」を「高脂血症」に寄せるような意味のずれが起きない
  そのかわり病名コード(BYOMEI_CD)は空になる。表示・印刷は問題ない。

  安全のために、こうしている。
    ・追加だけ。今ある行は消さない・書き換えない
    ・同じ病名が本人の既往歴として既に入っていれば飛ばす
    ・本人の既往歴が合計で -Max 件(既定4件)を超えないようにする
    ・T_KIOU の作り(列・NOT NULL・IDENTITY)をその場で読んでから組み立てる
    ・プレビューを出してから、その場で y を打つまで書き込まない
    ・書込前に T_KIOU を自動バックアップ、追加した行の一覧も残す
    ・健診ナビの結果入力画面で開いている人は書き換えない

  ★ 自動判定について
    健診ナビには T_KIOU_HANTEI という「病名 → 判定」の対応表がある。
    既往歴を入れたあとで自動判定をかけ直すと、対応表に載っている病名
    (高血圧・糖尿病 など) の項目の判定が変わる。
    かけ直さなければ、今の判定はそのままで、報告書に既往歴だけが増える。

.EXAMPLE
  既往歴取込.bat に 20260821.csv をドラッグ＆ドロップする

  powershell -ExecutionPolicy Bypass -File 既往歴取込.ps1 -Csv 20260821.csv
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Csv,                 # 東振協のCSV
    [string]$Ymd,                 # 受診日 (省略時はCSVの1列目から取る)
    [int]$Max = 4,                # 本人の既往歴の上限 (報告書の枠が4つ)
    [switch]$Commit,              # 付けると確認なしで書き込む
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

# form_import.ps1 から読み込み・接続・ロックまわりを借りる
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text'      'function Detect-MappingPath')
Invoke-Expression (Get-Part 'function Get-LocalConnFile'   'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Resolve-PkSeqByName' 'function Get-NaviName')
Invoke-Expression (Get-Part 'function Normalize-Name'      'function Select-TargetRows')
Invoke-Expression (Get-Part 'function Test-Locked'         'function Commit-Plan')
# Resolve-PkSeqByName が使う入れ物。借りた範囲の外で作られているので、ここで用意する。
$script:NaviByYmd = @{}

$BackupDir = Join-Path $PSScriptRoot 'backup'

# ---- CSVの既往歴の列 (東振協 250列版) ----
$COLS1 = 153..172      # 既往歴(1)
$COLS2 = 173..192      # 既往歴(2)
# 人の照合に使う列
$COL_YMD = 1; $COL_KANJI = 4; $COL_KANA = 5

function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }

# 型に合わせた「空」を返す。分からない型なら $null を返す (呼び側で止める)
function Empty-For([string]$t) {
    switch -Regex ($t.ToLower()) {
        '^(int|smallint|tinyint|bigint|bit|decimal|numeric|float|real|money|smallmoney)$' { return 0 }
        '^(char|varchar|nchar|nvarchar|text|ntext)$'                                      { return '' }
        '^(datetime|datetime2|smalldatetime|date)$'                                       { return (Get-Date) }
        default { return $null }
    }
}

# その列に入る長さを超えていないか
function Test-Fits([string]$s, $col) {
    if (-not $col) { return $true }
    $len = $col.Len
    if ($null -eq $len -or $len -lt 0) { return $true }
    if ($col.Type -match '^n') { return ($s.Length -le $len) }
    return ([System.Text.Encoding]::GetEncoding(932).GetByteCount($s) -le $len)
}

if (-not $Csv) { $Csv = Read-Host '東振協のCSVをドラッグ＆ドロップして Enter' }
$Csv = ($Csv -replace '^"|"$', '').Trim()
if (-not (Test-Path $Csv)) { throw "ファイルが見つかりません: $Csv" }

$rows = Read-FormCsv $Csv 'SJIS'
if ($rows.Count -lt 2) { throw "データが入っていません: $Csv" }
$head = $rows[0]
$data = $rows[1..($rows.Count - 1)]

if ($head.Count -lt 192) {
    throw ("列が足りません ({0}列)。東振協の250列のCSVを渡してください。" -f $head.Count)
}
# 見出しから病名を取る
$byomei = @{}
foreach ($c in ($COLS1 + $COLS2)) { $byomei[$c] = Normalize-Text $head[$c - 1] }

# (1)と(2)の見出しが同じ並びかを確かめる
$mismatch = @()
for ($i = 0; $i -lt $COLS1.Count; $i++) {
    $a = $byomei[$COLS1[$i]]; $b = $byomei[$COLS2[$i]]
    if ($a -ne $b) { $mismatch += ("{0}列[{1}] ≠ {2}列[{3}]" -f $COLS1[$i], $a, $COLS2[$i], $b) }
}

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 既往歴の取込' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($Csv))
Write-Host ("  入れる先: T_KIOU (人ごとの既往歴)  続柄=本人  1人 合計 {0} 件まで" -f $Max)
Write-Host '  CSVの文言をそのまま入れます (健診ナビの病名マスタに読み替えません)'
Write-Host ''
Write-Host ("  既往歴(1): {0}〜{1}列 / 既往歴(2): {2}〜{3}列  どちらも本人" `
    -f $COLS1[0], $COLS1[-1], $COLS2[0], $COLS2[-1]) -ForegroundColor DarkGray
Write-Host ("  病名: {0}" -f (($COLS1 | ForEach-Object { $byomei[$_] }) -join ' / ')) -ForegroundColor DarkGray
if ($mismatch.Count -gt 0) {
    Write-Host ''
    Write-Host '  ※ (1)と(2)の並びが違います。列の位置を確かめてください。' -ForegroundColor Yellow
    foreach ($m in $mismatch) { Write-Host ("     " + $m) -ForegroundColor Yellow }
}
Write-Host ''

if (-not $Ymd) { $Ymd = Normalize-Ymd (Get-Field $data[0] $COL_YMD) }
$Ymd = Normalize-Ymd $Ymd
if (-not $Ymd) { throw '受診日が読み取れません。-Ymd 2026/08/21 のように指定してください。' }
Write-Host ("  受診日: {0}" -f $Ymd)

$conn = Open-Db
try {
    # ---- T_KIOU の作りを読む ----
    $colsDt = Invoke-DbQuery $conn @"
SELECT c.COLUMN_NAME, c.DATA_TYPE, c.IS_NULLABLE, c.COLUMN_DEFAULT,
       c.CHARACTER_MAXIMUM_LENGTH AS MAXLEN,
       COLUMNPROPERTY(OBJECT_ID('T_KIOU'), c.COLUMN_NAME, 'IsIdentity') AS IS_IDENT,
       COLUMNPROPERTY(OBJECT_ID('T_KIOU'), c.COLUMN_NAME, 'IsComputed') AS IS_COMP
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME = 'T_KIOU'
ORDER BY c.ORDINAL_POSITION
"@
    if ($colsDt.Rows.Count -eq 0) { throw 'T_KIOU という表が見つかりません。' }
    $tbl = @{}
    foreach ($r in $colsDt.Rows) {
        $ml = -1
        if ($r.MAXLEN -ne [DBNull]::Value) { $ml = [int]$r.MAXLEN }
        $tbl[[string]$r.COLUMN_NAME] = New-Object PSObject -Property @{
            Name  = [string]$r.COLUMN_NAME
            Type  = ([string]$r.DATA_TYPE).ToLower()
            Null  = ([string]$r.IS_NULLABLE -eq 'YES')
            Def   = $(if ($r.COLUMN_DEFAULT -eq [DBNull]::Value) { '' } else { [string]$r.COLUMN_DEFAULT })
            Len   = $ml
            Ident = ($r.IS_IDENT -ne [DBNull]::Value -and [int]$r.IS_IDENT -eq 1)
            Comp  = ($r.IS_COMP  -ne [DBNull]::Value -and [int]$r.IS_COMP  -eq 1)
        }
    }
    foreach ($need in @('KOJIN_ID', 'RENBAN', 'BYOMEI')) {
        if (-not $tbl.ContainsKey($need)) { throw "T_KIOU に $need の列がありません。健診ナビの版が違います。" }
    }

    # ---- 続柄の使われ方を実データから読む ----
    $zkDt = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.ZOKUGARA),''))) AS CD,
       LTRIM(RTRIM(ISNULL(k.ZOKUGARAMEI,'')))                    AS MEI,
       COUNT(*) AS C
FROM T_KIOU k
GROUP BY LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.ZOKUGARA),''))),
         LTRIM(RTRIM(ISNULL(k.ZOKUGARAMEI,'')))
ORDER BY 3 DESC
"@
    $selfCd  = '0'
    $selfMei = '本人'
    $best = 0
    foreach ($r in $zkDt.Rows) {
        if ([string]$r.CD -eq '0' -and [int]$r.C -gt $best) { $best = [int]$r.C; $selfMei = [string]$r.MEI }
    }
    if ($selfMei -eq '') { $selfMei = '本人' }

    Write-Host ''
    Write-Host '=== いま T_KIOU に入っている続柄 ===' -ForegroundColor Cyan
    foreach ($r in $zkDt.Rows) {
        $mark = ''
        if ([string]$r.CD -eq $selfCd) { $mark = '  ← ここに入れます' }
        Write-Host ("  続柄CD[{0}] {1,-8} {2,6} 件{3}" -f $r.CD, $r.MEI, $r.C, $mark)
    }

    # ---- 受診者 (PK_SEQ ↔ KOJIN_ID) ----
    $kojinOf = @{}
    foreach ($r in (Invoke-DbQuery $conn `
        'SELECT PK_SEQ, KOJIN_ID FROM T_KENSIN WHERE D_KENSIN = @ymd AND F_TORIKESI = 0' `
        @{ ymd = $Ymd }).Rows) {
        $kojinOf[[string]$r.PK_SEQ] = [string]$r.KOJIN_ID
    }

    # ---- その人たちの今の既往歴 ----
    $exist = @{}
    foreach ($r in (Invoke-DbQuery $conn @"
SELECT k.KOJIN_ID,
       k.RENBAN,
       LTRIM(RTRIM(ISNULL(k.BYOMEI,'')))                          AS BYOMEI,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.ZOKUGARA),'')))  AS ZK
FROM T_KIOU k
WHERE k.KOJIN_ID IN (SELECT KOJIN_ID FROM T_KENSIN WHERE D_KENSIN = @ymd AND F_TORIKESI = 0)
"@ @{ ymd = $Ymd }).Rows) {
        $kj = [string]$r.KOJIN_ID
        if (-not $exist.ContainsKey($kj)) {
            $exist[$kj] = New-Object PSObject -Property @{ MaxRenban = 0; Self = @{}; SelfCount = 0 }
        }
        $n = 0
        if ([int]::TryParse(([string]$r.RENBAN).Trim(), [ref]$n)) {
            if ($n -gt $exist[$kj].MaxRenban) { $exist[$kj].MaxRenban = $n }
        }
        # 続柄が空か 0 の行を「本人の既往歴」として扱う
        if ([string]$r.ZK -eq '' -or [string]$r.ZK -eq $selfCd) {
            $exist[$kj].Self[[string]$r.BYOMEI] = 1
            $exist[$kj].SelfCount = $exist[$kj].SelfCount + 1
        }
    }

    # ---- 計画を立てる ----
    $plan    = @()
    $noHit   = 0
    $only1 = 0; $only2 = 0; $both = 0
    foreach ($f in $data) {
        $kanji = Normalize-Text (Get-Field $f $COL_KANJI)
        $kana  = Normalize-Text (Get-Field $f $COL_KANA)
        if ($kanji -eq '' -and $kana -eq '') { continue }
        $name = $(if ($kanji -ne '') { $kanji } else { $kana })

        # (1)(2) それぞれで印の立った病名を、列の順に拾う
        $h1 = @(); foreach ($c in $COLS1) { if ((Normalize-Text (Get-Field $f $c)) -ne '') { $h1 += $byomei[$c] } }
        $h2 = @(); foreach ($c in $COLS2) { if ((Normalize-Text (Get-Field $f $c)) -ne '') { $h2 += $byomei[$c] } }
        if ($h1.Count -eq 0 -and $h2.Count -eq 0) { $noHit++; continue }
        if ($h1.Count -gt 0 -and $h2.Count -gt 0) { $both++ }
        elseif ($h1.Count -gt 0) { $only1++ } else { $only2++ }

        # (1)を先に、(2)にしか無いものを後ろに足す。同じ病名は1回だけ。
        $hits = @(); $seen = @{}; $from = @{}
        foreach ($b in $h1) { if (-not $seen.ContainsKey($b)) { $seen[$b] = 1; $hits += $b; $from[$b] = '(1)' } }
        foreach ($b in $h2) {
            if ($seen.ContainsKey($b)) { $from[$b] = '(1)(2)' }
            else { $seen[$b] = 1; $hits += $b; $from[$b] = '(2)' }
        }

        $who = Resolve-PkSeqByName $conn $Ymd $kanji $kana
        if ($who.Status -ne 'OK') {
            $st = switch ($who.Status) {
                'NOTFOUND'  { '健診ナビに見つかりません' }
                'AMBIGUOUS' { "同名が{0}人いて特定できません" -f $who.Count }
                default     { $who.Status }
            }
            $plan += New-Object PSObject -Property @{
                氏名 = $name; 病名 = ($hits -join ' / '); 出所 = ''; 連番 = ''
                状態 = $st; PkSeq = $null; KojinId = $null }
            continue
        }
        $pk = [string]$who.PkSeq
        if (-not $kojinOf.ContainsKey($pk)) {
            $plan += New-Object PSObject -Property @{
                氏名 = $name; 病名 = ($hits -join ' / '); 出所 = ''; 連番 = ''
                状態 = '個人IDが取れません'; PkSeq = $pk; KojinId = $null }
            continue
        }
        $kj = $kojinOf[$pk]
        if (-not $exist.ContainsKey($kj)) {
            $exist[$kj] = New-Object PSObject -Property @{ MaxRenban = 0; Self = @{}; SelfCount = 0 }
        }
        $room   = $Max - $exist[$kj].SelfCount
        $renban = $exist[$kj].MaxRenban

        foreach ($b in $hits) {
            $rep = New-Object PSObject -Property @{
                氏名 = $name; 病名 = $b; 出所 = $from[$b]; 連番 = ''
                状態 = ''; PkSeq = $pk; KojinId = $kj }
            if ($exist[$kj].Self.ContainsKey($b)) { $rep.状態 = '既に入っています'; $plan += $rep; continue }
            if (-not (Test-Fits $b $tbl['BYOMEI'])) { $rep.状態 = '病名が長すぎます'; $plan += $rep; continue }
            if ($room -le 0) {
                $rep.状態 = ("本人の既往歴が既に{0}件あるので入れません" -f $Max); $plan += $rep; continue
            }
            $renban++
            $rep.連番 = $renban
            $rep.状態 = 'OK'
            $room--
            $exist[$kj].Self[$b] = 1
            $exist[$kj].SelfCount = $exist[$kj].SelfCount + 1
            $exist[$kj].MaxRenban = $renban
            $plan += $rep
        }
    }

    # ---- 出す ----
    Write-Host ''
    Write-Host '=== 既往歴(1)と(2)の関係 ===' -ForegroundColor Cyan
    Write-Host ("  (1)にだけ印: {0} 人 / (2)にだけ印: {1} 人 / 両方に印: {2} 人 / 既往歴なし: {3} 人" `
        -f $only1, $only2, $both, $noHit)

    Write-Host ''
    Write-Host '=== 入れる内容 ===' -ForegroundColor Cyan
    $plan | Select-Object 氏名, 病名, 出所, 連番, 状態 |
        Format-Table -AutoSize -Wrap | Out-String -Width 160 | Write-Host

    $ok   = @($plan | Where-Object { $_.状態 -eq 'OK' })
    $skip = @($plan | Where-Object { $_.状態 -eq '既に入っています' })
    $err  = @($plan | Where-Object { $_.状態 -notin @('OK', '既に入っています') })

    Write-Host '=== 病名ごとの件数 ===' -ForegroundColor Cyan
    foreach ($g in ($ok | Group-Object 病名 | Sort-Object Count -Descending)) {
        Write-Host ("  {0,3}件  {1}" -f $g.Count, $g.Name)
    }
    Write-Host ''
    Line
    Write-Host ("入れる: {0} 件 ({1} 人) / 既にあり: {2} 件 / 要確認: {3} 件 / 既往歴なし: {4} 人" `
        -f $ok.Count, (@($ok | Group-Object KojinId)).Count, $skip.Count, $err.Count, $noHit) `
        -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })
    Line

    # ---- 自動判定への影響を先に見せる ----
    if ($ok.Count -gt 0) {
        $names = @($ok | Select-Object -ExpandProperty 病名 -Unique)
        $inList = "N'" + (($names | ForEach-Object { $_ -replace "'", "''" }) -join "',N'") + "'"
        $hitDt = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(h.BYOMEI)) AS BYOMEI, LTRIM(RTRIM(h.KOMOKU_CD)) AS CD,
       LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) AS KIGO
FROM T_KIOU_HANTEI h
WHERE LTRIM(RTRIM(h.BYOMEI)) IN ($inList)
GROUP BY LTRIM(RTRIM(h.BYOMEI)), LTRIM(RTRIM(h.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,'')))
ORDER BY 1, 2
"@
        Write-Host ''
        Write-Host '=== 自動判定をかけ直したときに変わるもの ===' -ForegroundColor Yellow
        if ($hitDt.Rows.Count -eq 0) {
            Write-Host '  ありません。判定に使われる病名は今回の中にありませんでした。' -ForegroundColor Green
        } else {
            foreach ($r in $hitDt.Rows) {
                $n = @($ok | Where-Object { $_.病名 -eq [string]$r.BYOMEI }).Count
                Write-Host ("  {0,-14} {1,3}人  → 項目 {2} の判定が [{3}] になります" `
                    -f $r.BYOMEI, $n, $r.CD, $r.KIGO)
            }
            Write-Host '  ※ 自動判定をかけ直さなければ、今の判定はそのままです。' -ForegroundColor Yellow
        }
    }

    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ プレビューのみです。まだ何も書き込んでいません。' -ForegroundColor Yellow
        Write-Host ''
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    if ($ok.Count -eq 0) { Write-Host '入れるものがありません。'; return }

    # ---- INSERT の形を組み立てる ----
    #   IDENTITY と計算列は触らない。NOT NULL で既定値も無い列は、型に合う空で埋める。
    $fixed = @{}
    if ($tbl.ContainsKey('BYOMEI_CD'))   { $fixed['BYOMEI_CD']   = '' }
    if ($tbl.ContainsKey('CHIRYO'))      { $fixed['CHIRYO']      = '' }
    if ($tbl.ContainsKey('CHIRYOMEI'))   { $fixed['CHIRYOMEI']   = '' }
    if ($tbl.ContainsKey('BIKOU'))       { $fixed['BIKOU']       = '' }
    if ($tbl.ContainsKey('ZOKUGARA'))    {
        $fixed['ZOKUGARA'] = $(if ($tbl['ZOKUGARA'].Type -match '^(int|smallint|tinyint|bigint)$') { 0 } else { $selfCd })
    }
    if ($tbl.ContainsKey('ZOKUGARAMEI')) { $fixed['ZOKUGARAMEI'] = $selfMei }
    if ($tbl.ContainsKey('AGE'))         { $fixed['AGE']         = (Empty-For $tbl['AGE'].Type) }
    # 型が分からなくて空を決められなかったものは、はじめから入れない
    foreach ($k in @($fixed.Keys)) { if ($null -eq $fixed[$k]) { $fixed.Remove($k) } }

    $missing = @()
    foreach ($cn in $tbl.Keys) {
        $c = $tbl[$cn]
        if ($c.Ident -or $c.Comp) { continue }
        if ($cn -in @('KOJIN_ID', 'RENBAN', 'BYOMEI')) { continue }
        if ($fixed.ContainsKey($cn)) { continue }
        if ($c.Null) { continue }
        if ($c.Def -ne '') { continue }
        $e = Empty-For $c.Type
        if ($null -eq $e) { $missing += ("{0} ({1})" -f $cn, $c.Type); continue }
        $fixed[$cn] = $e
    }
    if ($missing.Count -gt 0) {
        throw ("T_KIOU に、何を入れればよいか分からない必須の列があります。書き込みません。`n  " + ($missing -join "`n  "))
    }

    $insCols = @('KOJIN_ID', 'RENBAN', 'BYOMEI') + @($fixed.Keys | Sort-Object)
    $insSql  = 'INSERT INTO T_KIOU (' + ($insCols -join ', ') + ') VALUES (' +
               (($insCols | ForEach-Object { '@' + $_ }) -join ', ') + ')'
    Write-Host ''
    Write-Host ("[SQL] " + $insSql) -ForegroundColor DarkGray

    # ---- 書き込む ----
    if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $undo    = Join-Path $BackupDir ("T_KIOU_追加分_{0}.csv" -f $stamp)
    $added   = @()
    $done    = 0

    foreach ($grp in ($ok | Group-Object KojinId)) {
        $kj = $grp.Name
        $pk = $grp.Group[0].PkSeq
        $lock = Test-Locked $conn $pk
        if ($lock) {
            Write-Host ("  [とばしました] {0} は健診ナビで編集中です ({1})" -f $grp.Group[0].氏名, $lock) -ForegroundColor Yellow
            continue
        }
        # 書込前の T_KIOU を丸ごと控える
        $bk = Join-Path $BackupDir ("T_KIOU_{0}_{1}.csv" -f $kj, $stamp)
        (Invoke-DbQuery $conn 'SELECT * FROM T_KIOU WHERE KOJIN_ID = @k' @{ k = $kj }) |
            Export-Csv -Path $bk -NoTypeInformation -Encoding UTF8

        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KIOU_IMPORT:" + $kj)
            $lock2 = Test-Locked $conn $pk $tran
            if ($lock2) { throw "書込直前に健診ナビで開かれました ($lock2)" }
            foreach ($t in $grp.Group) {
                $p = @{}
                foreach ($k in $fixed.Keys) { $p[$k] = $fixed[$k] }
                $p['KOJIN_ID'] = $kj
                $p['RENBAN']   = $(if ($tbl['RENBAN'].Type -match '^(int|smallint|tinyint|bigint)$') { [int]$t.連番 } else { [string]$t.連番 })
                $p['BYOMEI']   = $t.病名
                $n = Invoke-DbExec $conn $tran $insSql $p
                if ($n -ne 1) { throw ("INSERT影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.病名) }
                $added += New-Object PSObject -Property @{
                    KOJIN_ID = $kj; RENBAN = $t.連番; BYOMEI = $t.病名; 氏名 = $t.氏名 }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [書込] {0} : {1} 件" -f $grp.Group[0].氏名, $grp.Count) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }

    if ($added.Count -gt 0) {
        $added | Select-Object 氏名, KOJIN_ID, RENBAN, BYOMEI |
            Export-Csv -Path $undo -NoTypeInformation -Encoding UTF8
        Write-Host ("[控え] 追加した行の一覧: {0}" -f $undo) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host ("[完了] {0} 件の既往歴を入れました。" -f $done) -ForegroundColor Green
    Write-Host ''
    Write-Host '※ 結果報告書の「既往歴」欄に出ます (4枠まで)。' -ForegroundColor Yellow
    Write-Host '※ 病名コード(BYOMEI_CD)は空です。文字だけ入っています。' -ForegroundColor Yellow
    Write-Host '※ 自動判定をかけ直すと、上に出した項目の判定が変わります。' -ForegroundColor Yellow
    Write-Host '   今の判定のままにしたいなら、かけ直さないでください。' -ForegroundColor Yellow
}
finally { $conn.Close() }
