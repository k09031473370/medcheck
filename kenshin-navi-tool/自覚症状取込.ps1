<#
.SYNOPSIS
  東振協CSVの自覚症状を健診ナビに取り込む (自覚症状取込.ps1)

.DESCRIPTION
  東振協のCSVには自覚症状が「症状ごとに1列」で入っている (195〜207列)。
  該当する人はその列に 1 が立つ。
  健診ナビは 017107A〜J の10枠に「該当したものを順に詰める」形なので、
  該当した症状を前から順に枠へ入れる。

  CSVの文言をそのまま入れる。健診ナビの選択肢に読み替えない。
    ・「鼻づまり・鼻汁がでる」のように選択肢に無い症状も、そのまま残せる
    ・「顔や足がむくむ」を「手足がむくむ」に寄せるような意味のずれが起きない
  そのかわり選択肢に無い文言になるので、健診ナビの画面でその欄を
  選び直すと消える。表示・印刷は問題ない。

  安全のために、こうしている。
    ・空いている枠にだけ入れる。すでに入っているものは消さない
    ・同じ症状がすでに入っていれば飛ばす
    ・プレビューを出してから、その場で y を打つまで書き込まない
    ・書込前に T_KENSA を自動バックアップ
    ・健診ナビの結果入力画面で開いている人は書き換えない

.EXAMPLE
  自覚症状取込.bat に 20260821.csv をドラッグ＆ドロップする

  powershell -ExecutionPolicy Bypass -File 自覚症状取込.ps1 -Csv 20260821.csv
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Csv,                 # 東振協のCSV
    [string]$Ymd,                 # 受診日 (省略時はCSVの1列目から取る)
    [int]$Max = 4,                # 1人に入れる上限 (前から何個まで)
    [switch]$Commit,              # 付けると確認なしで書き込む
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

# form_import.ps1 から読み込み・接続・バックアップまわりを借りる
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
# Normalize-* / Read-TextShared / Parse-CsvText / Read-FormCsv / Get-Field まで
#   途中に「$XlsxLib = Join-Path $PSScriptRoot ...」という関数の外の行がある。
#   Invoke-Expression の中では $PSScriptRoot が空なのでそこで止まる。
#   CSVしか読まないので Excel の部品は要らない。その3行を飛ばして2回に分けて借りる。
Invoke-Expression (Get-Part 'function Normalize-Text'      '$XlsxLib = Join-Path')
Invoke-Expression (Get-Part 'function Test-XlsxEncrypted'  'function Detect-MappingPath')
Invoke-Expression (Get-Part 'function Get-LocalConnFile'   'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Resolve-PkSeqByName' 'function Get-NaviName')
Invoke-Expression (Get-Part 'function Normalize-Name'      'function Select-TargetRows')
Invoke-Expression (Get-Part 'function Test-Locked'         'function Commit-Plan')
# Resolve-PkSeqByName が使う入れ物。借りた範囲の外で作られているので、ここで用意する。
$script:NaviByYmd = @{}

$BackupDir = Join-Path $PSScriptRoot 'backup'

# ---- CSVの自覚症状の列 (東振協 250列版) ----
#   見出しの文字をそのまま健診ナビに入れる
$COLS = 195..207
# ---- 健診ナビの受け皿 ----
$SLOTS = @('017107A','017107B','017107C','017107D','017107E',
           '017107F','017107G','017107H','017107I','017107J')
# 人の照合に使う列 (東振協 250列版)
$COL_YMD = 1; $COL_KANJI = 4; $COL_KANA = 5

function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }

if (-not $Csv) { $Csv = Read-Host '東振協のCSVをドラッグ＆ドロップして Enter' }
$Csv = ($Csv -replace '^"|"$', '').Trim()
if (-not (Test-Path $Csv)) { throw "ファイルが見つかりません: $Csv" }

$rows = Read-FormCsv $Csv 'SJIS'
if ($rows.Count -lt 2) { throw "データが入っていません: $Csv" }
$head = $rows[0]
$data = $rows[1..($rows.Count - 1)]

# 見出しから症状名を取る。列数が違うCSVを渡されたら止める。
if ($head.Count -lt 207) {
    throw ("列が足りません ({0}列)。東振協の250列のCSVを渡してください。" -f $head.Count)
}
$symptom = @{}
foreach ($c in $COLS) { $symptom[$c] = Normalize-Text $head[$c - 1] }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 自覚症状の取込' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($Csv))
Write-Host ("  入れる先: 017107A〜J の空いている枠 (1人 最大 {0} 件)" -f $Max)
Write-Host '  CSVの文言をそのまま入れます (健診ナビの選択肢に読み替えません)'
Write-Host ''
Write-Host '  対象の列:' -ForegroundColor DarkGray
foreach ($c in $COLS) { Write-Host ("    {0,3}列  {1}" -f $c, $symptom[$c]) -ForegroundColor DarkGray }
Write-Host ''

if (-not $Ymd) { $Ymd = Normalize-Ymd (Get-Field $data[0] $COL_YMD) }
$Ymd = Normalize-Ymd $Ymd
if (-not $Ymd) { throw '受診日が読み取れません。-Ymd 2026/08/21 のように指定してください。' }
Write-Host ("  受診日: {0}" -f $Ymd)

$conn = Open-Db
try {
    # ---- いま入っている自覚症状を読む ----
    $slotList = "'" + ($SLOTS -join "','") + "'"
    $cur = @{}
    foreach ($r in (Invoke-DbQuery $conn @"
SELECT k.PK_SEQ, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, ISNULL(k.KEKKA,'') AS V
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($slotList)
"@ @{ ymd = $Ymd }).Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $cur.ContainsKey($pk)) { $cur[$pk] = @{} }
        $cur[$pk][[string]$r.CD] = Normalize-Text ([string]$r.V)
    }

    # ---- 計画を立てる ----
    $plan = @()
    $noHit = 0
    foreach ($f in $data) {
        $kanji = Normalize-Text (Get-Field $f $COL_KANJI)
        $kana  = Normalize-Text (Get-Field $f $COL_KANA)
        if ($kanji -eq '' -and $kana -eq '') { continue }

        # この人の症状を、列の順に拾う
        $hits = @()
        foreach ($c in $COLS) {
            if ((Normalize-Text (Get-Field $f $c)) -ne '') { $hits += $symptom[$c] }
        }
        if ($hits.Count -eq 0) { $noHit++; continue }
        $hits = @($hits | Select-Object -First $Max)

        $who = Resolve-PkSeqByName $conn $Ymd $kanji $kana
        $name = if ($kanji -ne '') { $kanji } else { $kana }
        if ($who.Status -ne 'OK') {
            $st = switch ($who.Status) {
                'NOTFOUND'  { '健診ナビに見つかりません' }
                'AMBIGUOUS' { "同名が{0}人いて特定できません" -f $who.Count }
                default     { $who.Status }
            }
            $plan += New-Object PSObject -Property @{
                氏名 = $name; 症状 = ($hits -join ' / '); 枠 = ''; 状態 = $st; PkSeq = $null; 文言 = '' }
            continue
        }
        $pk = [string]$who.PkSeq
        if (-not $cur.ContainsKey($pk)) {
            $plan += New-Object PSObject -Property @{
                氏名 = $name; 症状 = ($hits -join ' / '); 枠 = ''; 状態 = '自覚症状の枠がありません'; PkSeq = $pk; 文言 = '' }
            continue
        }

        $used = @{}
        foreach ($h in $hits) {
            $rep = New-Object PSObject -Property @{
                氏名 = $name; 症状 = $h; 枠 = ''; 状態 = ''; PkSeq = $pk; 文言 = $h }
            # すでに同じ症状が入っていないか
            $dup = $null
            foreach ($sc in $SLOTS) {
                if ($cur[$pk].ContainsKey($sc) -and $cur[$pk][$sc] -eq $h) { $dup = $sc; break }
            }
            if ($dup) { $rep.枠 = $dup; $rep.状態 = '既に入っています'; $plan += $rep; continue }
            # 空いている枠を探す
            $slot = $null
            foreach ($sc in $SLOTS) {
                if (-not $cur[$pk].ContainsKey($sc)) { continue }
                if ($used.ContainsKey($sc)) { continue }
                $v = $cur[$pk][$sc]
                if ($v -eq '' -or $v -eq '#') { $slot = $sc; break }
            }
            if (-not $slot) { $rep.状態 = '空いている枠がありません'; $plan += $rep; continue }
            $used[$slot] = 1
            $rep.枠 = $slot; $rep.状態 = 'OK'
            $plan += $rep
        }
    }

    # ---- 出す ----
    Write-Host ''
    Write-Host '=== 入れる内容 ===' -ForegroundColor Cyan
    $plan | Select-Object 氏名, 症状, 枠, 状態 |
        Format-Table -AutoSize -Wrap | Out-String -Width 160 | Write-Host

    $ok   = @($plan | Where-Object { $_.状態 -eq 'OK' })
    $skip = @($plan | Where-Object { $_.状態 -eq '既に入っています' })
    $err  = @($plan | Where-Object { $_.状態 -notin @('OK', '既に入っています') })
    Write-Host '=== 症状ごとの件数 ===' -ForegroundColor Cyan
    foreach ($g in ($ok | Group-Object 文言 | Sort-Object Count -Descending)) {
        Write-Host ("  {0,3}件  {1}" -f $g.Count, $g.Name)
    }
    Write-Host ''
    Line
    Write-Host ("入れる: {0} 件 / 既にあり: {1} 件 / 要確認: {2} 件 / 症状なし: {3} 人" `
        -f $ok.Count, $skip.Count, $err.Count, $noHit) `
        -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })
    Line
    Write-Host '※ 症状が1つも無い人には何も入れません (「特記事項なし」も入れません)' -ForegroundColor DarkGray

    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ プレビューのみです。まだ何も書き込んでいません。' -ForegroundColor Yellow
        Write-Host ''
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    if ($ok.Count -eq 0) { Write-Host '入れるものがありません。'; return }

    # ---- 書き込む ----
    $done = 0
    foreach ($grp in ($ok | Group-Object PkSeq)) {
        $pk = $grp.Name
        $lock = Test-Locked $conn $pk
        if ($lock) {
            Write-Host ("  [とばしました] {0} は健診ナビで編集中です ({1})" -f $grp.Group[0].氏名, $lock) -ForegroundColor Yellow
            continue
        }
        [void](Backup-Kensa $conn $pk)
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KENSA_IMPORT:" + $pk)
            $lock2 = Test-Locked $conn $pk $tran
            if ($lock2) { throw "書込直前に健診ナビで開かれました ($lock2)" }
            foreach ($t in $grp.Group) {
                $n = Invoke-DbExec $conn $tran `
                    'UPDATE T_KENSA SET KEKKA = @v WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' `
                    @{ v = $t.文言; p = $pk; cd = $t.枠 }
                if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.枠) }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [書込] {0} : {1} 件" -f $grp.Group[0].氏名, $grp.Count) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }

    Write-Host ''
    Write-Host ("[完了] {0} 件の自覚症状を入れました。" -f $done) -ForegroundColor Green
    Write-Host ''
    Write-Host '※ 結果報告書の「自覚症状」欄に出ます (4枠まで)。' -ForegroundColor Yellow
    Write-Host '※ 健診ナビの画面でこの欄を選び直すと、選択肢に無い文言は消えます。' -ForegroundColor Yellow
}
finally { $conn.Close() }
