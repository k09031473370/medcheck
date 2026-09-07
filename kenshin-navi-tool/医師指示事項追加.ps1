<#
.SYNOPSIS
  判定に応じた定型文を医師指示事項に追加する (医師指示事項追加.ps1)

.DESCRIPTION
  胸部X線・胃部X線・心電図の判定に応じて、決まった文章を
  医師指示事項 (総合判定N(編集) = 11001〜11009) の空いている枠に入れる。

    胸部X線  D → 308   F → 309
    胃部X線  D →  75   F →  69
    心電図   D →  58   F →  68

  ※ D・F は健診ナビの判定記号。東振協の C が D、東振協の D が F。

  安全のために、こうしています。
    ・空いている枠にだけ入れる。すでに入っている文章は絶対に消さない
    ・同じ番号の文章がすでにあれば飛ばす
    ・プレビューで全員分を出してから、別に書込を実行する
    ・書込前に T_KENSA を自動バックアップ
    ・健診ナビの結果入力画面で開いている人は書き換えない

  医師の指示欄なので、必ずプレビューを目で見てから書き込んでください。

.EXAMPLE
  医師指示事項追加.bat をダブルクリックする

  powershell -ExecutionPolicy Bypass -File 医師指示事項追加.ps1 -Ymd 2026/08/21
  powershell -ExecutionPolicy Bypass -File 医師指示事項追加.ps1 -Ymd 2026/08/21 -Commit
#>
[CmdletBinding()]
param(
    [string]$Ymd,          # 受診日 (2026/08/21)
    [switch]$Commit,       # 付けると実際に書き込む。付けなければプレビューだけ
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

# form_import.ps1 から接続・バックアップまわりを借りる (同じ作法で扱うため)
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

# ---- 入れる文章 ----
#   Item   … 判定を見る項目 (所見の1枠目)
#   D / F  … 健診ナビの判定記号ごとの、入れる番号
$RULES = @(
    @{ Name = '胸部X線'; Item = '077010A'; D = '308'; F = '309' }
    @{ Name = '胃部X線'; Item = '077300A'; D = '75';  F = '69'  }
    @{ Name = '心電図';  Item = '067112A'; D = '58';  F = '68'  }
)
$TEXT = @{
    '308' = '胸部X線検査にて異常所見を認めます。経過観察の必要があります。自覚症状があれば検査を受けて下さい。'
    '309' = '胸部X線検査にて異常所見を認めます。精密検査の必要があります。'
    '58'  = '心電図異常があります。経過観察を行い、自覚症状があれば受診して下さい。'
    '68'  = '心電図異常については、循環器外来を受診し、精査・治療が必要です。'
    '75'  = '胃部レントゲン検査にて異常所見を認めます。経過観察の必要があります。自覚症状があれば検査を受けて下さい。'
    '69'  = '胃部レントゲン検査にて異常所見を認めます。精密検査の必要があります。'
}
# 医師指示事項の枠 (総合判定N(編集))
$SLOTS = @('11001','11002','11003','11004','11005','11006','11007','11008','11009')

if (-not $Ymd) { $Ymd = Read-Host '受診日を入れてください (例 2026/08/21)' }
$Ymd = Normalize-Ymd $Ymd
if (-not $Ymd) { throw '受診日を 2026/08/21 のように入れてください。' }

function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 医師指示事項に定型文を追加' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host "  受診日: $Ymd"
Write-Host '  入れる先: 総合判定N(編集) 11001〜11009 の空いている枠'
Write-Host ''
foreach ($r in $RULES) {
    Write-Host ('    {0,-8} 判定D → {1}番 / 判定F → {2}番' -f $r.Name, $r.D, $r.F)
}
Write-Host ''

$conn = Open-Db
try {
    # ---- 判定が健診ナビの記号になっているかを先に確かめる ----
    # 東振協の C をそのまま入れたままだと、意味が食い違う (東振協C = 健診ナビD)。
    $items = "'" + (($RULES | ForEach-Object { $_.Item }) -join "','") + "'"
    $bad = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(k.HANTEI_KIGO)) AS H, COUNT(*) AS N
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($items)
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'C'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), LTRIM(RTRIM(k.HANTEI_KIGO))
"@ @{ ymd = $Ymd }
    if ($bad.Rows.Count -gt 0) {
        Write-Host '  ★止めます。判定が東振協の記号のままです。' -ForegroundColor Red
        foreach ($r in $bad.Rows) { Write-Host ("     {0} に判定C が {1} 件" -f $r.CD, $r.N) -ForegroundColor Red }
        Write-Host ''
        Write-Host '  健診ナビでは 東振協C = D / 東振協D = F です。' -ForegroundColor Yellow
        Write-Host '  先に mapping_tosinkyo250.csv と code_map.csv を差し替えて' -ForegroundColor Yellow
        Write-Host '  取込をやり直してから、もう一度実行してください。' -ForegroundColor Yellow
        return
    }

    # ---- 対象者と、その判定 ----
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, ISNULL(g.KANJI_SIMEI,'') AS 氏名, ISNULL(g.KANA_SIMEI,'') AS カナ,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H,
       LEFT(ISNULL(k.KEKKA,''), 24) AS 所見
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($items)
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('D','F')
ORDER BY g.KANA_SIMEI, k.KOMOKU_CD
"@ @{ ymd = $Ymd }).Rows

    if ($rows.Count -eq 0) {
        Write-Host '  判定が D または F の人はいませんでした。入れるものがありません。' -ForegroundColor Yellow
        return
    }

    # ---- いま入っている医師指示事項を読む ----
    $slotList = "'" + ($SLOTS -join "','") + "'"
    $cur = @{}   # PK_SEQ → @{ 枠CD → @{ Kekka; KekkaCd } }
    foreach ($r in (Invoke-DbQuery $conn @"
SELECT k.PK_SEQ, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD,
       ISNULL(k.KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($slotList)
"@ @{ ymd = $Ymd }).Rows) {
        $pk = [string]$r.PK_SEQ
        if (-not $cur.ContainsKey($pk)) { $cur[$pk] = @{} }
        $cur[$pk][[string]$r.CD] = @{ Kekka = (Normalize-Text ([string]$r.V)); KekkaCd = [string]$r.VC }
    }

    # ---- 計画を立てる ----
    $plan = @()
    $used = @{}    # PK_SEQ → 使う予定の枠 (同じ人に2つ入れるとき用)
    foreach ($r in $rows) {
        $pk    = [string]$r.PK_SEQ
        $cd    = [string]$r.CD
        $h     = [string]$r.H
        $rule  = $RULES | Where-Object { $_.Item -eq $cd } | Select-Object -First 1
        $no    = if ($h -eq 'D') { $rule.D } else { $rule.F }
        $body  = $TEXT[$no]

        $rep = New-Object PSObject -Property @{
            氏名 = [string]$r.氏名; 検査 = $rule.Name; 判定 = $h
            所見 = (Normalize-Text ([string]$r.所見))
            番号 = $no; 枠 = ''; 状態 = ''
            PkSeq = $pk; 文章 = $body
        }

        if (-not $cur.ContainsKey($pk)) { $rep.状態 = '医師指示事項の枠がありません'; $plan += $rep; continue }

        # すでに同じ番号か同じ文章が入っていないか
        $dup = $false
        foreach ($sc in $SLOTS) {
            if (-not $cur[$pk].ContainsKey($sc)) { continue }
            $c = $cur[$pk][$sc]
            if ($c.KekkaCd -eq $no)   { $dup = $true; $rep.枠 = $sc; break }
            if ($c.Kekka -eq $body)   { $dup = $true; $rep.枠 = $sc; break }
        }
        if ($dup) { $rep.状態 = '既に入っています'; $plan += $rep; continue }

        # 空いている枠を探す (この実行で使う予定の枠も埋まっているものとして扱う)
        if (-not $used.ContainsKey($pk)) { $used[$pk] = @{} }
        $slot = $null
        foreach ($sc in $SLOTS) {
            if (-not $cur[$pk].ContainsKey($sc)) { continue }        # その枠の行が無い
            if ($used[$pk].ContainsKey($sc))     { continue }        # この実行で使う予定
            $v = $cur[$pk][$sc].Kekka
            if ($v -eq '' -or $v -eq '#') { $slot = $sc; break }
        }
        if (-not $slot) { $rep.状態 = '空いている枠がありません'; $plan += $rep; continue }

        $used[$pk][$slot] = 1
        $rep.枠 = $slot
        $rep.状態 = 'OK'
        $plan += $rep
    }

    # ---- 出す ----
    Write-Host ''
    Write-Host '=== 追加する内容 ===' -ForegroundColor Cyan
    $plan | Select-Object 氏名, 検査, 判定, 番号, 枠, 状態, 所見 |
        Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host

    Write-Host '=== 入れる文章 ===' -ForegroundColor Cyan
    foreach ($no in ($plan | Where-Object { $_.状態 -eq 'OK' } | ForEach-Object { $_.番号 } | Sort-Object -Unique)) {
        $n = @($plan | Where-Object { $_.状態 -eq 'OK' -and $_.番号 -eq $no }).Count
        Write-Host ("  {0,-4} ({1}件)  {2}" -f $no, $n, $TEXT[$no])
    }

    $ok   = @($plan | Where-Object { $_.状態 -eq 'OK' })
    $skip = @($plan | Where-Object { $_.状態 -eq '既に入っています' })
    $err  = @($plan | Where-Object { $_.状態 -notin @('OK', '既に入っています') })
    Write-Host ''
    Line
    Write-Host ("追加: {0} 件 / 既にあり: {1} 件 / 要確認: {2} 件" -f $ok.Count, $skip.Count, $err.Count) `
        -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })
    Line

    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ プレビューのみです。まだ何も書き込んでいません。' -ForegroundColor Yellow
        Write-Host '   上の内容でよければ、もう一度実行して「書き込む」を選んでください。' -ForegroundColor Yellow
        Write-Host ''
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }
    if ($ok.Count -eq 0) { Write-Host '追加するものがありません。'; return }

    # ---- 書き込む ----
    $byPerson = $ok | Group-Object PkSeq
    $done = 0
    foreach ($grp in $byPerson) {
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
                    'UPDATE T_KENSA SET KEKKA = @v, KEKKA_CD = @c WHERE PK_SEQ = @p AND LTRIM(RTRIM(KOMOKU_CD)) = @cd' `
                    @{ v = $t.文章; c = $t.番号; p = $pk; cd = $t.枠 }
                if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.枠) }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [書込] {0} : {1} 件" -f $grp.Group[0].氏名, $grp.Count) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }

    Write-Host ''
    Write-Host ("[完了] {0} 件を医師指示事項に追加しました。" -f $done) -ForegroundColor Green
    Write-Host ''
    Write-Host '※ 健診ナビで結果報告書を開いて、文章が出ているか確かめてください。' -ForegroundColor Yellow
    Write-Host '※ このあと自動判定を回すと、追加した文章が消える可能性があります。' -ForegroundColor Yellow
    Write-Host '   自動判定は先に済ませてから、この作業をしてください。' -ForegroundColor Yellow
}
finally { $conn.Close() }
