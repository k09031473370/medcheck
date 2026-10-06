<#
.SYNOPSIS
  受付アプリのJSONから、健診ナビの予約取込用Excel(33列)を作る (予約取込Excel作成.ps1)

.DESCRIPTION
  会場の受付アプリが出す navi.json を読んで、
  まだ健診ナビに入っていない人だけの予約取込ファイルを作る。

  なぜ要るか
    ・健診ナビの予約取込はカナが無いと弾かれる (10/2 は103人中17人が入らなかった)
    ・当日会場で確認したカナは受付アプリの中にある。それを取り出してExcelにする
    ・受付番号が入らないと血液の検査依頼が出せないので、会期当日に間に合わせる必要がある

  やること
    1. JSONを読む (identity.kana / identity.birth_date が当日確認したカナ・生年月日)
    2. 健診ナビを読んで、もう入っている人を外す (氏名で突き合わせ。二重登録を防ぐ)
    3. 事業所コードを form\fussa_company_map.csv から引く
    4. コース A/B/C を FA/FB/FC に変換 (受付済みの人は当日のコース)
    5. オプションを form\fussa_option_map.csv で健診ナビのコード/名称に変換して
       「オプション検査1〜10」の列に書く (受付済みの人は当日のオプション。O 溶接ヒュームは書かない)
    6. 33列のExcelを作る
    7. Excelでは入れられないもの (採血/胃部/尿の変更・便の本数・受付メモ・コース変更の経緯) を
       「確認一覧.txt」に出す → 健診ナビの画面で対応してもらう

  同じ人を二度登録しない仕組み
    健診ナビに同じ氏名の人が既にいれば出さない (2)。
    その人の受付番号・カナ・生年月日は 受付結果取込.bat が既存の予約に書く。
    つまり「無い人はこのツールで新規」「ある人は受付結果取込で更新」に自動で分かれる。

  出さない人 (一覧に理由を出す)
    ・もう健診ナビに入っている人
    ・カナが空の人            … 取り込んでも弾かれるので
    ・事業所コードが引けない人  … 対応表にない会社
    ・コースが A/B/C 以外の人   … NONE や未知の記号
    ・受付アプリでキャンセルの人

  DBは読むだけ。健診ナビには何も書きません。
  できたExcelを健診ナビの予約取込にかけてください。

.EXAMPLE
  予約取込Excel作成.bat に navi.json をドラッグ＆ドロップ

  # 10/2 に実際に受付した人だけ出す (血液の検査依頼を出すため)
  powershell -ExecutionPolicy Bypass -File 予約取込Excel作成.ps1 -Json ...\navi.json -Ymd 2026/10/02
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Json,                      # 受付アプリが出した navi.json
    [string]$Ymd,                       # この日に実際に受付した人だけ出す (省略時は全員)
    [string]$Out,                       # 出力先フォルダ (既定: デスクトップ)
    [string]$DantaiLike = '福生',       # 健診ナビ側で既に入っている人をさがす範囲
    [string]$MapCsv,                    # 省略時は form\fussa_company_map.csv
    [string]$OptCsv,                    # 省略時は form\fussa_option_map.csv (記号 → 健診ナビのコード/名称)
    [ValidateSet('Code','Name')]
    [string]$OptionAs = 'Code',         # オプション検査の列に何を書くか。Code=OPJ014  Name=乳腺超音波検査
    [switch]$NoOptions,                 # 付けると、オプション検査の列を空にする (10/2 と同じ形。健診ナビの画面で付ける)
    [switch]$IncludeNoKana,             # 付けると、カナが無い人も出す (取込で弾かれます)
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
Invoke-Expression (Get-Part 'function Normalize-Text'    '$XlsxLib = Join-Path')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')

if (-not $MapCsv) { $MapCsv = Join-Path $PSScriptRoot 'form\fussa_company_map.csv' }
if (-not $OptCsv) { $OptCsv = Join-Path $PSScriptRoot 'form\fussa_option_map.csv' }

# 受付アプリのコース記号 → 健診ナビ
$COURSE = @{
    'A' = @{ CD = 'FA'; Name = '福生A' }
    'B' = @{ CD = 'FB'; Name = '福生B' }
    'C' = @{ CD = 'FC'; Name = '福生C' }
}
# 健診ナビの予約取込の見出し (33列)
$HEAD = @('No','氏名','氏名カナ','性別','生年月日','郵便番号','住所1(都道府県)','住所2(市区町村・丁目・番地)',
          '住所3(建物名・部屋番号)','電話番号','カルテ番号','保険者番号','保険証記号','保険証番号','保険証枝番号',
          '健保コード','健保名','事業所コード','事業所名','コースコード','コース名','予約日','予約時間') +
         (1..10 | ForEach-Object { "オプション検査$_" })

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
function ToNarrowKana([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return '' }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $n = [Microsoft.VisualBasic.Strings]::StrConv($s, [Microsoft.VisualBasic.VbStrConv]::Narrow, 1041)
    return ($n -replace '\s+', ' ').Trim()
}
function StartTime([string]$slot) {
    # 「10:00～10:15」→「10:00」
    if ($slot -match '(\d{1,2}:\d{2})') { return $Matches[1] }
    return ''
}

# ============================================================================
# 入力
# ============================================================================
if (-not $Json) { $Json = Read-Host '受付アプリのJSONをドラッグ＆ドロップして Enter' }
# 日付は「Date 2026/10/06」「2026-10-06」のように打たれても、日付の部分だけ拾う
if ($Ymd) {
    if ($Ymd -match '(\d{4})[/\-\.年](\d{1,2})[/\-\.月](\d{1,2})') {
        $Ymd = '{0}/{1:00}/{2:00}' -f $Matches[1], [int]$Matches[2], [int]$Matches[3]
    } else { throw "日付が読めません: [$Ymd]  2026/10/06 のように入れてください" }
}
$Json = ($Json -replace '^"|"$', '').Trim()
if (-not (Test-Path $Json)) { throw "見つかりません: $Json" }
if (-not (Test-Path $MapCsv)) { throw "事業所の対応表がありません: $MapCsv" }
if (-not (Test-Path $OptCsv)) { throw "オプションの対応表がありません: $OptCsv" }
if (-not $Out) { $Out = [Environment]::GetFolderPath('Desktop') }
if (-not (Test-Path $Out)) { throw "出力先がありません: $Out" }

$doc = Get-Content -LiteralPath $Json -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $doc.people) { throw 'このJSONに people がありません。受付アプリの『健診ナビ連携JSONを出力』で出したファイルを渡してください。' }

# 事業所名 → 団体CD
$cdOf = @{}
foreach ($r in (Import-Csv -Path $MapCsv -Encoding UTF8)) {
    $k = NoSpace $r.'振り分け表の事業所名'
    if ($k -ne '' -and ([string]$r.団体CD).Trim() -ne '') { $cdOf[$k] = ([string]$r.団体CD).Trim() }
}

# オプション記号 → 健診ナビ
$optOf = @{}
foreach ($r in (Import-Csv -Path $OptCsv -Encoding UTF8)) {
    $sym = ([string]$r.記号).Trim()
    if ($sym -eq '') { continue }
    $optOf[$sym] = @{ CD = ([string]$r.コード).Trim(); Name = ([string]$r.健診ナビの名称).Trim(); App = ([string]$r.受付アプリの名称).Trim() }
}

Write-Host ''
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ' 受付アプリのJSON → 健診ナビの予約取込Excel (33列)' -ForegroundColor Cyan
Write-Host ('=' * 78) -ForegroundColor Cyan
Write-Host ("  ファイル : {0}" -f (Split-Path $Json -Leaf))
Write-Host ("  出力日時 : {0}   全{1}人" -f $doc.exported_at, $doc.people.Count)
if ($Ymd) { Write-Host ("  絞り込み : {0} に実際に受付した人だけ" -f $Ymd) -ForegroundColor Yellow }
Write-Host ("  事業所の対応表 : {0} 社   オプションの対応表 : {1} 記号 (列には {2} を書く)" -f $cdOf.Count, $optOf.Count, $(if ($OptionAs -eq 'Code') { 'コード' } else { '名称' }))
Write-Host ''

$conn = Open-Db
try {
    # ---- 健診ナビに既に入っている人 ----
    # 予定日だけでなく実際に受付した日も見る (予定と違う日に来た人の予約も拾う)
    $dates = @($doc.people | ForEach-Object {
        [string]$_.planned.date
        if ($_.actual) { [string]$_.actual.checked_in_date }
    } | Where-Object { $_ } | Select-Object -Unique)
    $inList = ($dates | ForEach-Object { "'" + ($_ -replace '-', '/') + "'" }) -join ','
    # 団体名に「福生」が無い団体も、対応表の団体CDなら「もう入っている」の判定に含める
    $cdList = (($cdOf.Values | Select-Object -Unique | ForEach-Object { "'$_'" }) -join ',')
    if ($cdList -eq '') { $cdList = "''" }
    $exist = @((Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS KANJI, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ($inList)
  AND (d.MEISYO1 LIKE N'%$DantaiLike%' OR LTRIM(RTRIM(s.DANTAI_CD1)) IN ($cdList))
"@ @{}).Rows)
    $have = @{}
    foreach ($e in $exist) { $k = NoSpace $e.KANJI; if ($k -ne '') { $have[$k] = $e.YMD } }
    Write-Host ("  健診ナビに既に入っている福生の人 : {0} 人" -f $have.Count) -ForegroundColor DarkGray
    Write-Host ''

    # ========================================================================
    # 行を作る
    # ========================================================================
    $rows = @(); $skip = @{}
    function Skip($why, $who) {
        if (-not $script:skip.ContainsKey($why)) { $script:skip[$why] = @() }
        $script:skip[$why] += $who
    }
    foreach ($p in $doc.people) {
        $name = [string]$p.identity.name; if ($name -eq '') { $name = [string]$p.planned.name }
        $who  = "{0} {1} ({2})" -f $p.id, $name, $p.planned.company

        # 実際に受付した日でしぼる
        $ciDate = $(if ($p.actual) { [string]$p.actual.checked_in_date } else { '' })
        if ($Ymd) {
            $want = ($Ymd -replace '/', '-')
            if ($ciDate -ne $want) { continue }
        }
        if ($p.derived.is_cancelled) { Skip '受付アプリでキャンセル' $who; continue }
        if ($have.ContainsKey((NoSpace $name))) { Skip 'もう健診ナビに入っている' $who; continue }

        $kana = ToNarrowKana ([string]$p.identity.kana)    # 健診ナビは半角カナで持っているので合わせる
        if ($kana -eq '' -and -not $IncludeNoKana) { Skip 'カナが空 (取込で弾かれます)' $who; continue }

        $cd = $cdOf[(NoSpace $p.planned.company)]
        if (-not $cd) { Skip '事業所コードが引けない (対応表にない会社)' $who; continue }

        # 当日コースがあればそれを優先
        $cs = $(if ($p.actual -and [string]$p.actual.course -ne '') { [string]$p.actual.course } else { [string]$p.planned.course })
        if (-not $COURSE.ContainsKey($cs)) { Skip ("コースが [$cs] なので変換できない") $who; continue }

        $sex = ([string]$p.planned.gender).Trim()
        if ($sex -eq '') { Skip '性別が空 (取込で弾かれます)' $who; continue }

        # 予約日
        #   -Ymd で日を絞ったときは その日 (実際に受付した日)。
        #   絞っていないときは 予定日。テスト受付が残っていても予約日が狂わないようにする。
        $ymdOut = $(if ($Ymd) { $Ymd -replace '-', '/' } else { ([string]$p.planned.date) -replace '-', '/' })

        # オプション: 受付済みなら当日の内容 (actual.options が全体)。未受付なら予定。
        $received = ($p.actual -and $null -ne $p.actual.reception_number)
        $syms = @($(if ($received) { $p.actual.options } else { $p.planned.options }) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
        $optVals = @(); $optNote = @()
        if ($NoOptions -and $syms.Count -gt 0) { $optNote += ("オプション [{0}] は -NoOptions なので列に書かない → 健診ナビの画面で付ける" -f ($syms -join '.')) }
        foreach ($sym in $syms) {
            if ($NoOptions) { break }
            if (-not $optOf.ContainsKey($sym)) { $optNote += "記号 $sym は対応表にない"; continue }
            $o = $optOf[$sym]
            if ($o.CD -eq '') { $optNote += "$sym($($o.App)) は健診ナビにコードが無いので取込まない"; continue }
            $optVals += $(if ($OptionAs -eq 'Code') { $o.CD } else { $o.Name })
        }
        if ($optVals.Count -gt 10) { Skip 'オプションが10個を超える' $who; continue }

        # 確認一覧 (Excelには書けないので、健診ナビの画面で見てもらうこと)
        $memo = @()
        if ($received) {
            $a = $p.actual
            if ([string]$a.course -ne '' -and [string]$a.course -ne [string]$p.planned.course) { $memo += ("コース変更 予定[{0}] → 当日[{1}]" -f $p.planned.course, $a.course) }
            if ($a.exam_changes) {
                foreach ($pair in @(@('blood','採血'), @('stomach','胃部'), @('urine','尿'))) {
                    $v = [string]$a.exam_changes.($pair[0])
                    if ($v -ne '' -and $v -ne 'unchanged') { $memo += ("{0}={1}" -f $pair[1], $v) }
                }
            }
            if ($null -ne $a.stool_count -and [int]$a.stool_count -ne 0 -and [int]$a.stool_count -ne 2) { $memo += ("便 {0} 本" -f $a.stool_count) }
            if ($syms -contains 'M' -and $null -ne $a.stool_count -and [int]$a.stool_count -eq 0) { $memo += '大腸がん(M)ありだが便 0 本' }
            if ([string]$a.notes -ne '') { $memo += ("受付メモ: {0}" -f $a.notes) }
            if ($p.derived -and (@($p.derived.options_added) + @($p.derived.options_removed)).Count -gt 0) {
                $memo += ("オプション 追加[{0}] 中止[{1}]" -f (($p.derived.options_added) -join '.'), (($p.derived.options_removed) -join '.'))
            }
        }
        $memo += $optNote
        if (([string]$p.identity.birth_date) -eq '') { $memo += '生年月日が空' }

        $rows += [PSCustomObject]@{
            氏名 = $name; カナ = $kana; 性別 = $sex
            生年月日 = ([string]$p.identity.birth_date) -replace '-', '/'
            事業所CD = $cd; 事業所 = [string]$p.planned.company
            コースCD = $COURSE[$cs].CD; コース名 = $COURSE[$cs].Name
            予約日 = $ymdOut; 予約時間 = StartTime ([string]$p.planned.time)
            オプション = $optVals; オプション記号 = ($syms -join '.')
            受付番号 = $(if ($received) { [string]$p.actual.reception_number } else { '' })
            確認 = ($memo -join ' / ')
            ID = $p.id
        }
    }

    # ========================================================================
    # 一覧
    # ========================================================================
    Write-Host ('--- 出す人 : {0} 人 ---' -f $rows.Count) -ForegroundColor Cyan
    if ($rows.Count -gt 0) {
        foreach ($g in ($rows | Group-Object 予約日 | Sort-Object Name)) {
            $c = ($g.Group | Group-Object コースCD | ForEach-Object { "{0} {1}人" -f $_.Name, $_.Count }) -join ' / '
            Write-Host ("  {0}  {1} 人   {2}" -f $g.Name, $g.Count, $c)
        }
        $noBirth = @($rows | Where-Object { $_.生年月日 -eq '' }).Count
        if ($noBirth -gt 0) { Write-Host ("  ※ うち生年月日が空: {0} 人 (取り込めますが年齢判定が出ません)" -f $noBirth) -ForegroundColor Yellow }
        $withOpt = @($rows | Where-Object { $_.オプション.Count -gt 0 })
        Write-Host ("  オプションあり : {0} 人 / {1} 個" -f $withOpt.Count, (($withOpt | ForEach-Object { $_.オプション.Count }) | Measure-Object -Sum).Sum)
    }
    Write-Host ''
    if ($skip.Count -gt 0) {
        Write-Host '--- 出さない人 ---' -ForegroundColor Yellow
        foreach ($k in ($skip.Keys | Sort-Object)) {
            Write-Host ("  [{0}] {1} 人" -f $k, $skip[$k].Count) -ForegroundColor Yellow
            if ($k -notlike 'もう健診ナビ*') { $skip[$k] | ForEach-Object { Write-Host ('      ' + $_) -ForegroundColor DarkYellow } }
        }
        Write-Host ''
    }
    if ($rows.Count -eq 0) { Write-Host '出す人がいません。' -ForegroundColor Yellow; return }

    # ========================================================================
    # Excelを作る
    # ========================================================================
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $file  = Join-Path $Out ("予約取込_{0}{1}.xlsx" -f $(if ($Ymd) { ($Ymd -replace '[/-]','') + '_' } else { '' }), $stamp)
    $xl = $null
    try { $xl = New-Object -ComObject Excel.Application }
    catch { throw 'Excelが起動できません。このPCにExcelが入っているか確認してください。' }
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    try {
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = '予約取込'
        for ($c = 0; $c -lt $HEAD.Count; $c++) {
            $ws.Cells.Item(1, $c + 1).NumberFormat = '@'
            $ws.Cells.Item(1, $c + 1).Value2 = $HEAD[$c]
        }
        # 全部文字列で入れる (先頭の0や日付の自動変換を防ぐ)
        $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($rows.Count + 1, $HEAD.Count)).NumberFormat = '@'
        $r = 2
        foreach ($x in $rows) {
            $ws.Cells.Item($r, 1).Value2  = [string]($r - 1)      # No
            $ws.Cells.Item($r, 2).Value2  = $x.氏名
            $ws.Cells.Item($r, 3).Value2  = $x.カナ
            $ws.Cells.Item($r, 4).Value2  = $x.性別
            $ws.Cells.Item($r, 5).Value2  = $x.生年月日
            $ws.Cells.Item($r, 18).Value2 = $x.事業所CD
            $ws.Cells.Item($r, 19).Value2 = $x.事業所
            $ws.Cells.Item($r, 20).Value2 = $x.コースCD
            $ws.Cells.Item($r, 21).Value2 = $x.コース名
            $ws.Cells.Item($r, 22).Value2 = $x.予約日
            $ws.Cells.Item($r, 23).Value2 = $x.予約時間
            for ($k = 0; $k -lt $x.オプション.Count; $k++) { $ws.Cells.Item($r, 24 + $k).Value2 = [string]$x.オプション[$k] }
            $r++
        }
        [void]$ws.Columns.AutoFit()
        $wb.SaveAs($file, 51)
        $wb.Close($false)
    }
    finally {
        $xl.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    # 控え
    $csv = [System.IO.Path]::ChangeExtension($file, '.csv')
    $rows | Select-Object ID, 受付番号, 氏名, カナ, 性別, 生年月日, 事業所CD, 事業所, コースCD, 予約日, 予約時間,
                          @{ N='オプション'; E={ $_.オプション -join '.' } }, オプション記号, 確認 |
        Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

    # 確認一覧 (健診ナビの画面で見てもらうこと)
    $need = @($rows | Where-Object { $_.確認 -ne '' })
    $memoFile = [System.IO.Path]::ChangeExtension($file, '.確認一覧.txt')
    $lines = @("=== 確認一覧  $(Split-Path $file -Leaf)  ===", '',
               '健診ナビの予約取込では入れられないものです。取込のあと健診ナビの画面で対応してください。', '')
    if ($need.Count -eq 0) { $lines += '  (確認が必要な人はいません)' }
    foreach ($x in $need) { $lines += ("  受付{0,-4} {1} {2} ({3})" -f $x.受付番号, $x.ID, $x.氏名, $x.事業所); $lines += ("        {0}" -f $x.確認) }
    $lines | Out-File -FilePath $memoFile -Encoding Default

    Write-Host ("[出力] {0}  ({1} 人)" -f $file, $rows.Count) -ForegroundColor Green
    Write-Host ("[控え] {0}" -f $csv) -ForegroundColor DarkGray
    Write-Host ("[確認] {0}  ({1} 人)" -f $memoFile, $need.Count) -ForegroundColor $(if ($need.Count -gt 0) { 'Yellow' } else { 'DarkGray' })
    if ($need.Count -gt 0) {
        Write-Host ''
        Write-Host ('--- 確認一覧 ({0} 人。Excelには入りません。健診ナビの画面で対応) ---' -f $need.Count) -ForegroundColor Yellow
        foreach ($x in $need) { Write-Host ("  受付{0,-4} {1} {2}" -f $x.受付番号, $x.ID, $x.氏名) -ForegroundColor Yellow; Write-Host ("        {0}" -f $x.確認) -ForegroundColor DarkYellow }
    }
    Write-Host ''
    Write-Host '※ 健診ナビの予約取込にこのExcelをかけてください。' -ForegroundColor Yellow
    if ($NoOptions) { Write-Host '※ オプション検査の列は空です (-NoOptions)。オプションは確認一覧を見て健診ナビの画面で付けてください。' -ForegroundColor Yellow }
    else { Write-Host ("※ オプション検査の列には健診ナビの{0}を書いてあります。初回は1人だけで試してください。" -f $(if ($OptionAs -eq 'Code') { 'コード (OPJ014 など)' } else { '名称 (乳腺超音波検査 など)' })) -ForegroundColor Yellow }
    Write-Host '※ 取込が終わったら 受付結果取込.bat で受付番号を入れてください。' -ForegroundColor Yellow
}
finally {
    if ($conn -and $conn.State -eq 'Open') { $conn.Close() }
}
