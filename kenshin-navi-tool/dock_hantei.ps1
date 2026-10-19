<#
.SYNOPSIS
  人間ドック学会の判定を出して、健診ナビの判定と見比べる (dock_hantei.ps1)

.DESCRIPTION
  健診ナビのDBから値を読み、form\dock_kijun.csv (学会の判定区分 + 院内ルール) で判定する。
  判定の文字は学会の 4 段階: A 異常なし / B 軽度異常 / C 要再検査・生活改善 / D 要精密検査・治療。J (治療中) は使わない。
    ・項目 → くくり (脂質代謝など) → 総合 の順に、いちばん重い判定を上に上げる。
    ・院内ルールは基準表の「上限」行 (例: 視力は悪くても B) と「判定」行 (業者の読影をそのまま使う) で入れる。
    ・血糖・中性脂肪は 空腹時 → 随時 → 未判別 の順に値を探す。未判別 (食事の記録なし) は空腹時の基準で見る。
    ・血圧は 2 回目があれば 2 回目、無ければ 1 回目 (-BpFirst で 1 回目を使う)。
  健診ナビの判定は form\dock_moji.csv で学会の文字に読み替えて並べ、違うところに印を付ける。
  DBは読むだけで、一切変更しない。結果は out\ に HTML (ブラウザで開く) と CSV で出す。氏名は出さない。
  例:
    dock_hantei.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
#>
[CmdletBinding()]
param(
    [string]$Dates,
    [int]$UkeFrom = 0,
    [int]$UkeTo = 0,
    [switch]$BpFirst,
    [switch]$NoOpen,
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

$FormDir = Join-Path $PSScriptRoot 'form'
$OutDir  = Join-Path $PSScriptRoot 'out'
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
$Inv = [Globalization.CultureInfo]::InvariantCulture

$RANK  = @{ 'A' = 1; 'B' = 2; 'C' = 3; 'D' = 4 }
$LABEL = [ordered]@{ 'A' = '異常なし'; 'B' = '軽度異常'; 'C' = '要再検査・生活改善'; 'D' = '要精密検査・治療' }
# くくりの名前 → 健診ナビのくくり判定の項目CD (見比べ用)。並び順は画面の列の順
$KUKURI = [ordered]@{
    '身体計測' = '600020'; '視力' = '600040'; '聴力' = '600070'; '血圧' = '600150'; '胸部X線' = '600090'; '心電図' = '600130'
    '胃部' = '600160'; '眼底' = '600060'; '肝機能' = '600331'; '脂質代謝' = '600320'; '糖代謝' = '600340'; '腎機能' = '600220'
    '尿酸' = '600460'; '血液一般' = '600310'; '便潜血' = '600200'
}
$TOTAL_CD = '11000'   # 総合判定(編集)。無ければ 10000 (自動判定の総合) を見る
$TOTAL_AUTO = '10000'
# ツールが作る項目の元になる項目CD
$SRC_CDS = @('001211','001212','001253','001254','001255','001256','067012','067013','067018','067019',
             '068705','068708','H30-0011','H30-0012','068682','068702','R06-0003','R06-0002')

function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
function H([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }
function To-Num([string]$s) {
    if ($null -eq $s) { return $null }
    $t = [regex]::Replace($s.Trim(), '[０-９．－]', [System.Text.RegularExpressions.MatchEvaluator]{ param($m) [string][char]([int][char]$m.Value[0] - 0xFEE0) })
    $mt = [regex]::Match($t, '-?\d+(\.\d+)?')
    if (-not $mt.Success) { return $null }
    return [double]::Parse($mt.Value, $Inv)
}
function Read-Form([string]$name) {
    $path = Join-Path $FormDir $name
    if (-not (Test-Path $path)) { throw "基準表がありません: $path" }
    return @(Get-Content $path -Encoding UTF8 | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } | ConvertFrom-Csv)
}
function Worse([string]$a, [string]$b) {
    if (-not $RANK.ContainsKey($a)) { return $b }
    if (-not $RANK.ContainsKey($b)) { return $a }
    if ($RANK[$b] -gt $RANK[$a]) { return $b } else { return $a }
}

# ---- 基準表を読む ----
$Rules = Read-Form 'dock_kijun.csv'
$NAVI = @{}
foreach ($m in (Read-Form 'dock_moji.csv')) { $NAVI[([string]$m.'ナビ').Trim()] = ([string]$m.'学会').Trim() }
$bad = @()
$n = 0
foreach ($r in $Rules) {
    $n++
    foreach ($c in @('種類','項目','項目名','くくり','性別','以上','未満','文字','判定','出典','メモ')) {
        if ($null -eq $r.$c) { $r | Add-Member -NotePropertyName $c -NotePropertyValue '' -Force }
        $r.$c = ([string]$r.$c).Trim()
    }
    $r | Add-Member -NotePropertyName Lo -NotePropertyValue (To-Num $r.'以上') -Force
    $r | Add-Member -NotePropertyName Hi -NotePropertyValue (To-Num $r.'未満') -Force
    if ($r.'種類' -notin @('数値','文字','判定','上限')) { $bad += "$n 行目: 種類「$($r.'種類')」が分かりません" }
    if ($r.'種類' -ne '判定' -and -not $RANK.ContainsKey($r.'判定')) { $bad += "$n 行目 ($($r.'項目名')): 判定は A〜D で書いてください" }
    if ($r.'性別' -notin @('','男','女')) { $bad += "$n 行目 ($($r.'項目名')): 性別は 空欄・男・女" }
    if (-not $KUKURI.Contains($r.'くくり')) { $bad += "$n 行目 ($($r.'項目名')): くくり「$($r.'くくり')」は一覧にありません" }
    if ($r.'種類' -eq '文字') { try { [void][regex]::new($r.'文字') } catch { $bad += "$n 行目 ($($r.'項目名')): 文字の書き方が正しくありません" } }
}
if ($bad.Count -gt 0) { $bad | ForEach-Object { Write-Host " $_" -ForegroundColor Red }; throw '基準表 dock_kijun.csv を直してください' }
# 項目の並び (基準表の順) と、項目 → くくり・名前
$KEYS = [ordered]@{}
foreach ($r in $Rules) { if (-not $KEYS.Contains($r.'項目')) { $KEYS[$r.'項目'] = @{ Kukuri = $r.'くくり'; Name = $r.'項目名' } } }

# ---- 対象 ----
if (-not $Dates) { $Dates = Read-Host '受診日 (例 2026/09/03。いくつかなら 2026/09/03,2026/09/09)' }
$DL = @()
foreach ($d in @(($Dates -split '[,、\s]+') | Where-Object { $_ -ne '' })) {
    $t = $d.Trim() -replace '-', '/'
    if ($t -notmatch '^\d{4}/\d{1,2}/\d{1,2}$') { throw "受診日の形が違います: $d (例 2026/09/03)" }
    $DL += [datetime]::ParseExact($t, 'yyyy/M/d', $Inv).ToString('yyyy/MM/dd')
}
if ($DL.Count -eq 0) { throw '受診日を入れてください' }
$WHERE = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('" + ($DL -join "','") + "')"
if ($UkeFrom -gt 0) { $WHERE += " AND s.UKE_NO_KENSA >= $UkeFrom" }
if ($UkeTo -gt 0)   { $WHERE += " AND s.UKE_NO_KENSA <= $UkeTo" }

$cds = @($KEYS.Keys) + $SRC_CDS + @($KUKURI.Values) + @($TOTAL_CD, $TOTAL_AUTO)
$cds = @($cds | Where-Object { $_ -match '^[0-9A-Z-]+$' } | Select-Object -Unique)

Write-Host ''
Write-Host ' 人間ドック学会の判定を出して、健診ナビの判定と見比べる (DBは読むだけ)' -ForegroundColor Cyan
Write-Host (" 受診日 {0}{1}" -f ($DL -join '・'), $(if ($UkeFrom -or $UkeTo) { "  受付番号 $UkeFrom〜$UkeTo" } else { '' }))
$conn = Open-Db
try {
    $people = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ AS PK, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(s.UKE_NO_KENSA)) AS UKE, LTRIM(RTRIM(ISNULL(j.SEIBETU,''))) AS SEX
FROM T_KENSIN s JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
WHERE $WHERE
"@).Rows
    $items = (Invoke-DbQuery $conn @"
SELECT k.PK_SEQ AS PK, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS HK
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE $WHERE AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('$($cds -join "','")')
"@).Rows
} finally { $conn.Close() }
if (@($people).Count -eq 0) { Write-Host ' 該当する受診者がいません。' -ForegroundColor Yellow; return }
Write-Host (" {0} 人" -f @($people).Count)

$DB = @{}
foreach ($it in $items) {
    $pk = [string]$it.PK
    if (-not $DB.ContainsKey($pk)) { $DB[$pk] = @{} }
    $DB[$pk][[string]$it.CD] = @{ V = [string]$it.V; HK = [string]$it.HK }
}

# ---- 1 人分の値 (ツールが作る項目も含む) ----
function Get-Values($d) {
    $vals = @{}
    function Num($cd) { if ($d.ContainsKey($cd)) { return (To-Num $d[$cd].V) } return $null }
    foreach ($cd in $d.Keys) { $vals[$cd] = @{ V = $d[$cd].V; N = (To-Num $d[$cd].V); HK = $d[$cd].HK } }
    $ht = Num '001211'; $wt = Num '001212'
    if ($null -ne $ht -and $null -ne $wt -and $ht -gt 0) {
        $bmi = [math]::Round($wt / [math]::Pow($ht / 100, 2), 1)
        $vals['BMI'] = @{ V = $bmi.ToString('0.0', $Inv); N = $bmi; HK = '' }
    }
    $s2 = Num '001255'; $d2 = Num '001256'
    if (-not $BpFirst -and $null -ne $s2 -and $null -ne $d2) { $sbp = $s2; $dbp = $d2; $bpNote = '2回目' } else { $sbp = Num '001253'; $dbp = Num '001254'; $bpNote = '1回目' }
    if ($null -ne $sbp) { $vals['収縮期'] = @{ V = "$sbp ($bpNote)"; N = $sbp; HK = '' } }
    if ($null -ne $dbp) { $vals['拡張期'] = @{ V = "$dbp ($bpNote)"; N = $dbp; HK = '' } }
    $eyes = @()
    foreach ($pair in @(@('067018','067012','右'), @('067019','067013','左'))) {
        $v = Num $pair[0]; $kind = '矯正'
        if ($null -eq $v) { $v = Num $pair[1]; $kind = '裸眼' }
        if ($null -ne $v) { $eyes += [PSCustomObject]@{ N = $v; T = "$($pair[2]) $v ($kind)" } }
    }
    if ($eyes.Count -gt 0) {
        $worst = @($eyes | Sort-Object N)[0]
        $vals['視力'] = @{ V = (($eyes | ForEach-Object { $_.T }) -join ' / '); N = $worst.N; HK = '' }
    }
    $tc = Num '068705'; $hdl = Num '068708'
    if ($null -ne $tc -and $null -ne $hdl) { $nh = $tc - $hdl; $vals['nonHDL'] = @{ V = "$nh"; N = $nh; HK = '' } }
    foreach ($g in @(@('H30-0011','空腹時血糖','空腹時'), @('H30-0012','随時血糖','随時'), @('068682','空腹時血糖','食事の記録なし → 空腹時で見る')),
                    @(@('068702','空腹時中性脂肪','空腹時'), @('R06-0003','随時中性脂肪','随時'), @('R06-0002','空腹時中性脂肪','食事の記録なし → 空腹時で見る'))) {
        foreach ($c in $g) {
            $v = Num $c[0]
            if ($null -ne $v) { $vals[$c[1]] = @{ V = "$v ($($c[2]))"; N = $v; HK = '' }; break }
        }
    }
    return $vals
}

# ---- 1 項目の判定 ----
function Judge($key, [string]$sex, $info) {
    $res = ''; $note = ''
    foreach ($r in $Rules) {
        if ($r.'項目' -ne $key -or $r.'種類' -eq '上限') { continue }
        if ($r.'性別' -ne '' -and $r.'性別' -ne $sex) { continue }
        switch ($r.'種類') {
            '数値' { if ($null -ne $info.N -and ($null -eq $r.Lo -or $info.N -ge $r.Lo) -and ($null -eq $r.Hi -or $info.N -lt $r.Hi)) { $res = $r.'判定' } }
            '文字' { if ($info.V -ne '' -and $info.V -match $r.'文字') { $res = $r.'判定' } }
            '判定' { if ($info.HK -ne '') { if ($NAVI.ContainsKey($info.HK)) { $res = $NAVI[$info.HK]; $note = "健診ナビ $($info.HK) を読み替え" } else { $note = "判定「$($info.HK)」は dock_moji.csv に無い" } } }
        }
        if ($res -ne '') { break }
    }
    if ($res -eq '' -and $note -eq '') { $note = '基準なし' }
    foreach ($r in $Rules) {
        if ($r.'項目' -ne $key -or $r.'種類' -ne '上限' -or $res -eq '') { continue }
        if ($r.'性別' -ne '' -and $r.'性別' -ne $sex) { continue }
        if ($RANK[$res] -gt $RANK[$r.'判定']) { $note = "院内ルールで $res → $($r.'判定')"; $res = $r.'判定' }
    }
    return @{ L = $res; Note = $note }
}
function Navi-To([string]$hk) {
    if ($hk -eq '') { return '' }
    if ($hk -eq 'J') { return 'J' }
    if ($NAVI.ContainsKey($hk)) { return $NAVI[$hk] }
    return '?'
}

# ---- 全員 ----
$Results = @()
$Detail = @()
$NoRule = @{}
foreach ($pp in $people) {
    $pk = [string]$pp.PK
    $sex = switch ([string]$pp.SEX) { '1' { '男' } '2' { '女' } default { '' } }
    $d = $(if ($DB.ContainsKey($pk)) { $DB[$pk] } else { @{} })
    $vals = Get-Values $d
    $kk = [ordered]@{}
    foreach ($name in $KUKURI.Keys) { $kk[$name] = @{ L = ''; Items = @() } }
    foreach ($key in $KEYS.Keys) {
        if (-not $vals.ContainsKey($key)) { continue }
        $info = $vals[$key]
        if ($info.V -eq '' -and $info.HK -eq '') { continue }
        $jr = Judge $key $sex $info
        $kname = $KEYS[$key].Kukuri
        $kk[$kname].L = Worse $kk[$kname].L $jr.L
        $shown = $(if ($info.V -ne '') { $info.V } else { '(判定のみ)' })
        $kk[$kname].Items += ("{0} {1} → {2}{3}" -f $KEYS[$key].Name, $shown, $(if ($jr.L) { $jr.L } else { '-' }), $(if ($jr.Note) { " ($($jr.Note))" } else { '' }))
        if ($jr.Note -eq '基準なし' -or $jr.Note -like '*dock_moji*') { $NoRule["$($KEYS[$key].Name)|$shown"] = 1 + [int]$NoRule["$($KEYS[$key].Name)|$shown"] }
        $naviItem = $(if ($key -match '^[0-9A-Z-]+$' -and $d.ContainsKey($key)) { $d[$key].HK } else { '' })
        $Detail += [PSCustomObject][ordered]@{
            受診日 = $pp.YMD; 受付番号 = $pp.UKE; 性別 = $sex; くくり = $kname; 項目 = $KEYS[$key].Name; 値 = $shown
            学会判定 = $jr.L; メモ = $jr.Note; ナビ判定 = $naviItem; ナビ換算 = (Navi-To $naviItem)
        }
    }
    $total = ''
    $o = [ordered]@{ 受診日 = $pp.YMD; 受付番号 = $pp.UKE; 性別 = $sex; PK = $pk }
    foreach ($name in $KUKURI.Keys) {
        $total = Worse $total $kk[$name].L
        $nk = $(if ($d.ContainsKey($KUKURI[$name])) { $d[$KUKURI[$name]].HK } else { '' })
        $o[$name] = @{ L = $kk[$name].L; Navi = $nk; NaviTo = (Navi-To $nk); Items = $kk[$name].Items }
    }
    $nt = ''
    foreach ($tc in @($TOTAL_CD, $TOTAL_AUTO)) {
        if ($nt -eq '' -and $d.ContainsKey($tc)) {
            $nt = $d[$tc].HK
            if ($nt -eq '' -and $d[$tc].V -match '^\s*([A-J])') { $nt = $Matches[1] }
        }
    }
    $o['総合'] = @{ L = $total; Navi = $nt; NaviTo = (Navi-To $nt); Items = @() }
    $Results += [PSCustomObject]$o
}
$Results = @($Results | Sort-Object 受診日, @{ Expression = { $u = 0; if ([int]::TryParse([string]$_.受付番号, [ref]$u)) { $u } else { [int]::MaxValue } } })

function Cmp($c) {
    if ($c.L -eq '' -or $c.Navi -eq '') { return '' }
    if ($c.NaviTo -eq 'J') { return 'J' }
    if ($c.NaviTo -eq $c.L) { return '一致' }
    if (-not $RANK.ContainsKey($c.NaviTo)) { return '?' }
    if ($RANK[$c.L] -gt $RANK[$c.NaviTo]) { return '学会が重い' } else { return '学会が軽い' }
}

# ---- 見比べの集計 ----
$COLS = @($KUKURI.Keys) + @('総合')
$Sum = foreach ($name in $COLS) {
    $cnt = @{ '一致' = 0; '学会が重い' = 0; '学会が軽い' = 0; 'J' = 0; '?' = 0; '' = 0 }
    $dist = @{ 'A' = 0; 'B' = 0; 'C' = 0; 'D' = 0 }
    foreach ($r in $Results) { $cnt[(Cmp $r.$name)]++; if ($RANK.ContainsKey($r.$name.L)) { $dist[$r.$name.L]++ } }
    [PSCustomObject][ordered]@{ くくり = $name; A = $dist.A; B = $dist.B; C = $dist.C; D = $dist.D
        一致 = $cnt['一致']; 学会が重い = $cnt['学会が重い']; 学会が軽い = $cnt['学会が軽い']; ナビ治療中 = $cnt['J'] }
}
Line
$Sum | Format-Table -AutoSize | Out-String -Width 200 | Write-Host

# ---- 出力 ----
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath  = Join-Path $OutDir "dock_hantei_$ts.csv"
$htmlPath = Join-Path $OutDir "dock_hantei_$ts.html"
$Detail | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8

$sbH = New-Object System.Text.StringBuilder
function A([string]$s) { [void]$sbH.Append($s) }
A @"
<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>学会判定の見比べ</title>
<style>
body{font-family:'Meiryo','Yu Gothic',sans-serif;font-size:13px;margin:16px;color:#222;background:#fff}
h1{font-size:18px} h2{font-size:15px;margin-top:24px;border-left:4px solid #369;padding-left:6px}
table{border-collapse:collapse} th,td{border:1px solid #bbb;padding:2px 6px;text-align:center;white-space:nowrap}
th{background:#eef2f7;position:sticky;top:0} td.l{text-align:left}
.A{color:#1a7f37}.B{color:#555}.C{color:#b35900;font-weight:bold}.D{color:#c00;font-weight:bold}
td.heavy{background:#ffe0e0} td.light{background:#e0ecff} td.j{background:#f2f2f2}
small{color:#777} .note{color:#555;margin:4px 0} .warn{color:#c00}
</style></head><body>
<h1>人間ドック学会の判定 と 健診ナビの判定 の見比べ</h1>
<p class="note">受診日 $(H ($DL -join '・'))$(if ($UkeFrom -or $UkeTo) { H "　受付番号 $UkeFrom〜$UkeTo" })　／　$(@($Results).Count) 人　／　作成 $(Get-Date -Format 'yyyy/MM/dd HH:mm')　／　DBは読むだけ</p>
<p class="note">判定: A 異常なし ／ B 軽度異常 ／ C 要再検査・生活改善 ／ D 要精密検査・治療。　各マスの上が学会の判定、下の小さい字が健診ナビの判定 (→ 学会の文字に読み替え)。</p>
<p class="note"><span style="background:#ffe0e0">赤</span> = 学会の方が重い　<span style="background:#e0ecff">青</span> = 学会の方が軽い　<span style="background:#f2f2f2">灰</span> = 健診ナビが J (治療中)。マスにマウスを乗せると項目ごとの値と判定が出る。</p>
<p class="warn">基準表 form\dock_kijun.csv の「★要確認」の行は、学会の判定区分 (PDF) と照らしてから使ってください。</p>
<h2>1. くくりごとの人数と見比べ</h2><table><tr><th>くくり</th><th>A</th><th>B</th><th>C</th><th>D</th><th>一致</th><th>学会が重い</th><th>学会が軽い</th><th>ナビ治療中</th></tr>
"@
foreach ($s in $Sum) {
    A ("<tr><td class=l>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td></tr>" -f (H $s.'くくり'), $s.A, $s.B, $s.C, $s.D, $s.'一致', $s.'学会が重い', $s.'学会が軽い', $s.'ナビ治療中')
}
A '</table><h2>2. 1 人ずつ</h2><table><tr><th>受診日</th><th>受付番号</th><th>性別</th><th>総合</th>'
foreach ($name in $KUKURI.Keys) { A "<th>$(H $name)</th>" }
A '</tr>'
foreach ($r in $Results) {
    A ("<tr><td>{0}</td><td>{1}</td><td>{2}</td>" -f (H $r.'受診日'), (H $r.'受付番号'), (H $r.'性別'))
    foreach ($name in @('総合') + @($KUKURI.Keys)) {
        $c = $r.$name
        $cls = switch (Cmp $c) { '学会が重い' { 'heavy' } '学会が軽い' { 'light' } 'J' { 'j' } default { '' } }
        $tip = (@($c.Items) -join "`n")
        if ($c.Navi -ne '') { $tip += "`n健診ナビ: $($c.Navi)" }
        $naviTxt = $(if ($c.Navi -ne '') { "<br><small>ナビ $(H $c.Navi)$(if ($c.NaviTo -ne '' -and $c.NaviTo -ne $c.Navi -and $c.NaviTo -ne 'J') { '→' + (H $c.NaviTo) })</small>" } else { '' })
        A ("<td class='{0}' title='{1}'><span class='{2}'>{3}</span>{4}</td>" -f $cls, (H $tip.Trim()), (H $c.L), $(if ($c.L) { H $c.L } else { '<small>-</small>' }), $naviTxt)
    }
    A '</tr>'
}
A '</table>'
if ($NoRule.Count -gt 0) {
    A '<h2>3. 基準表に当てはまらなかった値 (判定を付けていない)</h2><table><tr><th>項目</th><th>値</th><th>人数</th></tr>'
    foreach ($k in ($NoRule.Keys | Sort-Object)) { $kv = $k -split '\|', 2; A ("<tr><td class=l>{0}</td><td class=l>{1}</td><td>{2}</td></tr>" -f (H $kv[0]), (H $kv[1]), $NoRule[$k]) }
    A '</table><p class="note">基準表に行を足すか、値の入れ方を確かめてください。</p>'
}
A '<h2>4. 使った基準表 (form\dock_kijun.csv)</h2><table><tr><th>種類</th><th>項目</th><th>くくり</th><th>性別</th><th>以上</th><th>未満</th><th>文字</th><th>判定</th><th>出典</th><th>メモ</th></tr>'
foreach ($r in $Rules) {
    A ("<tr><td>{0}</td><td class=l>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td class=l>{6}</td><td><span class='{7}'>{7}</span></td><td>{8}</td><td class=l>{9}</td></tr>" -f (H $r.'種類'), (H $r.'項目名'), (H $r.'くくり'), (H $r.'性別'), (H $r.'以上'), (H $r.'未満'), (H $r.'文字'), (H $r.'判定'), (H $r.'出典'), (H $r.'メモ'))
}
A '</table></body></html>'
[System.IO.File]::WriteAllText($htmlPath, $sbH.ToString(), (New-Object System.Text.UTF8Encoding $true))

Write-Host " 書き出しました (氏名なし):" -ForegroundColor Green
Write-Host "   $htmlPath"
Write-Host "   $csvPath  (項目ごとの値と判定)"
if ($NoRule.Count -gt 0) { Write-Host (" 基準表に当てはまらなかった値が {0} 種類あります (HTML の 3)。" -f $NoRule.Count) -ForegroundColor Yellow }
if (-not $NoOpen) { Start-Process $htmlPath }
