<#
.SYNOPSIS
  人間ドック学会の文字で判定を出して、健診ナビの判定と見比べる (dock_hantei.ps1)

.DESCRIPTION
  健診ナビのDBから値を読み、次の表で判定する。どの団体・コースでも使える。
    form\dock_kijun.csv       院内の決まり (手で直す)。ここに行がある項目は、こちらだけで判定する。上限・除外もここ
    form\dock_kijun_navi.csv  健診ナビの基準値の設定から作った表 (dock_setup.bat で作る。手で直さない)
    form\dock_kukuri.csv      項目 → くくり (脂質代謝など) の表 (手で決めたもの。無い項目は dock_kukuri_navi.csv = 健診ナビの実際の判定から作った表)
    form\dock_moji.csv        健診ナビの判定記号 (8段階) → 学会の文字 (4段階)
  判定の文字は A 異常なし / B 軽度異常 / C 要再検査・生活改善 / D 要精密検査・治療。J (治療中) は使わない。
  項目 → くくり → 総合 の順に、いちばん重い判定を上に上げる。
  健診ナビに入っている判定も学会の文字に読み替えて並べ、違うところに色を付ける (違いがゼロなら、記号が違うだけで結果は同じ)。
  DBは読むだけで、一切変更しない。結果は out\ に HTML (ブラウザで開く) と CSV で出す。氏名は出さない。
  例:
    dock_hantei.ps1 -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
    dock_hantei.ps1 -Dates 2026/08/21 -Course YSA1
#>
[CmdletBinding()]
param(
    [string]$Dates,
    [int]$UkeFrom = 0,
    [int]$UkeTo = 0,
    [string]$Course,
    [switch]$NoOpen,
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'
# 止まったときに何行目かを出す (原因を調べるため)
trap { Write-Host (" エラー: {0}`n   {1} 行目: {2}" -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber, ([string]$_.InvocationInfo.Line).Trim()) -ForegroundColor Red; break }

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
$RANK = @{ 'A' = 1; 'B' = 2; 'C' = 3; 'D' = 4 }
$OTHER = 'その他'
$TOTAL_CDS = @('11000', '10000')   # 総合判定(編集)、無ければ自動判定の総合

function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
# H は PowerShell の Get-History の短縮名なので関数名に使わない
function HtmlEnc([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }
function To-Num([string]$s) {
    if ($null -eq $s) { return $null }
    $t = $s.Trim()
    if ($t -match '[０-９．－]') { $t = $t.Normalize([System.Text.NormalizationForm]::FormKC) }   # 全角の数字・点・マイナスを半角に
    # 数字だけの結果のときだけ数値にする (「0.1未満」「凝固」などは文字の行で見る)
    if ($t -notmatch '^-?\d+(\.\d+)?$') { return $null }
    return [double]::Parse($t, $Inv)
}
function Read-Form([string]$name) {
    $path = Join-Path $FormDir $name
    if (-not (Test-Path $path)) { throw "表がありません: $path" }
    return @(Get-Content $path -Encoding UTF8 | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } | ConvertFrom-Csv)
}
function Worse([string]$a, [string]$b) {
    if (-not $RANK.ContainsKey($a)) { return $b }
    if (-not $RANK.ContainsKey($b)) { return $a }
    if ($RANK[$b] -gt $RANK[$a]) { return $b } else { return $a }
}

# ---- 表を読む ----
$NAVI = @{}
foreach ($m in (Read-Form 'dock_moji.csv')) { $NAVI[([string]$m.'ナビ').Trim()] = ([string]$m.'学会').Trim() }
$KUKURI = [ordered]@{}   # くくり名 → くくりCD
$ItemK = @{}             # 項目CD → @{ Kukuri; Name }
foreach ($k in (Read-Form 'dock_kukuri.csv')) {
    $cd = ([string]$k.'項目').Trim(); $kn = ([string]$k.'くくり').Trim(); $kc = ([string]$k.'くくりCD').Trim()
    if ($cd -eq '' -or $kn -eq '') { continue }
    $ItemK[$cd] = @{ Kukuri = $kn; Name = ([string]$k.'項目名').Trim() }
    if (-not $KUKURI.Contains($kn)) { $KUKURI[$kn] = $kc } elseif ($KUKURI[$kn] -eq '' -and $kc -ne '') { $KUKURI[$kn] = $kc }
}
# 手の表に無い項目は、dock_setup が健診ナビの実際の判定から作った表 (dock_kukuri_navi.csv) で補う
$CdName = @{}
foreach ($kn in $KUKURI.Keys) { if ($KUKURI[$kn] -ne '') { $CdName[$KUKURI[$kn]] = $kn } }
$kukAuto = Join-Path $FormDir 'dock_kukuri_navi.csv'
if (Test-Path $kukAuto) {
    foreach ($k in (Read-Form 'dock_kukuri_navi.csv')) {
        $cd = ([string]$k.'項目').Trim(); $kc = ([string]$k.'くくりCD').Trim()
        if ($cd -eq '' -or $kc -eq '' -or $ItemK.ContainsKey($cd)) { continue }
        $kn = $(if ($CdName.ContainsKey($kc)) { $CdName[$kc] } else { ([string]$k.'くくり').Trim() })
        if (-not $KUKURI.Contains($kn)) { $KUKURI[$kn] = $kc; $CdName[$kc] = $kn }
        $ItemK[$cd] = @{ Kukuri = $kn; Name = ([string]$k.'項目名').Trim() }
    }
}
$navPath = Join-Path $FormDir 'dock_kijun_navi.csv'
if (-not (Test-Path $navPath)) { throw '健診ナビの設定から作る表 form\dock_kijun_navi.csv がありません。先に dock_setup.bat を実行してください。' }

$COLS_RULE = @('種類','項目','項目名','基準CD','性別','年齢以上','年齢以下','以上','以下','未満','文字','判定','ナビ判定','出典','メモ')
$bad = @()
function Prep-Rules($list, [string]$file, [bool]$fromNavi) {
    # 速さのため、表の 1 行を hashtable にする (Add-Member は遅い)。文字の行は正規表現を先に作っておく
    $res = New-Object System.Collections.Generic.List[object]
    $n = 0
    foreach ($row0 in $list) {
        $n++
        $r = @{}
        foreach ($c in $COLS_RULE) { $v = $row0.$c; $r[$c] = $(if ($null -eq $v) { '' } else { ([string]$v).Trim() }) }
        if ($fromNavi) { $r['判定'] = $(if ($NAVI.ContainsKey($r['ナビ判定'])) { $NAVI[$r['ナビ判定']] } else { '' }) }   # dock_moji を直したらすぐ効くように
        $r.Lo = To-Num $r['以上']; $r.Le = To-Num $r['以下']; $r.Hi = To-Num $r['未満']; $r.AgeF = To-Num $r['年齢以上']; $r.AgeT = To-Num $r['年齢以下']
        $tag = "$file $n 行目 ($($r['項目']) $($r['項目名']))"
        if ($r['種類'] -notin @('数値','文字','判定','上限','除外')) { $script:bad += "${tag}: 種類「$($r['種類'])」が分かりません" }
        if ($r['種類'] -in @('数値','文字','上限') -and -not $RANK.ContainsKey($r['判定'])) { $script:bad += "${tag}: 判定は A〜D (健診ナビの表なら dock_moji.csv に「$($r['ナビ判定'])」が無い)" }
        if ($r['性別'] -notin @('','男','女')) { $script:bad += "${tag}: 性別は 空欄・男・女" }
        if ($r['種類'] -eq '文字') { try { $r.Rx = [regex]::new($r['文字']) } catch { $script:bad += "${tag}: 文字の書き方が正しくありません" } }
        $res.Add($r)
    }
    return ,$res
}
$Inai  = Prep-Rules (Read-Form 'dock_kijun.csv') 'dock_kijun.csv' $false
$NavR  = Prep-Rules (Read-Form 'dock_kijun_navi.csv') 'dock_kijun_navi.csv' $true
if ($bad.Count -gt 0) { $bad | Select-Object -First 30 | ForEach-Object { Write-Host " $_" -ForegroundColor Red }; throw '表を直してください (dock_kijun_navi.csv なら dock_setup.bat で作り直し)' }

$InaiBy = @{}; $CapBy = @{}; $NavBy = @{}; $Excl = @{}
foreach ($r in $Inai) {
    switch ($r.'種類') {
        '上限' { if (-not $CapBy.ContainsKey($r.'項目')) { $CapBy[$r.'項目'] = @() }; $CapBy[$r.'項目'] += $r }
        '除外' { $Excl[$r.'項目'] = $true }
        default { if (-not $InaiBy.ContainsKey($r.'項目')) { $InaiBy[$r.'項目'] = @() }; $InaiBy[$r.'項目'] += $r }
    }
}
foreach ($r in $NavR) { if (-not $NavBy.ContainsKey($r.'項目')) { $NavBy[$r.'項目'] = @() }; $NavBy[$r.'項目'] += $r }

# ---- 対象 ----
$ask = -not $Dates
if ($ask) { $Dates = Read-Host '受診日 (例 2026/09/03。いくつかなら 2026/09/03,2026/09/09)' }
$DL = @()
foreach ($d in @(($Dates -split '[,、\s]+') | Where-Object { $_ -ne '' })) {
    $t = $d.Trim() -replace '-', '/'
    if ($t -notmatch '^\d{4}/\d{1,2}/\d{1,2}$') { throw "受診日の形が違います: $d (例 2026/09/03)" }
    $DL += [datetime]::ParseExact($t, 'yyyy/M/d', $Inv).ToString('yyyy/MM/dd')
}
if ($DL.Count -eq 0) { throw '受診日を入れてください' }
if ($ask) {
    $u = Read-Host '受付番号の範囲 (例 4631-4831。全員なら空のまま Enter)'
    if ($u -match '^\s*(\d+)\s*[-〜~]\s*(\d+)\s*$') { $UkeFrom = [int]$Matches[1]; $UkeTo = [int]$Matches[2] }
}
$WHERE = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('" + ($DL -join "','") + "')"
if ($UkeFrom -gt 0) { $WHERE += " AND s.UKE_NO_KENSA >= $UkeFrom" }
if ($UkeTo -gt 0)   { $WHERE += " AND s.UKE_NO_KENSA <= $UkeTo" }
$CL = @(($Course -split '[,、\s]+') | Where-Object { $_ -match '^[0-9A-Za-z_-]+$' })
if ($CL.Count -gt 0) { $WHERE += " AND LTRIM(RTRIM(s.COURSE_CD)) IN ('" + ($CL -join "','") + "')" }

Write-Host ''
Write-Host ' 学会の文字で判定を出して、健診ナビの判定と見比べる (DBは読むだけ)' -ForegroundColor Cyan
$cond = "受診日 $($DL -join '・')" + $(if ($UkeFrom -or $UkeTo) { "  受付番号 $UkeFrom〜$UkeTo" } else { '' }) + $(if ($CL.Count) { "  コース $($CL -join ',')" } else { '' })
Write-Host " $cond"
$conn = Open-Db
try {
    $people = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ AS PK, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(s.UKE_NO_KENSA)) AS UKE,
       LTRIM(RTRIM(ISNULL(j.SEIBETU,''))) AS SEX, j.D_SEINEN AS BIRTH, s.D_KENSIN AS KDATE,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS COURSE, LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), c1.KIJUN_CD),''))) AS KIJUN
FROM T_KENSIN s
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_COURSE1 c1 ON c1.COURSE_CD = s.COURSE_CD AND c1.DANTAI_CD1 = s.DANTAI_CD1
WHERE $WHERE
"@).Rows
    $items = (Invoke-DbQuery $conn @"
SELECT k.PK_SEQ AS PK, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V,
       LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS VC, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS HK
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE $WHERE
"@).Rows
} finally { $conn.Close() }
if (@($people).Count -eq 0) { Write-Host ' 該当する受診者がいません。' -ForegroundColor Yellow; return }
Write-Host (" {0} 人" -f @($people).Count)

$DB = @{}
foreach ($it in $items) {
    $pk = [string]$it.PK
    if (-not $DB.ContainsKey($pk)) { $DB[$pk] = @{} }
    $DB[$pk][[string]$it.CD] = @{ V = [string]$it.V; VC = [string]$it.VC; HK = [string]$it.HK; N = (To-Num ([string]$it.V)) }
}

# ---- 判定 ----
function Cond-Ok($r, $who) {
    if ($r.'基準CD' -ne '' -and $who.Kijun -ne '' -and $r.'基準CD' -ne $who.Kijun) { return $false }
    if ($r.'性別' -ne '' -and $r.'性別' -ne $who.Sex) { return $false }
    if ($null -ne $who.Age) {
        if ($null -ne $r.AgeF -and $who.Age -lt $r.AgeF) { return $false }
        if ($null -ne $r.AgeT -and $who.Age -gt $r.AgeT) { return $false }
    }
    return $true
}
function Eval-Rule($r, $info) {
    switch ($r.'種類') {
        '数値' {
            if ($null -eq $info.N) { return '' }
            if ($null -ne $r.Lo -and $info.N -lt $r.Lo) { return '' }
            if ($null -ne $r.Le -and $info.N -gt $r.Le) { return '' }
            if ($null -ne $r.Hi -and $info.N -ge $r.Hi) { return '' }
            return $r.'判定'
        }
        '文字' { if (($info.V -ne '' -and $r.Rx.IsMatch($info.V)) -or ($info.VC -ne '' -and $r.Rx.IsMatch($info.VC))) { return $r.'判定' }; return '' }
        '判定' { if ($info.HK -ne '' -and $NAVI.ContainsKey($info.HK)) { return $NAVI[$info.HK] }; return '' }
    }
    return ''
}
$MEMO = @{}
$HasAge = @{}
foreach ($lst in @($Inai, $NavR)) { foreach ($r in $lst) { if ($null -ne $r.AgeF -or $null -ne $r.AgeT) { $HasAge[$r['項目']] = $true } } }   # List どうしは + でつなげない (PowerShell 5.1 でエラーになる)
function Judge([string]$key, $who, $info) {
    # 同じ項目・同じ値・同じ性別 (・年齢の決まりがあれば年齢) の判定は覚えておいて使い回す
    $mk = "$key|$($who.Sex)|$($who.Kijun)|$(if ($HasAge.ContainsKey($key)) { $who.Age })|$($info.V)|$($info.VC)|$($info.HK)"
    if ($MEMO.ContainsKey($mk)) { return $MEMO[$mk] }
    $out = Judge-Core $key $who $info
    $MEMO[$mk] = $out
    return $out
}
function Judge-Core([string]$key, $who, $info) {
    $res = ''; $note = ''
    if ($InaiBy.ContainsKey($key)) {
        foreach ($r in $InaiBy[$key]) { if (Cond-Ok $r $who) { $res = Eval-Rule $r $info; if ($res -ne '') { $note = "院内: $($r.'メモ')"; break } } }
        if ($res -eq '') { $note = '基準なし (院内の表)' }
    } elseif ($NavBy.ContainsKey($key)) {
        foreach ($r in $NavBy[$key]) { if (Cond-Ok $r $who) { $res = Worse $res (Eval-Rule $r $info) } }
        $note = $(if ($res -eq '') { '基準なし (健診ナビの設定)' } else { '健診ナビの設定' })
    } else {
        return $null
    }
    if ($res -ne '' -and $CapBy.ContainsKey($key)) {
        foreach ($r in $CapBy[$key]) { if ((Cond-Ok $r $who) -and $RANK[$res] -gt $RANK[$r.'判定']) { $note = "院内: $($r.'メモ') ($res → $($r.'判定'))"; $res = $r.'判定' } }
    }
    return @{ L = $res; Note = $note }
}
function Navi-To([string]$hk) {
    if ($hk -eq '') { return '' }
    if ($hk -eq 'J') { return 'J' }
    if ($NAVI.ContainsKey($hk)) { return $NAVI[$hk] }
    return '?'
}
function Cmp($c) {
    if ($c.L -eq '' -or $c.Navi -eq '') { return '' }
    if ($c.NaviTo -eq 'J') { return 'J' }
    if ($c.NaviTo -eq $c.L) { return '一致' }
    if (-not $RANK.ContainsKey($c.NaviTo)) { return '?' }
    if ($RANK[$c.L] -gt $RANK[$c.NaviTo]) { return '学会が重い' } else { return '学会が軽い' }
}

$Results = @()
$Detail = New-Object System.Collections.Generic.List[object]   # 4 万行近くになるので += は使わない
$NoRule = @{}
$useOther = $false
$done = 0
foreach ($pp in $people) {
    $done++
    if ($done % 10 -eq 0) { Write-Host -NoNewline '.' }
    $pk = [string]$pp.PK
    $who = @{ Sex = $(switch ([string]$pp.SEX) { '1' { '男' } '2' { '女' } default { '' } }); Kijun = [string]$pp.KIJUN; Age = $null }
    $bd = [datetime]::MinValue
    if ($pp.BIRTH -isnot [DBNull] -and [datetime]::TryParse([string]$pp.BIRTH, [ref]$bd) -and $pp.KDATE -isnot [DBNull]) {
        $kd = [datetime]$pp.KDATE
        $age = $kd.Year - $bd.Year; if ($kd.Month -lt $bd.Month -or ($kd.Month -eq $bd.Month -and $kd.Day -lt $bd.Day)) { $age-- }
        $who.Age = [double]$age
    }
    $d = $(if ($DB.ContainsKey($pk)) { $DB[$pk] } else { @{} })
    # BMI の値が無ければ身長・体重から出す
    if ((-not $d.ContainsKey('001217') -or $null -eq $d['001217'].N) -and $d.ContainsKey('001211') -and $d.ContainsKey('001212') -and $d['001211'].N -gt 0 -and $null -ne $d['001212'].N) {
        $bmi = [math]::Round($d['001212'].N / [math]::Pow($d['001211'].N / 100, 2), 1)
        $d['001217'] = @{ V = $bmi.ToString('0.0', $Inv); VC = ''; HK = ''; N = $bmi }
    }
    $kk = [ordered]@{}
    foreach ($name in $KUKURI.Keys) { $kk[$name] = @{ L = ''; Items = @() } }
    $kk[$OTHER] = @{ L = ''; Items = @() }
    foreach ($key in ($d.Keys | Sort-Object)) {
        if ($Excl.ContainsKey($key)) { continue }
        $info = $d[$key]
        if ($info.V -eq '' -and $info.VC -eq '' -and $info.HK -eq '') { continue }
        $jr = Judge $key $who $info
        if ($null -eq $jr) { continue }
        $kname = $(if ($ItemK.ContainsKey($key)) { $ItemK[$key].Kukuri } else { $OTHER })
        if ($kname -eq $OTHER) { $useOther = $true }
        $iname = $(if ($ItemK.ContainsKey($key) -and $ItemK[$key].Name) { $ItemK[$key].Name } elseif ($NavBy.ContainsKey($key)) { $NavBy[$key][0].'項目名' } elseif ($InaiBy.ContainsKey($key)) { $InaiBy[$key][0].'項目名' } else { $key })
        $kk[$kname].L = Worse $kk[$kname].L $jr.L
        $shown = $(if ($info.V -ne '') { $info.V } elseif ($info.VC -ne '') { "(結果CD $($info.VC))" } else { '(判定のみ)' })
        $kk[$kname].Items += ("{0} {1} → {2} ({3})" -f $iname, $shown, $(if ($jr.L) { $jr.L } else { '-' }), $jr.Note)
        if ($jr.L -eq '') { $nk = "$key $iname|$shown"; $NoRule[$nk] = 1 + [int]$NoRule[$nk] }
        $Detail.Add([PSCustomObject][ordered]@{
            受診日 = $pp.YMD; 受付番号 = $pp.UKE; 性別 = $who.Sex; 年齢 = $who.Age; コース = $pp.COURSE; くくり = $kname; 項目CD = $key; 項目 = $iname; 値 = $shown
            学会判定 = $jr.L; 根拠 = $jr.Note; ナビ判定 = $info.HK; ナビ換算 = (Navi-To $info.HK); 見比べ = (Cmp @{ L = $jr.L; Navi = $info.HK; NaviTo = (Navi-To $info.HK) })
        })
    }
    $total = ''
    $o = [ordered]@{ 受診日 = $pp.YMD; 受付番号 = $pp.UKE; 性別 = $who.Sex; コース = $pp.COURSE }
    foreach ($name in @($KUKURI.Keys) + @($OTHER)) {
        if ($name -ne $OTHER) { $total = Worse $total $kk[$name].L }
        $ccd = $(if ($name -eq $OTHER) { '' } else { $KUKURI[$name] })
        $nkh = $(if ($ccd -ne '' -and $d.ContainsKey($ccd)) { $d[$ccd].HK } else { '' })
        $o[$name] = @{ L = $kk[$name].L; Navi = $nkh; NaviTo = (Navi-To $nkh); Items = $kk[$name].Items }
    }
    $nt = ''
    foreach ($tc in $TOTAL_CDS) {
        if ($nt -eq '' -and $d.ContainsKey($tc)) {
            $nt = $d[$tc].HK
            if ($nt -eq '' -and $d[$tc].V -match '^\s*([A-J])') { $nt = $Matches[1] }
        }
    }
    $o['総合'] = @{ L = $total; Navi = $nt; NaviTo = (Navi-To $nt); Items = @() }
    $Results += [PSCustomObject]$o
}
Write-Host ''
$Results = @($Results | Sort-Object 受診日, @{ Expression = { $x = 0; if ([int]::TryParse([string]$_.'受付番号', [ref]$x)) { $x } else { [int]::MaxValue } } })

# ---- 見比べの集計 ----
$SHOW = @('総合') + @($KUKURI.Keys) + $(if ($useOther) { @($OTHER) } else { @() })
$Sum = foreach ($name in $SHOW) {
    $cnt = @{ '一致' = 0; '学会が重い' = 0; '学会が軽い' = 0; 'J' = 0; '?' = 0; '' = 0 }
    $dist = @{ 'A' = 0; 'B' = 0; 'C' = 0; 'D' = 0 }
    foreach ($r in $Results) { $cnt[(Cmp $r.$name)]++; if ($RANK.ContainsKey($r.$name.L)) { $dist[$r.$name.L]++ } }
    [PSCustomObject][ordered]@{ くくり = $name; A = $dist.A; B = $dist.B; C = $dist.C; D = $dist.D
        一致 = $cnt['一致']; 学会が重い = $cnt['学会が重い']; 学会が軽い = $cnt['学会が軽い']; ナビ治療中 = $cnt['J'] }
}
$ItemSum = @($Detail | Where-Object { $_.'見比べ' -in @('学会が重い','学会が軽い') } | Group-Object 項目CD, 項目, 見比べ | Sort-Object Count -Descending |
    ForEach-Object { [PSCustomObject][ordered]@{ 項目 = $_.Group[0].'項目'; 項目CD = $_.Group[0].'項目CD'; 見比べ = $_.Group[0].'見比べ'; 人数 = $_.Count } })
Line
$Sum | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
if ($ItemSum.Count -gt 0) {
    Write-Host ' 項目ごとに健診ナビと違ったもの:' -ForegroundColor Yellow
    $ItemSum | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
}

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
<h1>人間ドック学会の文字の判定 と 健診ナビの判定 の見比べ</h1>
<p class="note">$(HtmlEnc $cond)　／　$(@($Results).Count) 人　／　作成 $(Get-Date -Format 'yyyy/MM/dd HH:mm')　／　DBは読むだけ</p>
<p class="note">判定: A 異常なし ／ B 軽度異常 ／ C 要再検査・生活改善 ／ D 要精密検査・治療。　マスの上が学会の文字の判定、下の小さい字が健診ナビの判定 (→ 学会の文字に読み替え)。</p>
<p class="note"><span style="background:#ffe0e0">赤</span> = 学会の方が重い　<span style="background:#e0ecff">青</span> = 学会の方が軽い　<span style="background:#f2f2f2">灰</span> = 健診ナビが J (治療中)。マスにマウスを乗せると項目ごとの値・判定・根拠が出る。</p>
<p class="note">基準: 院内の表 dock_kijun.csv ($(@($Inai).Count) 行) → 健診ナビの設定から作った表 dock_kijun_navi.csv ($(@($NavR).Count) 行)。</p>
<h2>1. くくりごとの人数と見比べ</h2><table><tr><th>くくり</th><th>A</th><th>B</th><th>C</th><th>D</th><th>一致</th><th>学会が重い</th><th>学会が軽い</th><th>ナビ治療中</th></tr>
"@
foreach ($s in $Sum) {
    A ("<tr><td class=l>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td></tr>" -f (HtmlEnc $s.'くくり'), $s.A, $s.B, $s.C, $s.D, $s.'一致', $s.'学会が重い', $s.'学会が軽い', $s.'ナビ治療中')
}
A '</table>'
if ($ItemSum.Count -gt 0) {
    A '<h2>2. 項目ごとに健診ナビと違ったもの (院内の表に行を足す手がかり)</h2><table><tr><th>項目</th><th>項目CD</th><th>見比べ</th><th>人数</th></tr>'
    foreach ($s in $ItemSum) { A ("<tr><td class=l>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f (HtmlEnc $s.'項目'), (HtmlEnc $s.'項目CD'), (HtmlEnc $s.'見比べ'), $s.'人数') }
    A '</table>'
}
A '<h2>3. 1 人ずつ</h2><table><tr><th>受診日</th><th>受付番号</th><th>性別</th><th>コース</th>'
foreach ($name in $SHOW) { A "<th>$(HtmlEnc $name)</th>" }
A '</tr>'
foreach ($r in $Results) {
    A ("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td>" -f (HtmlEnc $r.'受診日'), (HtmlEnc $r.'受付番号'), (HtmlEnc $r.'性別'), (HtmlEnc $r.'コース'))
    foreach ($name in $SHOW) {
        $c = $r.$name
        $cls = switch (Cmp $c) { '学会が重い' { 'heavy' } '学会が軽い' { 'light' } 'J' { 'j' } default { '' } }
        $tip = (@($c.Items) -join "`n")
        if ($c.Navi -ne '') { $tip += "`n健診ナビ: $($c.Navi)" }
        $naviTxt = $(if ($c.Navi -ne '') { "<br><small>ナビ $(HtmlEnc $c.Navi)$(if ($c.NaviTo -ne '' -and $c.NaviTo -ne $c.Navi -and $c.NaviTo -ne 'J') { '→' + (HtmlEnc $c.NaviTo) })</small>" } else { '' })
        A ("<td class='{0}' title='{1}'><span class='{2}'>{3}</span>{4}</td>" -f $cls, (HtmlEnc $tip.Trim()), (HtmlEnc $c.L), $(if ($c.L) { HtmlEnc $c.L } else { '<small>-</small>' }), $naviTxt)
    }
    A '</tr>'
}
A '</table>'
if ($NoRule.Count -gt 0) {
    A '<h2>4. 基準に当てはまらなかった値 (判定を付けていない)</h2><table><tr><th>項目</th><th>値</th><th>人数</th></tr>'
    foreach ($k in ($NoRule.Keys | Sort-Object)) { $kv = $k -split '\|', 2; A ("<tr><td class=l>{0}</td><td class=l>{1}</td><td>{2}</td></tr>" -f (HtmlEnc $kv[0]), (HtmlEnc $kv[1]), $NoRule[$k]) }
    A '</table><p class="note">院内の表 dock_kijun.csv に行を足すか、値の入れ方を確かめてください。</p>'
}
A '<h2>5. 院内の表 (form\dock_kijun.csv)</h2><table><tr><th>種類</th><th>項目</th><th>性別</th><th>以上</th><th>以下</th><th>未満</th><th>文字</th><th>判定</th><th>出典</th><th>メモ</th></tr>'
foreach ($r in $Inai) {
    A ("<tr><td>{0}</td><td class=l>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td class=l>{6}</td><td><span class='{7}'>{7}</span></td><td>{8}</td><td class=l>{9}</td></tr>" -f (HtmlEnc $r.'種類'), (HtmlEnc "$($r.'項目') $($r.'項目名')"), (HtmlEnc $r.'性別'), (HtmlEnc $r.'以上'), (HtmlEnc $r.'以下'), (HtmlEnc $r.'未満'), (HtmlEnc $r.'文字'), (HtmlEnc $r.'判定'), (HtmlEnc $r.'出典'), (HtmlEnc $r.'メモ'))
}
A '</table></body></html>'
[System.IO.File]::WriteAllText($htmlPath, $sbH.ToString(), (New-Object System.Text.UTF8Encoding $true))

Write-Host " 書き出しました (氏名なし):" -ForegroundColor Green
Write-Host "   $htmlPath"
Write-Host "   $csvPath  (項目ごとの値・判定・根拠)"
if ($NoRule.Count -gt 0) { Write-Host (" 基準に当てはまらなかった値が {0} 種類あります (HTML の 4)。" -f $NoRule.Count) -ForegroundColor Yellow }
if (-not $NoOpen) { Start-Process $htmlPath }
