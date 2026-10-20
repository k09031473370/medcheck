<#
.SYNOPSIS
  健診ナビの基準値の設定を読んで、学会判定の表 form\dock_kijun_navi.csv を作る (dock_setup.ps1)

.DESCRIPTION
  T_KIJUN1 (数値項目) と T_KIJUN2 (定性項目) の、コースで使われている基準セットの行を全部読み、
  健診ナビの判定記号 (8段階) を dock_moji.csv で学会の文字 (4段階) に読み替えて書き出す。
  院内の決まり (健診ナビの設定より優先するもの) は form\dock_kijun.csv に手で書く。この表は手で直さない (作り直すと消える)。
  健診ナビの設定を直したら、もう一度 dock_setup.bat を実行する。
  あわせて r_dock_setup.txt に、項目 → くくり の表 (dock_kukuri.csv) に無い項目や、
  健診ナビの「くくり」の設定がありそうなテーブルを書き出す。
  DBは読むだけで、一切変更しない (マスタにも書かない)。
#>
[CmdletBinding()]
param(
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
$OutCsv  = Join-Path $FormDir 'dock_kijun_navi.csv'
$Report  = Join-Path $PSScriptRoot 'r_dock_setup.txt'
function W($t) { $t | Out-File $Report -Append -Encoding Default }
function Read-Form([string]$name) {
    $path = Join-Path $FormDir $name
    if (-not (Test-Path $path)) { throw "表がありません: $path" }
    return @(Get-Content $path -Encoding UTF8 | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' } | ConvertFrom-Csv)
}
function To-NumS([string]$s) {
    $mt = [regex]::Match([string]$s, '-?\d+(\.\d+)?')
    if (-not $mt.Success) { return $null }
    return [double]::Parse($mt.Value, [Globalization.CultureInfo]::InvariantCulture)
}
function Is-Num([string]$s) { return ([string]$s).Trim() -match '^-?\d+(\.\d+)?$' }
function Query-Long($conn, [string]$sql) {
    # 過去の判定を数える問い合わせは時間がかかるので、待ち時間を長くする (form_import の Invoke-DbQuery は既定の 30 秒)
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $sql; $cmd.CommandTimeout = 900
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt); $cmd.Dispose()
    return ,$dt
}
function Col($row, [string]$name) {
    if (-not $row.Table.Columns.Contains($name)) { return '' }
    $v = $row[$name]
    if ($v -is [DBNull] -or $null -eq $v) { return '' }
    return ([string]$v).Trim()
}

$NAVI = @{}
foreach ($m in (Read-Form 'dock_moji.csv')) { $NAVI[([string]$m.'ナビ').Trim()] = ([string]$m.'学会').Trim() }
$KUK = @{}; $KUKCD = @{}; $CDNAME = @{}
foreach ($k in (Read-Form 'dock_kukuri.csv')) {
    $kcd = ([string]$k.'くくりCD').Trim()
    $KUK[([string]$k.'項目').Trim()] = ([string]$k.'くくり').Trim()
    $KUKCD[([string]$k.'項目').Trim()] = $kcd
    if ($kcd -ne '' -and -not $CDNAME.ContainsKey($kcd)) { $CDNAME[$kcd] = ([string]$k.'くくり').Trim() }
}
$OutKuk = Join-Path $FormDir 'dock_kukuri_navi.csv'

Write-Host ''
Write-Host ' 健診ナビの基準値の設定を読んで、学会判定の表を作る (DBは読むだけ)' -ForegroundColor Cyan
$conn = Open-Db
try {
    $sets = @((Invoke-DbQuery $conn "SELECT DISTINCT LTRIM(RTRIM(CONVERT(varchar(20), KIJUN_CD))) AS K, COUNT(*) AS N FROM T_COURSE1 GROUP BY LTRIM(RTRIM(CONVERT(varchar(20), KIJUN_CD)))").Rows)
    $setList = @($sets | ForEach-Object { [string]$_.K } | Where-Object { $_ -match '^[0-9A-Za-z]+$' })
    if ($setList.Count -eq 0) { $setList = @('1') }
    $inSets = "('" + ($setList -join "','") + "')"
    $k1 = @((Invoke-DbQuery $conn "SELECT * FROM T_KIJUN1 WHERE LTRIM(RTRIM(CONVERT(varchar(20), KIJUN_CD))) IN $inSets").Rows)
    $k2 = @((Invoke-DbQuery $conn "SELECT * FROM T_KIJUN2 WHERE LTRIM(RTRIM(CONVERT(varchar(20), KIJUN_CD))) IN $inSets").Rows)
    $names = @{}
    foreach ($r in (Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS NM FROM T_KOMOKU").Rows) { $names[[string]$r.CD] = [string]$r.NM }
    $hanteis = Invoke-DbQuery $conn "SELECT * FROM T_HANTEIS"
    # 健診ナビが実際に付けた判定 (過去 2 年、値と判定ごとの件数)。同じ範囲に違う判定の行が 2 つあるとき、実際に効いている方を選ぶのに使う
    $codes = @((@($k1) + @($k2)) | ForEach-Object { Col $_ 'KOMOKU_CD' } | Where-Object { $_ -match '^[0-9A-Za-z-]+$' } | Select-Object -Unique)
    $actual = @()
    $actualErr = ''
    if ($codes.Count -gt 0) {
        Write-Host ' 健診ナビが過去 2 年に付けた判定を数えています (数分かかることがあります) ...' -ForegroundColor DarkGray
        try {
            $actual = @((Query-Long $conn @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(j.SEIBETU,''))) AS SX, LTRIM(RTRIM(k.KEKKA)) AS V, LTRIM(RTRIM(k.HANTEI_KIGO)) AS H, COUNT(*) AS N
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= DATEADD(year, -2, GETDATE())
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('$($codes -join "','")')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(j.SEIBETU,''))), LTRIM(RTRIM(k.KEKKA)), LTRIM(RTRIM(k.HANTEI_KIGO))
"@).Rows)
        } catch { $actualErr = $_.Exception.Message }
    }
    # くくりの作りの手がかり: くくり項目 (600xxx) の計算式・区分と、くくりを含むセット
    $kukuriKomoku = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 名前, LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 区分, CALC_ORDER AS 計算順, LTRIM(RTRIM(ISNULL(CALC_FORM,''))) AS 計算式, LTRIM(RTRIM(ISNULL(KEKKA_GRP,''))) AS 結果グループ FROM T_KOMOKU WHERE KOMOKU_CD LIKE '600%' OR KOMOKU_CD LIKE '610%' ORDER BY KOMOKU_CD"
    # 項目 → くくり を、健診ナビが過去 1 年に付けた判定から推し量る:
    #   その項目が悪い人で、くくりの判定がその項目より軽いことがほとんど無く (= くくりに入っている)、
    #   くくりの判定がその項目と同じ (= その項目がくくりを重くした) 人がいちばん多いくくり
    $kukStat = @()
    $kukErr = ''
    if ($codes.Count -gt 0) {
        Write-Host ' 項目とくくりの関係を数えています (数分かかることがあります) ...' -ForegroundColor DarkGray
        $RK = "CASE LTRIM(RTRIM(k.HANTEI_KIGO)) WHEN 'A' THEN 1 WHEN 'B' THEN 2 WHEN 'C' THEN 3 WHEN 'D' THEN 4 WHEN 'E' THEN 5 WHEN 'D3' THEN 5 WHEN 'F' THEN 6 WHEN 'D2' THEN 6 WHEN 'G' THEN 7 WHEN 'D1' THEN 7 ELSE 0 END"
        try {
            $kukStat = @((Query-Long $conn @"
WITH R AS (
  SELECT k.PK_SEQ AS PK, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, $RK AS RK
  FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
  WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= DATEADD(year, -1, GETDATE()) AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
    AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600%' OR LTRIM(RTRIM(k.KOMOKU_CD)) IN ('$($codes -join "','")'))
)
SELECT x.CD AS ITEM, g.CD AS KUK,
       SUM(CASE WHEN x.RK > 1 THEN 1 ELSE 0 END) AS ABN,
       SUM(CASE WHEN x.RK > g.RK THEN 1 ELSE 0 END) AS VIOL,
       SUM(CASE WHEN x.RK = g.RK AND x.RK > 1 THEN 1 ELSE 0 END) AS EXPL
FROM R x JOIN R g ON g.PK = x.PK
WHERE g.CD LIKE '600%' AND x.CD NOT LIKE '600%' AND x.RK > 0 AND g.RK > 0
GROUP BY x.CD, g.CD
"@).Rows)
        } catch { $kukErr = $_.Exception.Message }
    }
    $kukuriSet = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(t.SET_CD)) AS セット, LTRIM(RTRIM(t.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 名前 FROM T_SET2 t LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(t.KOMOKU_CD)) WHERE t.SET_CD IN (SELECT SET_CD FROM T_SET2 WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('600320','600340','600310')) ORDER BY 1, 2"
    $komokuCols = Invoke-DbQuery $conn "SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KOMOKU' ORDER BY ORDINAL_POSITION"
    # くくり (どの項目がどのくくりに入るか) の設定がありそうなテーブル: KOMOKU_CD 列を持ち、くくりの 600320 (脂質代謝) が入っているもの
    $cand = @((Invoke-DbQuery $conn @"
SELECT DISTINCT c.TABLE_NAME AS T FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE c.COLUMN_NAME = 'KOMOKU_CD' AND c.TABLE_NAME NOT LIKE 'T_KENSA%' AND c.TABLE_NAME NOT LIKE '%LOG%'
"@).Rows | ForEach-Object { [string]$_.T })
    $found = @()
    foreach ($t in $cand) {
        if ($t -notmatch '^[A-Za-z0-9_]+$') { continue }
        try {
            $cnt = (Invoke-DbQuery $conn "SELECT COUNT(*) AS N FROM [$t] WHERE LTRIM(RTRIM(CONVERT(varchar(50), KOMOKU_CD))) IN ('600320','600340','600310')").Rows[0].N
            if ([int]$cnt -gt 0) {
                $cols = (Invoke-DbQuery $conn "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @t ORDER BY ORDINAL_POSITION" @{ t = $t }).Rows | ForEach-Object { [string]$_.COLUMN_NAME }
                $found += [PSCustomObject]@{ テーブル = $t; くくりの行 = [int]$cnt; 列 = ($cols -join ',') }
            }
        } catch { }
    }
} finally { $conn.Close() }

# ---- 行を作る ----
$cand2 = @()
$skip = @()
$sexName = @{ '' = ''; '0' = ''; '1' = '男'; '2' = '女' }
function New-Row([string]$kind, $r, [string]$lo, [string]$hi, [string]$pat, [string]$memo) {
    $cd = Col $r 'KOMOKU_CD'; $hk = Col $r 'HANTEI_KIGO'; $sx = Col $r 'SEIBETU'
    return [PSCustomObject][ordered]@{
        種類 = $kind; 項目 = $cd; 項目名 = $names[$cd]; 基準CD = (Col $r 'KIJUN_CD'); 性別 = $(if ($sexName.ContainsKey($sx)) { $sexName[$sx] } else { $sx })
        年齢以上 = (Col $r 'NENREI_F'); 年齢以下 = (Col $r 'NENREI_T'); 以上 = $lo; 以下 = $hi; 未満 = ''; 文字 = $pat
        判定 = $NAVI[$hk]; ナビ判定 = $hk; 出典 = '健診ナビ'; メモ = $memo
    }
}
foreach ($r in $k1) {
    $cd = Col $r 'KOMOKU_CD'; $hk = Col $r 'HANTEI_KIGO'
    if (-not $NAVI.ContainsKey($hk)) { $skip += "数値 $cd $($names[$cd]) 判定「$hk」"; continue }
    $lo = Col $r 'SAISYO'; $hi = Col $r 'SAIDAI'; $ren = "連番 $(Col $r 'RENBAN')"
    if (($lo -ne '' -and -not (Is-Num $lo)) -or ($hi -ne '' -and -not (Is-Num $hi))) {
        # 「凝固」「多数」「0.1未満」「0.30以下」のような文字の結果は、その文字と同じときだけ当てる
        $txt = $(if ($lo -ne '' -and -not (Is-Num $lo)) { $lo } else { $hi })
        $cand2 += (New-Row '文字' $r '' '' ('^' + [regex]::Escape($txt) + '$') "$ren / 文字 $txt")
    } else {
        $cand2 += (New-Row '数値' $r $lo $hi '' $ren)
    }
}
foreach ($r in $k2) {
    $cd = Col $r 'KOMOKU_CD'; $hk = Col $r 'HANTEI_KIGO'; $te = Col $r 'TEISEI'
    if (-not $NAVI.ContainsKey($hk)) { $skip += "文字 $cd $($names[$cd]) 「$te」 判定「$hk」"; continue }
    if ($te -eq '') { $skip += "文字 $cd $($names[$cd]) 判定 $hk (定性の文字が空)"; continue }
    $cand2 += (New-Row '文字' $r '' '' ('^' + [regex]::Escape($te) + '$') "連番 $(Col $r 'RENBAN') / 定性 $te")
}

# ---- 健診ナビが実際に付けた判定と照らす ----
$ACT = @{}
foreach ($a in $actual) {
    $cd = [string]$a.CD
    if (-not $ACT.ContainsKey($cd)) { $ACT[$cd] = New-Object System.Collections.Generic.List[object] }
    $v = [string]$a.V
    $ACT[$cd].Add(@{ SX = $(if ($sexName.ContainsKey([string]$a.SX)) { $sexName[[string]$a.SX] } else { '' }); V = $v; N = $(if (Is-Num $v) { [double]::Parse($v.Trim(), [Globalization.CultureInfo]::InvariantCulture) } else { $null }); H = [string]$a.H; C = [int]$a.N })
}
function Sex-Ok([string]$a, [string]$b) { return ($a -eq '' -or $b -eq '' -or $a -eq $b) }
$INFO = @{}
$idx = 0
foreach ($row in $cand2) {
    $row | Add-Member -NotePropertyName _i -NotePropertyValue $idx -Force
    $idx++
    $ag = 0; $di = 0
    $isTxt = ($row.'種類' -eq '文字'); $rlo = To-NumS $row.'以上'; $rhi = To-NumS $row.'以下'
    if ($ACT.ContainsKey($row.'項目')) {
        foreach ($a in $ACT[$row.'項目']) {
            if (-not (Sex-Ok $row.'性別' $a.SX)) { continue }
            if ($isTxt) { $hit = ($a.V -match $row.'文字') }
            else { $hit = ($null -ne $a.N -and ($null -eq $rlo -or $a.N -ge $rlo) -and ($null -eq $rhi -or $a.N -le $rhi)) }
            if ($hit) { if ($a.H -eq $row.'ナビ判定') { $ag += $a.C } else { $di += $a.C } }
        }
    }
    $INFO[$row._i] = @{ Ag = $ag; Di = $di; Score = $(if ($ag + $di -gt 0) { $ag / ($ag + $di) } else { -1 }) }
    $row.'メモ' += " / 実際の判定 一致 $ag・不一致 $di"
}
function Overlap($x, $y) {
    if ($x.'項目' -ne $y.'項目' -or $x.'基準CD' -ne $y.'基準CD' -or -not (Sex-Ok $x.'性別' $y.'性別')) { return $false }
    if ($x.'種類' -ne $y.'種類') { return $false }
    if ($x.'種類' -eq '文字') { return ($x.'文字' -eq $y.'文字') }
    $xl = To-NumS $x.'以上'; $xh = To-NumS $x.'以下'; $yl = To-NumS $y.'以上'; $yh = To-NumS $y.'以下'
    $lo = [math]::Max($(if ($null -eq $xl) { -1e18 } else { $xl }), $(if ($null -eq $yl) { -1e18 } else { $yl }))
    $hi = [math]::Min($(if ($null -eq $xh) { 1e18 } else { $xh }), $(if ($null -eq $yh) { 1e18 } else { $yh }))
    return ($lo -le $hi)
}
function Tie-Win([string]$p, [string]$q) {
    $pd = $p -match '^D\d$'; $qd = $q -match '^D\d$'
    if ($pd -and -not $qd) { return 2 }
    if ($qd -and -not $pd) { return 1 }
    if ($p -eq 'C' -and $q -match '^[DEFG]$') { return 2 }
    if ($q -eq 'C' -and $p -match '^[DEFG]$') { return 1 }
    return 0
}
$drop = @{}
$conflicts = @()
foreach ($g in ($cand2 | Group-Object 項目)) {
    $list = @($g.Group)
    for ($a = 0; $a -lt $list.Count; $a++) {
        for ($b = $a + 1; $b -lt $list.Count; $b++) {
            $x = $list[$a]; $y = $list[$b]
            if ($x.'ナビ判定' -eq $y.'ナビ判定' -or -not (Overlap $x $y)) { continue }
            $sx = $INFO[$x._i].Score; $sy = $INFO[$y._i].Score
            $res = '決められない (どちらも使う。重い方になる)'
            if ($sx -gt $sy) { $drop[$y._i] = $true; $res = "$($x.'ナビ判定') を使う" } elseif ($sy -gt $sx) { $drop[$x._i] = $true; $res = "$($y.'ナビ判定') を使う" }
            else {
                # 実際の判定で決められないとき: 学会そのままの組 (D1/D2/D3、C) より当院の組 (D/E/F/G など) を使う (5 で効いているのはいつも当院の組だった)
                $tw = Tie-Win $x.'ナビ判定' $y.'ナビ判定'
                if ($tw -eq 1) { $drop[$y._i] = $true; $res = "$($x.'ナビ判定') を使う (当院の組)" } elseif ($tw -eq 2) { $drop[$x._i] = $true; $res = "$($y.'ナビ判定') を使う (当院の組)" }
            }
            $conflicts += ("  {0} {1} {2}  [{3}〜{4}{5} {6} 一致{7}/不一致{8}] と [{9}〜{10}{11} {12} 一致{13}/不一致{14}] → {15}" -f $x.'項目', $x.'項目名', $x.'性別',
                $x.'以上', $x.'以下', $(if ($x.'種類' -eq '文字') { $x.'文字' } else { '' }), $x.'ナビ判定', $INFO[$x._i].Ag, $INFO[$x._i].Di,
                $y.'以上', $y.'以下', $(if ($y.'種類' -eq '文字') { $y.'文字' } else { '' }), $y.'ナビ判定', $INFO[$y._i].Ag, $INFO[$y._i].Di, $res)
        }
    }
}
$rows = @($cand2 | Where-Object { -not $drop.ContainsKey($_._i) })
$poor = @($rows | Where-Object { $INFO[$_._i].Di -ge 5 -and $INFO[$_._i].Di -gt $INFO[$_._i].Ag })
$rows = @($rows | Sort-Object 項目, 基準CD, 性別, @{ Expression = { $n = 0; [void][int]::TryParse(($_.'メモ' -replace '^連番 (\d+).*$', '$1'), [ref]$n); $n } })

$head = @(
    "# 健診ナビの基準値の設定から作った表 ($(Get-Date -Format 'yyyy/MM/dd HH:mm') dock_setup.ps1)。★手で直さない。作り直すと消える",
    '# 院内の決まりは dock_kijun.csv に書く (そちらが優先)。判定 = ナビ判定を dock_moji.csv で学会の文字にしたもの',
    "# 基準セット (T_COURSE1.KIJUN_CD): $($setList -join ', ')"
)
$body = @($rows | Select-Object * -ExcludeProperty _i | ConvertTo-Csv -NoTypeInformation)
$lines = @($body[0]) + $head + @($body | Select-Object -Skip 1)
[System.IO.File]::WriteAllLines($OutCsv, [string[]]$lines, (New-Object System.Text.UTF8Encoding $true))

# ---- 項目 → くくり (健診ナビの実際の判定から) ----
$KNAME = @{}; $KOLD = @{}
foreach ($r in $kukuriKomoku.Rows) {
    $KNAME[[string]$r.'項目CD'] = (([string]$r.'名前') -replace '^【判定】', '' -replace '×$', '').Trim()
    if (([string]$r.'名前') -match '×') { $KOLD[[string]$r.'項目CD'] = $true }   # 名前に × が付いたくくりは今は使っていない (循環器× など)
}
$kukRows = @()
$kukNone = @()
foreach ($g in ($kukStat | Group-Object ITEM)) {
    # 偶然の重なりを拾わないよう、悪い人の 2 割以上でくくりと同じ判定になっていることも条件にする
    $okList = @($g.Group | Where-Object { -not $KOLD.ContainsKey([string]$_.KUK) -and [int]$_.ABN -ge 3 -and [int]$_.EXPL -ge [math]::Max(2, [math]::Ceiling(0.2 * [int]$_.ABN)) -and [int]$_.VIOL -le [math]::Max(1, [math]::Floor(0.05 * [int]$_.ABN)) } |
        Sort-Object @{ Expression = { [int]$_.EXPL }; Descending = $true }, @{ Expression = { [int]$_.VIOL } })
    $item = [string]$g.Name
    if ($okList.Count -eq 0) {
        $mx = @($g.Group | Sort-Object @{ Expression = { [int]$_.ABN }; Descending = $true })[0]
        $kukNone += ("  {0} {1}  (悪い人 {2} 人)" -f $item, $names[$item], $mx.ABN)
        continue
    }
    $best = $okList[0]; $kc = [string]$best.KUK
    $kukRows += [PSCustomObject][ordered]@{
        項目 = $item; 項目名 = $names[$item]; くくり = $(if ($CDNAME.ContainsKey($kc)) { $CDNAME[$kc] } elseif ($KNAME.ContainsKey($kc)) { $KNAME[$kc] } else { $kc }); くくりCD = $kc
        悪い人 = [int]$best.ABN; くくりと同じ = [int]$best.EXPL; くくりより重い = [int]$best.VIOL
    }
}
$kukRows = @($kukRows | Sort-Object くくりCD, 項目)
$kb = @($kukRows | ConvertTo-Csv -NoTypeInformation)
$kl = @($(if ($kb.Count) { $kb[0] } else { '"項目","項目名","くくり","くくりCD"' })) + @(
    "# 項目 → くくり を健診ナビが過去 1 年に付けた判定から推し量った表 ($(Get-Date -Format 'yyyy/MM/dd HH:mm') dock_setup.ps1)。★手で直さない",
    '# 手で決めたものは dock_kukuri.csv に書く (そちらが優先)。悪い人 = その項目が B 以上の人、くくりと同じ = くくりの判定がその項目と同じ人、くくりより重い = 項目の方が重かった人'
) + @($kb | Select-Object -Skip 1)
[System.IO.File]::WriteAllLines($OutKuk, [string[]]$kl, (New-Object System.Text.UTF8Encoding $true))

# ---- 報告 ----
"=== 健診ナビの基準値の設定から 学会判定の表を作った $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $Report -Encoding Default
W ''
W "書き出し: $OutCsv"
W ("  数値 {0} 行 / 文字 {1} 行 / 項目 {2} 個" -f @($rows | Where-Object { $_.'種類' -eq '数値' }).Count, @($rows | Where-Object { $_.'種類' -eq '文字' }).Count, @($rows | Select-Object -ExpandProperty 項目 -Unique).Count)
W ''
W '--- 1. コースが使っている基準セット (T_COURSE1.KIJUN_CD) ---'
foreach ($s in $sets) { W ("  {0} : {1} コース" -f $s.K, $s.N) }
W ''
W '--- 2. ナビ判定ごとの行数 (→ 学会の文字) ---'
foreach ($g in ($rows | Group-Object ナビ判定 | Sort-Object Name)) { W ("  {0} → {1} : {2} 行" -f $g.Name, $(if ($NAVI.ContainsKey($g.Name)) { $NAVI[$g.Name] } else { '(dock_moji に無い)' }), $g.Count) }
W ''
W "--- 3. 使わなかった行 (判定が A〜J の 1 文字でない など) $($skip.Count) 行 ---"
foreach ($s in ($skip | Select-Object -First 200)) { W "  $s" }
W ''
W '--- 4. ★項目 → くくり の表 (dock_kukuri.csv) に無い項目 (判定は出すが総合には入れない) ---'
foreach ($g in ($rows | Group-Object 項目 | Sort-Object Name)) {
    if (-not $KUK.ContainsKey($g.Name)) { W ("  {0}  {1}  ({2} 行)" -f $g.Name, $g.Group[0].'項目名', $g.Count) }
}
W ''
W "--- 5. 同じ範囲に違う判定の行があったもの (健診ナビが過去 2 年に実際に付けた判定と合う方を使う。使わなかった行 $($drop.Count) 行) ---"
if ($actualErr) { W "  ★過去の判定を数えられませんでした: $actualErr" }
foreach ($c in $conflicts) { W $c }
W ''
W "--- 5b. 使う行なのに、実際の判定と合わないことが多いもの $($poor.Count) 行 (院内の決まりがある・手で直している などの手がかり) ---"
foreach ($p in $poor) { W ("  {0} {1} {2}  {3}〜{4}{5} {6}  一致 {7}・不一致 {8}" -f $p.'項目', $p.'項目名', $p.'性別', $p.'以上', $p.'以下', $(if ($p.'種類' -eq '文字') { $p.'文字' } else { '' }), $p.'ナビ判定', $INFO[$p._i].Ag, $INFO[$p._i].Di) }
W ''
W "--- 6. ★項目 → くくり を健診ナビの実際の判定から推し量った表 (form\dock_kukuri_navi.csv) $($kukRows.Count) 項目 ---"
if ($kukErr) { W "  ★数えられませんでした: $kukErr" }
W '  ・dock_kukuri.csv (手で決めた表) と違うもの:'
foreach ($r in $kukRows) { if ($KUKCD.ContainsKey($r.'項目') -and $KUKCD[$r.'項目'] -ne $r.'くくりCD') { W ("    {0} {1}: 手の表 {2} ({3}) / 実際の判定 {4} ({5})  悪い人 {6}・同じ {7}・重い {8}" -f $r.'項目', $r.'項目名', $KUK[$r.'項目'], $KUKCD[$r.'項目'], $r.'くくり', $r.'くくりCD', $r.'悪い人', $r.'くくりと同じ', $r.'くくりより重い') } }
W '  ・手の表に無く、実際の判定から足したもの:'
foreach ($g in ($kukRows | Where-Object { -not $KUKCD.ContainsKey($_.'項目') } | Group-Object くくりCD)) { W ("    {0} {1}: {2}" -f $g.Name, $g.Group[0].'くくり', (($g.Group | ForEach-Object { "$($_.'項目名')($($_.'項目') 悪い人$($_.'悪い人')・同じ$($_.'くくりと同じ')・重い$($_.'くくりより重い'))" }) -join '、')) }
W "  ・どのくくりか決められなかった項目 $($kukNone.Count) 個 (悪い人が少ない・どのくくりとも合わない):"
foreach ($t in ($kukNone | Select-Object -First 80)) { W "  $t" }
W ''
W '=== 読み方 ==='
W '  4 に出た項目は dock_kukuri.csv に「項目,項目名,くくり,くくりCD」の行を足すと総合判定に入る。'
W '  6 の「手の表と違うもの」は、どちらが正しいか確かめる。実際の判定の方が正しければ dock_kukuri.csv の行を直すか消す (消せば実際の判定の表が使われる)。'
Write-Host " 書き出しました: $OutCsv" -ForegroundColor Green
Write-Host " 報告: $Report"
Start-Process notepad.exe $Report
