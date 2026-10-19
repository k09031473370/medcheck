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
function Col($row, [string]$name) {
    if (-not $row.Table.Columns.Contains($name)) { return '' }
    $v = $row[$name]
    if ($v -is [DBNull] -or $null -eq $v) { return '' }
    return ([string]$v).Trim()
}

$NAVI = @{}
foreach ($m in (Read-Form 'dock_moji.csv')) { $NAVI[([string]$m.'ナビ').Trim()] = ([string]$m.'学会').Trim() }
$KUK = @{}
foreach ($k in (Read-Form 'dock_kukuri.csv')) { $KUK[([string]$k.'項目').Trim()] = ([string]$k.'くくり').Trim() }

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
$rows = @()
$skip = @()
$sexName = @{ '' = ''; '0' = ''; '1' = '男'; '2' = '女' }
foreach ($r in $k1) {
    $cd = Col $r 'KOMOKU_CD'; $hk = Col $r 'HANTEI_KIGO'
    if ($hk -notmatch '^[A-J]$') { $skip += "数値 $cd $($names[$cd]) 判定「$hk」"; continue }
    $sx = Col $r 'SEIBETU'
    $rows += [PSCustomObject][ordered]@{
        種類 = '数値'; 項目 = $cd; 項目名 = $names[$cd]; 基準CD = (Col $r 'KIJUN_CD'); 性別 = $(if ($sexName.ContainsKey($sx)) { $sexName[$sx] } else { $sx })
        年齢以上 = (Col $r 'NENREI_F'); 年齢以下 = (Col $r 'NENREI_T'); 以上 = (Col $r 'SAISYO'); 以下 = (Col $r 'SAIDAI'); 未満 = ''; 文字 = ''
        判定 = $(if ($NAVI.ContainsKey($hk)) { $NAVI[$hk] } else { '' }); ナビ判定 = $hk; 出典 = '健診ナビ'; メモ = "連番 $(Col $r 'RENBAN')"
    }
}
foreach ($r in $k2) {
    $cd = Col $r 'KOMOKU_CD'; $hk = Col $r 'HANTEI_KIGO'; $te = Col $r 'TEISEI'
    if ($hk -notmatch '^[A-J]$') { $skip += "文字 $cd $($names[$cd]) 「$te」 判定「$hk」"; continue }
    if ($te -eq '') { $skip += "文字 $cd $($names[$cd]) 判定 $hk (定性の文字が空)"; continue }
    $sx = Col $r 'SEIBETU'
    $rows += [PSCustomObject][ordered]@{
        種類 = '文字'; 項目 = $cd; 項目名 = $names[$cd]; 基準CD = (Col $r 'KIJUN_CD'); 性別 = $(if ($sexName.ContainsKey($sx)) { $sexName[$sx] } else { $sx })
        年齢以上 = (Col $r 'NENREI_F'); 年齢以下 = (Col $r 'NENREI_T'); 以上 = ''; 以下 = ''; 未満 = ''; 文字 = ('^' + [regex]::Escape($te) + '$')
        判定 = $(if ($NAVI.ContainsKey($hk)) { $NAVI[$hk] } else { '' }); ナビ判定 = $hk; 出典 = '健診ナビ'; メモ = "連番 $(Col $r 'RENBAN') / 定性 $te"
    }
}
$rows = @($rows | Sort-Object 項目, 基準CD, 性別, @{ Expression = { $n = 0; [void][int]::TryParse(($_.'メモ' -replace '^連番 (\d+).*$', '$1'), [ref]$n); $n } })

$head = @(
    "# 健診ナビの基準値の設定から作った表 ($(Get-Date -Format 'yyyy/MM/dd HH:mm') dock_setup.ps1)。★手で直さない。作り直すと消える",
    '# 院内の決まりは dock_kijun.csv に書く (そちらが優先)。判定 = ナビ判定を dock_moji.csv で学会の文字にしたもの',
    "# 基準セット (T_COURSE1.KIJUN_CD): $($setList -join ', ')"
)
$body = @($rows | ConvertTo-Csv -NoTypeInformation)
$lines = @($body[0]) + $head + @($body | Select-Object -Skip 1)
[System.IO.File]::WriteAllLines($OutCsv, [string[]]$lines, (New-Object System.Text.UTF8Encoding $true))

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
W '--- 5. 同じ項目・性別で範囲が重なっている行 (重い方を使う) ---'
foreach ($g in ($rows | Where-Object { $_.'種類' -eq '数値' } | Group-Object 項目, 基準CD, 性別, 年齢以上, 年齢以下)) {
    $list = @($g.Group)
    for ($a = 0; $a -lt $list.Count; $a++) {
        for ($b = $a + 1; $b -lt $list.Count; $b++) {
            $x = $list[$a]; $y = $list[$b]
            if ($x.'ナビ判定' -eq $y.'ナビ判定') { continue }
            $xl = To-NumS $x.'以上'; $xh = To-NumS $x.'以下'; $yl = To-NumS $y.'以上'; $yh = To-NumS $y.'以下'
            $lo = [math]::Max($(if ($null -eq $xl) { -1e9 } else { $xl }), $(if ($null -eq $yl) { -1e9 } else { $yl }))
            $hi = [math]::Min($(if ($null -eq $xh) { 1e9 } else { $xh }), $(if ($null -eq $yh) { 1e9 } else { $yh }))
            if ($lo -le $hi) { W ("  {0} {1} {2}  {3}〜{4} {5} と {6}〜{7} {8}" -f $x.'項目', $x.'項目名', $x.'性別', $x.'以上', $x.'以下', $x.'ナビ判定', $y.'以上', $y.'以下', $y.'ナビ判定') }
        }
    }
}
W ''
W '--- 6. ★くくりの設定がありそうなテーブル (KOMOKU_CD に 600320・600340・600310 が入っている) ---'
if ($found.Count -eq 0) { W '  見つかりませんでした' } else { foreach ($f in $found) { W ("  {0} ({1} 行): {2}" -f $f.'テーブル', $f.'くくりの行', $f.'列') } }
W ''
W '--- 7. 参考: T_KOMOKU の列 ---'
W (($komokuCols.Rows | ForEach-Object { "$($_.'列')($($_.'型'))" }) -join ', ')
W ''
W '--- 8. 参考: 判定記号の意味 (T_HANTEIS) ---'
W ($hanteis | Format-Table -AutoSize | Out-String -Width 200)
W '=== 読み方 ==='
W '  4 に出た項目は dock_kukuri.csv に「項目,項目名,くくり,くくりCD」の行を足すと総合判定に入る。'
W '  6 は、健診ナビの「くくり」の決まりがどのテーブルにあるかの手がかり。貼ってもらえれば dock_kukuri.csv を自動で作れるようにする。'
Write-Host " 書き出しました: $OutCsv" -ForegroundColor Green
Write-Host " 報告: $Report"
notepad $Report
