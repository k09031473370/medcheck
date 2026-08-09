<#
.SYNOPSIS
  会社宛の請求一覧CSVを作る (seikyu_list.ps1)

.DESCRIPTION
  健診ナビのDBを読むだけで、指定日・指定会社の受診者の請求一覧を作る。
  DBには一切書き込まない。

  1人1行:
    受付番号 / 氏名 / 年齢 / コース / 基本料金(会社負担) / オプション計 /
    調整(胃部X線未実施の減額など) / 会社請求額 / 健保請求(参考) / 内訳

  金額の出どころ:
    基本料金     … T_COURSE3 (団体料金=会社負担、健保料金=協会けんぽ等へ請求する分)
    オプション   … T_RYOUKIN (受付で料金入力した明細。GOUKEIを合算)
    調整         … form\seikyu_rules.csv (胃部X線未実施の減額・PSA加算など)

  実施の判定:
    胃部X線未実施 … 077300A の枠があり結果が空
    PSA実施       … 067818 に結果が入っている

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File seikyu_list.ps1 -Ymd 2026/06/26 -Dantai サンテック
#>
[CmdletBinding()]
param(
    [string]$Ymd,                 # 受診日 'YYYY/MM/DD'
    [string]$Dantai,              # 会社名の一部 (M_DANTAI.MEISYO1 を部分一致)
    [string]$OutCsv,              # 出力先。省略時はデスクトップに 請求一覧_会社_日付.csv
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)

$ErrorActionPreference = 'Stop'

# form_import.ps1 から接続まわりを借りる (restore_backup.ps1 と同じやり方)
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Normalize-KenNo')
Invoke-Expression (Get-Part 'function Normalize-Ymd' 'function Parse-CsvText')
Invoke-Expression (Get-Part 'function Resolve-ConnectionString' 'function Get-CurrentKensa')

if (-not $Ymd)    { $Ymd    = Read-Host '受診日 (例 2026/06/26)' }
if (-not $Dantai) { $Dantai = Read-Host '会社名の一部 (例 サンテック)' }
$y = Normalize-Ymd $Ymd
if (-not $y) { throw "受診日「$Ymd」を解釈できません。2026/06/26 のように入力してください。" }
$Dantai = $Dantai.Trim()
if ($Dantai -eq '') { throw '会社名が空です。' }

# ---- 調整ルール (form\seikyu_rules.csv) ----
#   会社,コースCD,条件,金額,備考
#   会社は部分一致(空=全社)。コースCDは空=全コース。条件は 胃部X線未実施 / PSA実施。
#   金額は 会社請求額への増減 (減額はマイナスで書く)。
$rulesPath = Join-Path $PSScriptRoot 'form\seikyu_rules.csv'
$rules = @()
if (Test-Path $rulesPath) {
    $rules = @(Import-Csv $rulesPath -Encoding UTF8 | Where-Object {
        (Normalize-Text $_.条件) -ne '' -and -not ((Normalize-Text $_.会社).StartsWith('#')) })
}

$conn = Open-Db
try {
    # ---- 対象者 ----
    $people = Invoke-DbQuery $conn @'
SELECT s.PK_SEQ, s.UKE_NO_KENSA, j.KANJI_SIMEI, j.D_SEINEN, s.COURSE_CD,
       c1.MEISYO AS COURSE_MEI, d.MEISYO1 AS DANTAI_MEI,
       ISNULL(c3.DANTAI_RYOUKIN,0) AS DANTAI_RYOUKIN, ISNULL(c3.KENPO_RYOUKIN,0) AS KENPO_RYOUKIN
FROM T_KENSIN s
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c1 ON c1.COURSE_CD = s.COURSE_CD AND c1.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE3 c3 ON c3.COURSE_CD = s.COURSE_CD AND c3.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN = @y AND s.F_TORIKESI = 0 AND d.MEISYO1 LIKE '%' + @d + '%'
ORDER BY LEN(LTRIM(RTRIM(s.UKE_NO_KENSA))), s.UKE_NO_KENSA
'@ @{ y = $y; d = $Dantai }

    if ($people.Rows.Count -eq 0) {
        Write-Host "該当者がいません (受診日=$y 会社名に「$Dantai」を含む)" -ForegroundColor Yellow
        return
    }
    $dantaiMei = [string]$people.Rows[0].DANTAI_MEI
    Write-Host ("対象: {0} / {1} / {2} 人" -f $dantaiMei, $y, $people.Rows.Count) -ForegroundColor Cyan

    $lines = @()
    $warn = @()
    $sumBase = 0; $sumOpt = 0; $sumAdj = 0; $sumBill = 0; $sumKenpo = 0

    foreach ($r in $people.Rows) {
        $pk = $r.PK_SEQ
        $name = Normalize-Text ([string]$r.KANJI_SIMEI)
        $uke = Normalize-Text ([string]$r.UKE_NO_KENSA)
        $course = Normalize-Text ([string]$r.COURSE_CD)
        $courseMei = Normalize-Text ([string]$r.COURSE_MEI)

        # 年齢 (受診日時点。N_NENREIが空の環境なので生年月日から計算)
        $age = ''
        $sei = Normalize-Text ([string]$r.D_SEINEN)
        $dt = [datetime]::MinValue
        if ([datetime]::TryParse($sei, [ref]$dt)) {
            $kdt = [datetime]::Parse($y)
            $a = $kdt.Year - $dt.Year
            if ($kdt -lt $dt.AddYears($a)) { $a-- }
            $age = $a
        }

        $base = [int]$r.DANTAI_RYOUKIN
        $kenpo = [int]$r.KENPO_RYOUKIN

        # ---- オプション明細 (T_RYOUKIN。受付で入力した料金) ----
        $optSum = 0
        $optNames = @()
        $opt = Invoke-DbQuery $conn @'
SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS CD, ISNULL(r.GOUKEI,0) AS GOUKEI, k.MEISYO1
FROM T_RYOUKIN r
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
WHERE r.PK_SEQ = @p AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) <> ''
'@ @{ p = $pk }
        foreach ($o in $opt.Rows) {
            $g = [int]$o.GOUKEI
            if ($g -eq 0) { continue }
            $optSum += $g
            $nm = Normalize-Text ([string]$o.MEISYO1)
            if ($nm -eq '') { $nm = [string]$o.CD }
            $optNames += ('{0} {1}円' -f ($nm -replace '[【】]', ''), $g)
        }

        # ---- 実施状況 (その人の全検査行を1回で読んで、条件はメモリ上で判定する) ----
        $kensa = Invoke-DbQuery $conn @'
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS CD, KEKKA FROM T_KENSA WHERE PK_SEQ = @p
'@ @{ p = $pk }
        $waku = @{}   # 項目CD → 結果 (枠がある項目だけキーが存在する)
        foreach ($k in $kensa.Rows) {
            $waku[[string]$k.CD] = Normalize-Text ([string]$k.KEKKA)
        }
        function Test-Waku([string]$cd) { return $waku.ContainsKey($cd) }                       # 枠がある
        function Test-Done([string]$cd) { return ($waku.ContainsKey($cd) -and $waku[$cd] -ne '') } # 結果あり

        $ixMiss   = ((Test-Waku '077300A') -and -not (Test-Done '077300A'))  # 胃部X線: 枠があるのに結果が無い
        $psaDone  = (Test-Done '067818')                                     # PSA: 結果あり
        # 便潜血: 1回目=069245 / 2回目=069246
        $benFrame = ((Test-Waku '069245') -or (Test-Waku '069246'))
        $benCount = @('069245','069246' | Where-Object { Test-Done $_ }).Count

        # ---- 調整ルール適用 ----
        $adjSum = 0
        $adjNames = @()
        foreach ($rule in $rules) {
            $rc = Normalize-Text $rule.会社
            $rcs = Normalize-Text $rule.コースCD
            $cond = Normalize-Text $rule.条件
            if ($rc -ne '' -and $dantaiMei -notlike "*$rc*") { continue }
            if ($rcs -ne '' -and $rcs -ne $course) { continue }
            $hit = $false
            switch -Regex ($cond) {
                '^胃部X線未実施$'   { $hit = $ixMiss; break }
                '^PSA実施$'        { $hit = $psaDone; break }
                '^便潜血未実施$'    { $hit = ($benFrame -and $benCount -eq 0); break }
                '^便潜血1本のみ$'   { $hit = ($benFrame -and $benCount -eq 1); break }
                # 汎用: 「実施:項目CD」=その項目に結果がある / 「未実施:項目CD」=枠があるのに結果が無い
                # 新しい加減算が出てきたら、コードを直さずルール行の追加だけで対応できる。
                '^実施:(.+)$'      { $hit = (Test-Done $Matches[1].Trim()); break }
                '^未実施:(.+)$'    { $c = $Matches[1].Trim(); $hit = ((Test-Waku $c) -and -not (Test-Done $c)); break }
                default            { $warn += "ルールの条件「$cond」は未対応です (seikyu_rules.csv)"; continue }
            }
            if (-not $hit) { continue }
            $amt = 0
            [void][int]::TryParse((Normalize-Text $rule.金額), [ref]$amt)
            if ($amt -eq 0) {
                $warn += ("{0} {1}: 「{2}」に該当しますが金額が未設定です (seikyu_rules.csv に金額を入れてください)" -f $uke, $name, $cond)
            }
            $adjSum += $amt
            $adjNames += ('{0} {1}円' -f $cond, $amt)
        }

        $bill = $base + $optSum + $adjSum
        $sumBase += $base; $sumOpt += $optSum; $sumAdj += $adjSum; $sumBill += $bill; $sumKenpo += $kenpo

        $lines += [pscustomobject]@{
            受付番号 = $uke; 氏名 = $name; 年齢 = $age
            コース = $courseMei
            基本料金 = $base
            オプション = $optSum
            調整 = $adjSum
            会社請求額 = $bill
            健保請求_参考 = $kenpo
            胃部X線 = $(if (-not (Test-Waku '077300A')) { '対象外' } elseif ($ixMiss) { '未実施' } else { '実施' })
            便潜血 = $(if (-not $benFrame) { '対象外' } else { "$benCount本" })
            PSA = $(if ($psaDone) { '実施' } else { '' })
            内訳 = (@($optNames + $adjNames) -join ' / ')
        }
    }

    # 合計行
    $lines += [pscustomobject]@{
        受付番号 = ''; 氏名 = ('合計 ' + $people.Rows.Count + '名'); 年齢 = ''
        コース = ''
        基本料金 = $sumBase; オプション = $sumOpt; 調整 = $sumAdj
        会社請求額 = $sumBill; 健保請求_参考 = $sumKenpo
        胃部X線 = ''; 便潜血 = ''; PSA = ''; 内訳 = ''
    }

    if (-not $OutCsv) {
        $desk = [Environment]::GetFolderPath('Desktop')
        $OutCsv = Join-Path $desk ('請求一覧_{0}_{1}.csv' -f $dantaiMei, ($y -replace '/', ''))
    }
    $lines | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding Default

    Write-Host ''
    $lines | Format-Table 受付番号, 氏名, 年齢, コース, 基本料金, オプション, 調整, 会社請求額, 胃部X線, 便潜血, PSA -AutoSize |
        Out-String -Width 220 | Write-Host
    Write-Host ("会社請求 合計: {0:N0} 円 (基本 {1:N0} + オプション {2:N0} + 調整 {3:N0})" -f $sumBill, $sumBase, $sumOpt, $sumAdj) -ForegroundColor Green
    Write-Host ("健保への請求 (参考): {0:N0} 円" -f $sumKenpo)
    Write-Host ''
    Write-Host "出力: $OutCsv" -ForegroundColor Cyan
    if ($warn.Count -gt 0) {
        Write-Host ''
        Write-Host '【確認が必要】' -ForegroundColor Yellow
        $warn | Sort-Object -Unique | ForEach-Object { Write-Host "  ・$_" -ForegroundColor Yellow }
    }
}
finally { $conn.Close() }
