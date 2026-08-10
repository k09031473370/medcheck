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

# 通常は 請求一覧.bat が set /p で聞いて引数で渡してくる。
# 直接呼ばれて引数が無いときだけ、ここで聞く。
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

# ---- 状態別の基本料金 (form\seikyu_prices.csv) ----
#   会社,コースCD,状態,会社請求,健保請求
#   協会けんぽの料金表そのままの形。胃部X線や便潜血をやらなかった人は
#   「減額」ではなく状態別の料金行に切り替える(料金表がそういう作りのため)。
#   この表が最優先で、行が無いコースだけ健診ナビのマスタ(T_COURSE3)を使う。
$pricesPath = Join-Path $PSScriptRoot 'form\seikyu_prices.csv'
$prices = @()
if (Test-Path $pricesPath) {
    $prices = @(Import-Csv $pricesPath -Encoding UTF8 | Where-Object {
        (Normalize-Text $_.状態) -ne '' -and -not ((Normalize-Text $_.会社).StartsWith('#')) })
}
function Find-Price([string]$dantaiMei, [string]$course, [string]$state) {
    foreach ($p in $prices) {
        $pc = Normalize-Text $p.会社
        if ($pc -ne '' -and $dantaiMei -notlike "*$pc*") { continue }
        if ((Normalize-Text $p.コースCD) -ne $course) { continue }
        if ((Normalize-Text $p.状態) -ne $state) { continue }
        return $p
    }
    return $null
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
    $chk  = @()   # 減額になった人 (目視確認用)
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

        # マスタの料金は「料金表CSVに行が無いとき」の代替。年度が古いことがある
        $masterBase = [int]$r.DANTAI_RYOUKIN
        $masterKenpo = [int]$r.KENPO_RYOUKIN

        # ---- オプション明細 (T_RYOUKIN。受付で入力した料金) ----
        $optSum = 0
        $optKenpo = 0   # オプションの健保負担分 (マンモ・子宮など協会補助のあるもの)
        $optNames = @()
        $optCds = @{}   # 金額つきで入っていたオプションの項目CD (二重加算防止に使う)
        $opt = Invoke-DbQuery $conn @'
SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS CD, ISNULL(r.GOUKEI,0) AS GOUKEI,
       ISNULL(r.KENPO_RYOUKIN,0) AS KENPO_RYOUKIN, k.MEISYO1
FROM T_RYOUKIN r
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
WHERE r.PK_SEQ = @p AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) <> ''
'@ @{ p = $pk }
        foreach ($o in $opt.Rows) {
            $optKenpo += [int]$o.KENPO_RYOUKIN
            $g = [int]$o.GOUKEI
            if ($g -eq 0) { continue }
            $optSum += $g
            $optCds[[string]$o.CD] = $g
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

        # 料金表に状態別の行があるコースは、そのコースに検査が含まれている証拠。
        # その場合は「枠が無い」も未実施として扱う(巡回でバリウム車が無い日などは
        # 健診ナビに枠すら作られないことがあるため、枠の有無だけでは取りこぼす)。
        $ixPriced  = ($null -ne (Find-Price $dantaiMei $course '胃部X線未実施'))
        $benPriced = ($null -ne (Find-Price $dantaiMei $course '便潜血未実施'))

        $ixMiss   = (-not (Test-Done '077300A')) -and ((Test-Waku '077300A') -or $ixPriced)
        $psaDone  = (Test-Done '067818')                                     # PSA: 結果あり
        # 便潜血: 1回目=069245 / 2回目=069246
        $benFrame = ((Test-Waku '069245') -or (Test-Waku '069246') -or $benPriced)
        $benCount = @('069245','069246' | Where-Object { Test-Done $_ }).Count
        # 尿検査(尿蛋白069206 / 尿糖069207 / 尿潜血069211)。1つも結果が無ければ未実施。
        # 協会は税抜260円(税込286=自己負担80+請求206)を減額する。
        # 2026/06/26 サンテックの請求明細書で確認。
        $nyoFrame = @('069206','069207','069211' | Where-Object { Test-Waku $_ }).Count -gt 0
        $nyoMiss  = ($nyoFrame -and (@('069206','069207','069211' | Where-Object { Test-Done $_ }).Count -eq 0))

        # ---- 状態を決めて、基本料金を料金表から引く ----
        # 胃部X線と便潜血は独立に欠けるので、組合せも状態として持つ。
        # 例: 「胃部X線未実施+便潜血未実施」。料金表にその行があればそれを使い、
        #     無ければ胃部X線だけの行へ落として警告を出す。
        $state = '通常'
        $benState = ''
        if ($benFrame -and $benCount -eq 0) { $benState = '便潜血未実施' }
        elseif ($benFrame -and $benCount -eq 1) { $benState = '便潜血1本のみ' }

        if ($ixMiss -and $benState -ne '') {
            $combo = '胃部X線未実施+' + $benState
            if (Find-Price $dantaiMei $course $combo) {
                $state = $combo
            } else {
                $state = '胃部X線未実施'
                $warn += ("{0} {1}: 胃部X線と便潜血の両方が未実施です。「{2}」の料金が seikyu_prices.csv に無いため胃欠の料金にしています" -f $uke, $name, $combo)
            }
        }
        elseif ($ixMiss)        { $state = '胃部X線未実施' }
        elseif ($benState -ne '') { $state = $benState }

        # 減額になる人は金額が変わるので、必ず一覧に出して目視確認してもらう。
        # 「結果がまだ入っていないだけ」を「受けなかった」と取り違えると請求を誤る。
        if ($state -ne '通常') {
            $chk += ("{0} {1} … {2} (胃部X線の枠 {3} / 便潜血 {4}本)" -f $uke, $name, $state,
                     $(if (Test-Waku '077300A') { 'あり' } else { 'なし' }), $benCount)
        }

        $priceRow = Find-Price $dantaiMei $course $state
        if (-not $priceRow -and $state -ne '通常') {
            $warn += ("{0} {1}: 状態「{2}」の料金が seikyu_prices.csv にありません (通常料金で計算)" -f $uke, $name, $state)
            $priceRow = Find-Price $dantaiMei $course '通常'
        }
        if ($priceRow) {
            $base = 0; $kenpo = 0
            [void][int]::TryParse((Normalize-Text $priceRow.会社請求), [ref]$base)
            [void][int]::TryParse((Normalize-Text $priceRow.健保請求), [ref]$kenpo)
        } else {
            $base = $masterBase; $kenpo = $masterKenpo
            $warn += ("コース {0}: seikyu_prices.csv に料金が無いため健診ナビのマスタの額 ({1}円) を使いました。年度が古い可能性があるので確認してください" -f $course, $masterBase)
        }
        $kenpo += $optKenpo   # マンモ・子宮など、オプションにも協会補助があるぶんを合算

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
                '^PSA実施$'        {
                    $hit = $psaDone
                    # 受付の料金入力にPSAが金額つきで入っていれば、二重になるのでルール加算はしない
                    if ($hit -and (@('OPJ001','OPI001','OP0002') | Where-Object { $optCds.ContainsKey($_) })) {
                        $hit = $false
                    }
                    break
                }
                '^便潜血未実施$'    { $hit = ($benFrame -and $benCount -eq 0); break }
                '^便潜血1本のみ$'   { $hit = ($benFrame -and $benCount -eq 1); break }
                '^尿検査未実施$'    { $hit = $nyoMiss; break }
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
            料金区分 = $state
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
        料金区分 = ''
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
    $lines | Format-Table 受付番号, 氏名, 年齢, コース, 料金区分, 基本料金, オプション, 調整, 会社請求額, 胃部X線, 便潜血, PSA -AutoSize |
        Out-String -Width 220 | Write-Host
    Write-Host ("会社請求 合計: {0:N0} 円 (基本 {1:N0} + オプション {2:N0} + 調整 {3:N0})" -f $sumBill, $sumBase, $sumOpt, $sumAdj) -ForegroundColor Green
    Write-Host ("健保への請求 (参考): {0:N0} 円" -f $sumKenpo)
    Write-Host ''
    Write-Host "出力: $OutCsv" -ForegroundColor Cyan
    if ($chk.Count -gt 0) {
        Write-Host ''
        Write-Host ('【減額になった人 {0} 名】結果がまだ入っていないだけでないか確認してください' -f $chk.Count) -ForegroundColor Yellow
        $chk | ForEach-Object { Write-Host "  ・$_" }
    }
    if ($warn.Count -gt 0) {
        Write-Host ''
        Write-Host '【確認が必要】' -ForegroundColor Yellow
        $warn | Sort-Object -Unique | ForEach-Object { Write-Host "  ・$_" -ForegroundColor Yellow }
    }
}
finally { $conn.Close() }
