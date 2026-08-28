<#
.SYNOPSIS
  取込の練習用データ(CSV)を作る (make_testdata.ps1)

.DESCRIPTION
  DBに実際に入っている受診者と検査枠を読んで、取込ツールに読ませる
  「結果が入ったファイル」をそれらしい値で作る。
  ダミーの受診者を作る必要はない。すでにある受診枠をそのまま使う。

  ★DBは読むだけ。一切書き込まない。
  ★作ったCSVを start.bat で読み込めば、取込の一連の流れを試せる。
    書込先は 接続先.txt が指す先なので、練習用DBに向けてから使うこと。

  値はそれらしい範囲の乱数。医学的に正しい組合せではないので、
  取込の練習にだけ使うこと。

.EXAMPLE
  # その日の受診者から、芝浦AI様式の練習データを作る
  powershell -ExecutionPolicy Bypass -File make_testdata.ps1 -Ymd 2026/06/26

  # 様式と人数と出力先を指定
  powershell -ExecutionPolicy Bypass -File make_testdata.ps1 -Ymd 2026/06/26 -Mapping mapping_rian.csv -Max 10 -Out C:\temp\test.csv

  # どの日に何人いるか一覧を出す
  powershell -ExecutionPolicy Bypass -File make_testdata.ps1 -List
#>
[CmdletBinding()]
param(
    [string]$Ymd,                                  # 受診日 (例 2026/06/26)
    [string]$Mapping = 'mapping_shibaura_ai2.csv', # どの様式で作るか
    [string]$Out,                                  # 出力先。省略時はデスクトップ
    [int]$Max = 0,                                 # 人数の上限 (0=全員)
    [switch]$List,                                 # 受診日の一覧を出して終わる
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
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Normalize-Ymd')
Invoke-Expression (Get-Part 'function Normalize-Ymd' 'function Parse-CsvText')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Get-CurrentKensa')

$MapDir = Join-Path $PSScriptRoot 'form'

# ---- それらしい値を作る。項目コードごとの範囲 ----
# ここに無い数値項目は 1〜10 の適当な値になる。
$RANGE = @{
    '001211' = @(150.0, 180.0, 1)   # 身長
    '001212' = @( 45.0,  90.0, 1)   # 体重
    '001215' = @( 65.0,  95.0, 1)   # 腹囲
    '001253' = @(105, 138, 0)       # 最高血圧
    '001254' = @( 60,  88, 0)       # 最低血圧
    '001255' = @(105, 138, 0)       # 最高血圧2回目
    '001256' = @( 60,  88, 0)       # 最低血圧2回目
}
$VISION = @('0.7', '0.9', '1.0', '1.2', '1.5')

$rand = New-Object System.Random 20260828
function New-Num($cd) {
    if ($RANGE.ContainsKey($cd)) {
        $lo, $hi, $dp = $RANGE[$cd]
        $v = $lo + ($rand.NextDouble() * ($hi - $lo))
        if ($dp -eq 0) { return [string][int][math]::Round($v) }
        return [string][math]::Round($v, $dp)
    }
    return [string]($rand.Next(1, 11))
}

# ---- 対応表を読む ----
$mapPath = if ([System.IO.Path]::IsPathRooted($Mapping)) { $Mapping } else { Join-Path $MapDir $Mapping }
if (-not (Test-Path $mapPath)) { throw "対応表がありません: $mapPath" }
$mapRows = @(Import-Csv -Path $mapPath -Encoding UTF8 | Where-Object {
    ([string]$_.Col).Trim() -ne '' -and -not ([string]$_.Col).StartsWith('#')
})
if ($mapRows.Count -eq 0) { throw "対応表に行がありません: $mapPath" }

# 対応表の DETECT=HEADER から、見出し行に必ず入れる語を拾う
$detectKws = @()
foreach ($ln in (Get-Content $mapPath -TotalCount 8 -Encoding UTF8)) {
    if ($ln -match '^#\s*DETECT\s*=\s*(?i)HEADER\s*:\s*(.+)$') {
        $detectKws = @($Matches[1] -split '[,|]' | ForEach-Object { $_.Trim() } |
                       Where-Object { $_ -ne '' -and -not $_.StartsWith('!') })
        break
    }
}

# 列番号 → 何を入れるか
$slots = @{}
foreach ($m in $mapRows) {
    $c = [int]$m.Col
    $slots[$c] = @{ Kind = [string]$m.Kind; Cd = ([string]$m.KOMOKU_CD).Trim(); Label = [string]$m.Label }
    $c2 = ([string]$m.Col2).Trim()
    if ($c2 -ne '') { $slots[[int]$c2] = @{ Kind = 'SHUBETSU'; Cd = ''; Label = ([string]$m.Label + ' 種別') } }
}
$nCol = ($slots.Keys | Measure-Object -Maximum).Maximum

$conn = Open-Db
try {
    # ---- 受診日の一覧 ----
    if ($List -or -not $Ymd) {
        $dt = Invoke-DbQuery $conn @'
SELECT TOP 40 g.KEN_YMD AS 受診日, COUNT(*) AS 人数
FROM T_KANJA_G g
GROUP BY g.KEN_YMD
ORDER BY g.KEN_YMD DESC
'@ $null
        Write-Host ''
        Write-Host '=== 受付済みの受診日 (新しい順) ===' -ForegroundColor Cyan
        foreach ($r in $dt.Rows) { Write-Host ("  {0}   {1,4} 人" -f $r.受診日, [int]$r.人数) }
        Write-Host ''
        Write-Host '  -Ymd <受診日> を付けて実行すると、その日の練習データを作ります。' -ForegroundColor Yellow
        return
    }

    $y = Normalize-Ymd $Ymd
    if (-not $y) { $y = $Ymd }

    # ---- その日の受診者 ----
    $people = Invoke-DbQuery $conn @'
SELECT g.KEN_NO, g.PK_SEQ, j.KANJI_SIMEI, j.KANA_SIMEI
FROM T_KANJA_G g
JOIN T_KENSIN s ON s.PK_SEQ = g.PK_SEQ
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
WHERE g.KEN_YMD = @y AND s.F_TORIKESI = 0
ORDER BY LEN(LTRIM(RTRIM(g.KEN_NO))), g.KEN_NO
'@ @{ y = $y }
    if ($people.Rows.Count -eq 0) {
        throw "その受診日に受付済みの人がいません: $y  (-List で受診日の一覧が出ます)"
    }

    $rows = @($people.Rows)
    if ($Max -gt 0 -and $rows.Count -gt $Max) { $rows = $rows[0..($Max - 1)] }

    Write-Host ''
    Write-Host '=== 取込の練習データを作る ===' -ForegroundColor Cyan
    Write-Host ("  受診日 : {0}" -f $y)
    Write-Host ("  人数   : {0} 人" -f $rows.Count)
    Write-Host ("  様式   : {0} ({1} 列)" -f (Split-Path $mapPath -Leaf), $nCol)
    Write-Host '  ※ DBは読むだけです。値はそれらしい乱数で、医学的な整合はありません。' -ForegroundColor DarkGray

    # ---- 見出し行 ----
    $head = @()
    $qn = 0
    for ($c = 1; $c -le $nCol; $c++) {
        if (-not $slots.ContainsKey($c)) { $head += ''; continue }
        $s = $slots[$c]
        $lb = $s.Label
        if ($lb -like '問診*') { $qn++; $lb = ('Q{0} {1}' -f $qn, ($lb -replace '^問診\d*\s*', '')) }
        elseif ($s.Kind -eq 'KENNO')  { $lb = '受付NO' }
        elseif ($s.Kind -eq 'KENYMD') { $lb = '受診日' }
        $head += $lb
    }
    # 自動判別に必要な語が見出しに無ければ、最後の空き列に足す
    $joined = $head -join ','
    foreach ($k in $detectKws) {
        if ($joined -notlike "*$k*") {
            $head += $k
            $joined = $head -join ','
        }
    }

    # ---- 1人1行 ----
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add(($head | ForEach-Object { if ($_ -match '[,"]') { '"' + $_.Replace('"','""') + '"' } else { $_ } }) -join ',')

    foreach ($p in $rows) {
        $kenNo = Normalize-KenNo ([string]$p.KEN_NO)
        # その人が実際に持っている検査枠だけを埋める (枠なしのエラーを出さないため)
        $have = @{}
        $k = Invoke-DbQuery $conn 'SELECT LTRIM(RTRIM(KOMOKU_CD)) AS CD FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $p.PK_SEQ }
        foreach ($r in $k.Rows) { $have[[string]$r.CD] = 1 }

        $vals = @()
        for ($c = 1; $c -le $head.Count; $c++) {
            if (-not $slots.ContainsKey($c)) { $vals += ''; continue }
            $s = $slots[$c]
            $v = ''
            switch ($s.Kind) {
                'KENYMD'    { $v = $y }
                'KENNO'     { $v = $kenNo }
                'NAMEKANJI' { $v = Normalize-Text ([string]$p.KANJI_SIMEI) }
                'NAMEKANA'  { $v = Normalize-Text ([string]$p.KANA_SIMEI) }
                'SHUBETSU'  { $v = '1' }                        # 視力の種別 1=裸眼
                'IGNORE'    { $v = '' }
                'NOFRAME'   { $v = '' }
                default {
                    # 枠が無い項目は空欄のままにする
                    if ($s.Cd -ne '' -and -not $have.ContainsKey($s.Cd)) { $v = '' }
                    else {
                        switch ($s.Kind) {
                            'VALUE'    { $v = New-Num $s.Cd }
                            'VISION'   { $v = $VISION[$rand.Next(0, $VISION.Count)] }
                            'MONSHIN'  { $v = '1' }
                            'NYOU'     { $v = '1' }
                            'CHORYOKU' { $v = '1' }
                            'SHOKENCD'  { $v = '1' }            # 1 = 異常なし
                            'SHOKENCD2' { $v = '1' }
                            'SHOKENCD5' { $v = '' }
                            'SHOKEN'    { $v = '' }
                            'SHOKEN2'   { $v = '' }
                            'KOJINNO'   { $v = '' }             # 個人マスタを書き換えるので空に
                            default     { $v = '' }
                        }
                    }
                }
            }
            $vals += $v
        }
        $lines.Add(($vals | ForEach-Object { if ($_ -match '[,"]') { '"' + $_.Replace('"','""') + '"' } else { $_ } }) -join ',')
    }

    if (-not $Out) {
        $desk = [Environment]::GetFolderPath('Desktop')
        $Out = Join-Path $desk ('練習データ_{0}_{1}.csv' -f ($y -replace '[/\-]',''), ((Split-Path $mapPath -Leaf) -replace '^mapping_|\.csv$',''))
    }
    # Excel から保存したCSVと同じ Shift_JIS で書く
    [System.IO.File]::WriteAllLines($Out, $lines, [System.Text.Encoding]::GetEncoding(932))

    # ---- 出来たファイルを、本物の判別器にかけて確かめる ----
    $Csv = $Out
    Invoke-Expression (Get-Part 'function Detect-MappingPath' 'function Load-Mapping')
    $det = Detect-MappingPath $Out
    Write-Host ''
    Write-Host ("出力: {0}  ({1} 行)" -f $Out, $lines.Count) -ForegroundColor Cyan
    if ($det) {
        $detName = Split-Path $det -Leaf
        if ($detName -eq (Split-Path $mapPath -Leaf)) {
            Write-Host ("[確認] 取込ツールは この様式を {0} と判別します。想定どおりです。" -f $detName) -ForegroundColor Green
        } else {
            Write-Host ("[注意] 取込ツールは {0} と判別しました。想定は {1} です。" -f $detName, (Split-Path $mapPath -Leaf)) -ForegroundColor Red
            Write-Host '        取込時に「対応表」を手で選び直してください。' -ForegroundColor Yellow
        }
    } else {
        Write-Host '[注意] 取込ツールが様式を判別できませんでした。取込時に対応表を手で選んでください。' -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host '次の手順:' -ForegroundColor Yellow
    Write-Host '  1. start.bat を開く'
    Write-Host '  2. 「参照...」でこのCSVを選ぶ'
    Write-Host '  3. 「3. プレビュー」で中身を確認する (ここではまだ書き込みません)'
    Write-Host '  4. 問題なければ「4. 書込実行」'
    Write-Host '  5. 戻したくなったら 元に戻す.bat'
}
finally { $conn.Close() }
