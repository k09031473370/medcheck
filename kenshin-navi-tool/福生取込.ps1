<#
.SYNOPSIS
  福生商工会 当日の取込を通しでやる (福生取込.ps1)
  福生取込.bat に受付アプリの JSON をドラッグするだけ。

.DESCRIPTION
  中身は今まで別々に動かしていた2本をそのまま順に呼ぶ。
    受付結果取込.ps1     … 健診ナビに予約がある人へ受付番号 (プレビュー → Y)
    予約取込Excel作成.ps1 … 予約が無い人の Excel を作る
  その前に「事前チェック」で、今日の人の団体にコースが無い／対応表が健診ナビと違う を
  Excel を作る前に見つけて止める (10/6 に赤いセルで止まった原因)。

  流れ
    1. 事前チェック (読むだけ)
    2. 受付番号を入れる        (受付結果取込.ps1 → プレビュー → Y)
    3. 未登録の人の Excel を作る (予約取込Excel作成.ps1)
       → 健診ナビの予約データ取込を開く (開けなければデスクトップを開く)
       → 「個人マスタ登録」「予約登録」を押したら Enter
    4. もう一度受付番号を入れる (3 で登録した人の分)
    5. ★ の一覧と 確認一覧.txt を表示

  健診ナビの予約データ取込を自動で開くには
    form\yoyaku_import_app.txt に exe のフルパスを1行書いておく。
    無ければデスクトップのフォルダを開くので、人が Excel をドロップする。

.EXAMPLE
  福生取込.bat に navi.json をドラッグ
  powershell -ExecutionPolicy Bypass -File 福生取込.ps1 -Json D:\fussa_reception_20261006_130147.json
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Json,
    [string]$Ymd,                       # 省略時は JSON の中の一番新しい受付日
    [string]$DantaiLike = '福生',
    [switch]$SkipCheck,                 # 事前チェックで止まったときに、承知のうえで続けるとき
    [switch]$CheckOnly,                 # 事前チェックだけして終わる (画面付きツール用。問題があれば終了コード 2)
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot
$ps  = (Get-Command powershell).Source

function Banner([string]$t) {
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor Cyan
    Write-Host (' ' + $t) -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor Cyan
}
function Pause-Enter([string]$msg) {
    Write-Host ''
    Write-Host $msg -ForegroundColor Yellow
    [void](Read-Host '  Enter を押すと続きます')
}
# 氏名の照合用: 空白を取り、旧字・異体字を新字に揃える (壽→寿, 髙→高, 﨑→崎 …)
#   健診ナビは個人マスタの字 (旧字のことがある)、受付アプリは名簿の字 (新字) なので、そのままだと別人扱いになる。
$script:OldChars = '壽髙﨑嵜齋齊澤邊邉國廣濱櫻龍惠眞禮與榮德瀧嶋嶌條曾冨靍萬彌藏巖穗祿舩渕籐檜栁槇薭圓濟靑淸瀨黑塲增敎橫每步鷗鹽亞惡應獻劍繪聰戶靜讓驗體醫學寫實將眾數壯莊單戰雙寶豐賴彥晉濵'
$script:NewChars = '寿高崎崎斎斉沢辺辺国広浜桜竜恵真礼与栄徳滝島島条曽富鶴万弥蔵巌穂禄船淵藤桧柳槙稗円済青清瀬黒場増教横毎歩鴎塩亜悪応献剣絵聡戸静譲験体医学写実将衆数壮荘単戦双宝豊頼彦晋浜'
function NoSpace([string]$s) {
    if ($null -eq $s) { return '' }
    $t = (($s -replace '[\s　]', '').Trim())
    if ($t -eq '') { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $t.ToCharArray()) {
        $i = $script:OldChars.IndexOf($ch)
        if ($i -ge 0) { [void]$sb.Append($script:NewChars[$i]) } else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

# ---- 入力 ----
if (-not $Json) { $Json = Read-Host '受付アプリの JSON をドラッグ＆ドロップして Enter' }
$Json = ($Json -replace '^"|"$', '').Trim()
if (-not (Test-Path $Json)) { throw "見つかりません: $Json" }
if ([System.IO.Path]::GetExtension($Json).ToLower() -ne '.json') { throw "JSON ではありません: $Json  (Excel ではなく受付アプリの JSON をドラッグしてください)" }
foreach ($f in @('受付結果取込.ps1', '予約取込Excel作成.ps1', 'form_import.ps1', 'form\fussa_company_map.csv', 'form\fussa_option_map.csv')) {
    if (-not (Test-Path (Join-Path $dir $f))) { throw "$f がありません: $dir" }
}

$doc = Get-Content -LiteralPath $Json -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $doc.people) { throw 'この JSON に people がありません。受付アプリの「健診ナビ連携JSON」を渡してください。' }

# 今日の日付 = JSON の中で一番新しい受付日
if (-not $Ymd) {
    $Ymd = [string](@($doc.people | Where-Object { $_.actual -and $_.actual.checked_in_date } |
             ForEach-Object { [string]$_.actual.checked_in_date } | Sort-Object -Descending) | Select-Object -First 1)
    if (-not $Ymd) { throw 'JSON に受付済みの人がいません。' }
}
$Ymd = ($Ymd -replace '-', '/')
$recv = @($doc.people | Where-Object { $_.actual -and $null -ne $_.actual.reception_number -and (([string]$_.actual.checked_in_date) -replace '-', '/') -eq $Ymd })

Banner "福生 当日取込  $Ymd"
Write-Host ("  JSON      : {0}" -f (Split-Path $Json -Leaf))
Write-Host ("  出力日時  : {0}   revision {1}" -f $doc.exported_at, $doc.revision)
Write-Host ("  {0} に受付した人 : {1} 人" -f $Ymd, $recv.Count)
if ($recv.Count -eq 0) { throw "$Ymd に受付した人が JSON にいません。日付が違うなら -Ymd で指定してください。" }

# ============================================================================
# 1. 事前チェック (読むだけ)
# ============================================================================
Banner '1. 事前チェック (団体のコース・対応表)'
$Core = Join-Path $dir 'form_import.ps1'
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text'    '$XlsxLib = Join-Path')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')

$COURSE = @{ 'A' = 'FA'; 'B' = 'FB'; 'C' = 'FC' }
$cdOf = @{}
foreach ($r in (Import-Csv -Path (Join-Path $dir 'form\fussa_company_map.csv') -Encoding UTF8)) {
    $k = NoSpace $r.'振り分け表の事業所名'
    if ($k -ne '' -and ([string]$r.団体CD).Trim() -ne '') { $cdOf[$k] = ([string]$r.団体CD).Trim() }
}

$problems = @()
$conn = Open-Db
try {
    # 健診ナビに既に予約がある人は、事前チェックの対象から外す (手で作った人・前回入れた人)
    $have = @{}
    $mapCds = (($cdOf.Values | Select-Object -Unique | ForEach-Object { "'$_'" }) -join ',')
    if ($mapCds -eq '') { $mapCds = "''" }
    $dtH = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS KANJI
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '$Ymd'
  AND (d.MEISYO1 LIKE N'%$DantaiLike%' OR LTRIM(RTRIM(s.DANTAI_CD1)) IN ($mapCds))
"@ @{}
    foreach ($r in $dtH.Rows) { $k = NoSpace $r.KANJI; if ($k -ne '') { $have[$k] = 1 } }
    $skipped = 0

    # 今日の人の 団体CD と コース
    $need = @{}   # 団体CD -> @{ Name=会社名; Courses=@{FA=人数} }
    foreach ($p in $recv) {
        if ($have.ContainsKey((NoSpace $p.identity.name))) { $skipped++; continue }
        $co = NoSpace $p.planned.company
        $cd = $cdOf[$co]
        if (-not $cd) { $problems += "対応表に無い会社: $($p.planned.company)  ($($p.identity.name))"; continue }
        $cs = $(if ($p.actual -and [string]$p.actual.course -ne '') { [string]$p.actual.course } else { [string]$p.planned.course })
        if (-not $COURSE.ContainsKey($cs)) { $problems += "コースが A/B/C でない: $($p.identity.name) [$cs]"; continue }
        if (-not $need.ContainsKey($cd)) { $need[$cd] = @{ Name = $p.planned.company; Courses = @{} } }
        $need[$cd].Courses[$COURSE[$cs]] = 1
    }
    if ($need.Count -gt 0) {
        $inCd = (($need.Keys | ForEach-Object { "'$_'" }) -join ',')
        $dt = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(d.DANTAI_CD1)) AS CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS NAME,
       ISNULL(STUFF((SELECT ',' + LTRIM(RTRIM(c.COURSE_CD)) FROM T_COURSE1 c WHERE c.DANTAI_CD1 = d.DANTAI_CD1 FOR XML PATH('')), 1, 1, ''), '') AS COURSES
FROM T_DANTAI1 d WHERE LTRIM(RTRIM(d.DANTAI_CD1)) IN ($inCd)
"@ @{}
        $navi = @{}
        foreach ($r in $dt.Rows) { $navi[[string]$r.CD] = @{ Name = [string]$r.NAME; Courses = @(([string]$r.COURSES) -split ',') } }
        foreach ($cd in $need.Keys) {
            $n = $need[$cd]
            if (-not $navi.ContainsKey($cd)) { $problems += "団体CD $cd が健診ナビに無い (対応表: $($n.Name))"; continue }
            $nm = $navi[$cd].Name
            if ($nm -notlike "*$DantaiLike*" -and $nm -ne 'ILL CLIPPER BARBER SHOP') {
                $problems += "団体CD $cd は健診ナビでは「$nm」。対応表の「$($n.Name)」と合っているか確認"
            }
            foreach ($c in $n.Courses.Keys) {
                if ($navi[$cd].Courses -notcontains $c) { $problems += "団体「$nm」($cd) に コース $c が無い → 健診ナビの画面で足してから" }
            }
        }
    }
}
finally { if ($conn -and $conn.State -eq 'Open') { $conn.Close() } }

if ($skipped -gt 0) { Write-Host ("  (健診ナビに予約がある {0} 人はチェック対象外)" -f $skipped) -ForegroundColor DarkGray }
if ($problems.Count -eq 0) {
    Write-Host ("  OK  {0} 団体ともコースあり・対応表も一致" -f $need.Count) -ForegroundColor Green
    if ($CheckOnly) { exit 0 }
}
elseif ($CheckOnly) {
    Write-Host ('  ★ 直してから進めてください ({0} 件)' -f $problems.Count) -ForegroundColor Red
    $problems | Sort-Object -Unique | ForEach-Object { Write-Host ('    ' + $_) -ForegroundColor Yellow }
    exit 2
}
else {
    Write-Host ('  ★ 直してから進めてください ({0} 件)' -f $problems.Count) -ForegroundColor Red
    $problems | Sort-Object -Unique | ForEach-Object { Write-Host ('    ' + $_) -ForegroundColor Yellow }
    if (-not $SkipCheck) {
        Write-Host ''
        Write-Host '  上を直してから、もう一度 bat を流してください。' -ForegroundColor Yellow
        Write-Host '  直さずに進めると、その人は健診ナビの取込で赤になります (他の人は入ります)。' -ForegroundColor DarkYellow
        $a = Read-Host '  それでも進めるなら Y'
        if ($a -ne 'Y' -and $a -ne 'y') { Write-Host '止めました。'; return }
    }
}

# ============================================================================
# 2. 受付番号を入れる (予約がある人)
# ============================================================================
Banner '2. 受付番号を入れる (健診ナビに予約がある人)'
& $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dir '受付結果取込.ps1') -Json $Json -DantaiLike $DantaiLike
if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) { Write-Host '  受付結果取込 がエラーで終わりました。上のメッセージを確認してください。' -ForegroundColor Red }

# ============================================================================
# 3. 未登録の人の Excel → 健診ナビの予約データ取込
# ============================================================================
Banner "3. 健診ナビに予約が無い人の Excel を作る ($Ymd)"
& $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dir '予約取込Excel作成.ps1') -Json $Json -Ymd $Ymd -DantaiLike $DantaiLike 2>&1 | Tee-Object -Variable excelLog | Out-Host
$xlsx = $null
foreach ($line in @($excelLog)) {
    if ([string]$line -match '\[出力\]\s+(.+?\.xlsx)') { $xlsx = $Matches[1].Trim() }
}
if (-not $xlsx) {
    Write-Host ''
    Write-Host '  新しく登録する人はいません。4 へ進みます。' -ForegroundColor Green
}
else {
    Write-Host ''
    Write-Host ("  Excel: {0}" -f $xlsx) -ForegroundColor Green
    # 健診ナビの予約データ取込を開く (exe のパスが form\yoyaku_import_app.txt にあれば)
    $appTxt = Join-Path $dir 'form\yoyaku_import_app.txt'
    $opened = $false
    if (Test-Path $appTxt) {
        $exe = (Get-Content $appTxt -Encoding UTF8 | Where-Object { $_.Trim() -ne '' -and -not $_.Trim().StartsWith('#') } | Select-Object -First 1)
        if ($exe) { $exe = $exe.Trim() }
        if ($exe -and (Test-Path $exe)) {
            try { Start-Process -FilePath $exe; $opened = $true; Write-Host '  健診ナビの「予約データ取込」を開きました。' -ForegroundColor Green }
            catch { Write-Host "  予約データ取込を開けませんでした: $($_.Exception.Message)" -ForegroundColor Yellow }
        }
    }
    # Excel の場所をエクスプローラーで選択状態にして開く
    try { Start-Process explorer.exe -ArgumentList "/select,`"$xlsx`"" } catch { }
    if (-not $opened) { Write-Host '  健診ナビのメニューから「予約データ取込」を開いてください。' -ForegroundColor Yellow }
    Pause-Enter @"
  次をやってください:
    1. エクスプローラーで選ばれている Excel を、健診ナビの「予約データ取込」にドロップ
    2. 赤いセルが無ければ 「個人マスタ登録」 → 「予約登録」
       (赤があれば、その行を Excel から消して保存し、ドロップし直す。消した人は後で)
    3. 「予約登録中…」が終わるまで待つ
"@
}

# ============================================================================
# 4. もう一度 受付番号 (3 で登録した人)
# ============================================================================
if ($xlsx) {
    Banner '4. 受付番号を入れる (いま登録した人)'
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $dir '受付結果取込.ps1') -Json $Json -DantaiLike $DantaiLike
}

# ============================================================================
# 5. まとめ
# ============================================================================
Banner '5. 残っていること'
Write-Host '  上の 4 (無ければ 2) の「お知らせ」で ★ が付いた人を、健診ナビの画面で対応してください。'
if ($xlsx) {
    $memo = [System.IO.Path]::ChangeExtension($xlsx, '.確認一覧.txt')
    if (Test-Path $memo) {
        Write-Host ''
        Write-Host ("  確認一覧: {0}" -f $memo) -ForegroundColor DarkGray
        Get-Content $memo -Encoding Default | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor DarkYellow }
    }
}
Write-Host ''
Write-Host '  「更新する人」の数が 3 で登録した人数より少ないときは、足りない人が' -ForegroundColor DarkGray
Write-Host '  「見つかりません」に出ています。健診ナビで予約登録ができているか確認して、' -ForegroundColor DarkGray
Write-Host '  もう一度この bat を流せば、その人だけ入ります。' -ForegroundColor DarkGray
Write-Host ''
Write-Host '完了。' -ForegroundColor Green
