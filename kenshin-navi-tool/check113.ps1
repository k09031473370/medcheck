<#
  石屋: 外部点検 (2 回目) の指摘を確かめる (check113.ps1)。DB は読むだけ。結果は out\check113_<日時>_*.csv と画面。
   A. 自動判定の前に保存した医師指示事項 (backup\comment_keep_*.csv) のうち、今どの枠にも同じ文が無いもの (先生に見てもらう)
   B. 医師指示事項が 11 行を超える人 / 11012 以降に文がある人 (結果票に印字されないおそれ)
   C. 【判定】眼底 600060 と【協会判定】眼底指導区分 610100 が食い違う人
   D. 服薬歴 (問診 084001〜084003) と 既往歴 (高血圧 19 / 糖尿病 48 / 高脂血症 47、薬剤治療中) と J の対応がずれている人
  check113.bat から実行。
#>
[CmdletBinding()]
param(
    [string]$Keep = 'comment_keep_20261018_091339.csv,comment_keep_20261018_100522.csv,comment_keep_20261018_100745.csv',
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
$BackupDir = Join-Path $PSScriptRoot 'backup'
$OutDir = Join-Path $PSScriptRoot 'out'
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
$ts = Get-Date -Format 'yyyyMMdd_HHmm'
function Save([object[]]$rows, [string]$tag) {
    if ($rows.Count -eq 0) { return '' }
    $f = Join-Path $OutDir ("check113_{0}_{1}.csv" -f $ts, $tag)
    $rows | Export-Csv -Path $f -NoTypeInformation -Encoding UTF8
    return $f
}
function Squash([string]$t) { return ((Normalize-Text $t) -replace '\s', '') }

$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$SLOTS = @(1..20 | ForEach-Object { [string](11000 + $_) })
$slotLit = "'" + ($SLOTS -join "','") + "'"
$YES = @('1','はい','有','あり')

$conn = Open-Db
try {
    $base = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.KOJIN_ID, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, ISNULL(k.KEKKA,'') AS V, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND (LTRIM(RTRIM(k.KOMOKU_CD)) IN ($slotLit) OR LTRIM(RTRIM(k.KOMOKU_CD)) IN ('600060','610100','600150','600340','600320','084001','084002','084003'))
WHERE $ISHIYA
"@).Rows
    $kiou = (Invoke-DbQuery $conn @"
SELECT DISTINCT kk.KOJIN_ID, LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), kk.BYOMEI_CD),''))) AS BC, LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), kk.CHIRYO),''))) AS C
FROM T_KIOU kk WHERE kk.KOJIN_ID IN (SELECT s.KOJIN_ID FROM T_KENSIN s WHERE $ISHIYA)
"@).Rows
} finally { $conn.Close() }

$Ppl = [ordered]@{}
foreach ($r in $base) {
    $pk = [string]$r.PK_SEQ
    if (-not $Ppl.Contains($pk)) { $Ppl[$pk] = @{ Kojin = [string]$r.KOJIN_ID; Uke = [string]$r.UKE; Ymd = [string]$r.YMD; Name = [string]$r.NAME; C = @{} } }
    $Ppl[$pk].C[[string]$r.CD] = @{ V = (Normalize-Text ([string]$r.V)); H = ([string]$r.H).ToUpper() }
}
function Comments($c) { return @($SLOTS | Where-Object { $c.ContainsKey($_) -and $c[$_].V -ne '' -and $c[$_].V -ne '#' } | ForEach-Object { $c[$_].V }) }

# 分類: 先生に見てもらう必要が高いものを「その他」に残す
$TEMPLATES = @(
    '腎機能検査に軽度異常があります。経過観察し、定期的に検査して下さい。', '腎機能検査に異常があります。精査を受けて下さい。',
    '【糖代謝】　高値です。経過観察の必要があります。', '【糖代謝】　糖尿病が疑われますので精密検査の必要があります。',
    '【中性脂肪】　高値です。飲酒に注意しバランスの良い食事と適度な運動を心がけ、経過観察の必要があります。', '【中性脂肪】　高値です。精密検査の必要があります。',
    '眼底に所見を認めますので、経過観察が必要です。症状があれば眼科を受診して下さい。', '眼底に異常があります。眼科受診し、精密検査を受けて下さい。'
) | ForEach-Object { Squash $_ }
$JK = @( @{ Kk = '600150'; Pat = '血圧' }, @{ Kk = '600340'; Pat = '糖代謝|糖尿病|血糖|ＨｂＡ１ｃ|HbA1c' }, @{ Kk = '600320'; Pat = '脂質|中性脂肪|コレステロール|ＬＤＬ|ＨＤＬ' } )
function Classify([string]$old, $c, $now) {
    if ($TEMPLATES -contains (Squash $old)) { return '1 ツールの定型文 (今は健診ナビの文などで説明あり)' }
    foreach ($j in $JK) { if ($old -match $j.Pat -and $old -notmatch '^\s*【尿糖】' -and $c.ContainsKey($j.Kk) -and $c[$j.Kk].H -eq 'J') { return '2 治療中(J) のため治療継続の文に変わった' } }
    if ($old -match '胸部' -and @($now | Where-Object { $_ -match '胸部' }).Count -gt 0) { return '3 胸部X線の文を差し替えた (先生確認)' }
    if ($old -match '^\s*(【[^】]+】)') { $tag = $Matches[1]; if (@($now | Where-Object { $_.StartsWith($tag) }).Count -gt 0) { return '4 同じ【】の文が自動判定で言い回しを変えて入っている' } }
    return '5 その他 (先生確認)'
}
# ---- A. 消えた医師指示事項 ----
$A = @()
foreach ($name in ($Keep -split ',')) {
    $f = Join-Path $BackupDir $name.Trim()
    if (-not (Test-Path $f)) { Write-Host ("  保存ファイルがありません: {0}" -f $f) -ForegroundColor Yellow; continue }
    foreach ($k in @(Import-Csv -Path $f -Encoding UTF8)) {
        $pk = [string]$k.PK_SEQ
        if (-not $Ppl.Contains($pk)) { continue }
        $now = Comments $Ppl[$pk].C
        $nowSq = @($now | ForEach-Object { Squash $_ })
        if ($nowSq -contains (Squash $k.値)) { continue }
        $A += [PSCustomObject][ordered]@{ 受診日 = $Ppl[$pk].Ymd; 受付番号 = $Ppl[$pk].Uke; 氏名 = $Ppl[$pk].Name; 分類 = (Classify $k.値 $Ppl[$pk].C $now)
            自動判定前にあった文 = $k.値; 今の医師指示事項 = ($now -join ' / ') }
    }
}
# ---- B. 行数 ----
$B = @()
foreach ($pk in $Ppl.Keys) {
    $c = $Ppl[$pk].C; $n = (Comments $c).Count
    $late = @($SLOTS | Where-Object { [int]$_ -ge 11012 -and $c.ContainsKey($_) -and $c[$_].V -ne '' -and $c[$_].V -ne '#' })
    $long = @((Comments $c) | Where-Object { $_.Length -gt 60 }).Count
    if ($n -gt 11 -or $late.Count -gt 0 -or ($n -ge 10 -and $long -gt 0)) {
        $B += [PSCustomObject][ordered]@{ 受診日 = $Ppl[$pk].Ymd; 受付番号 = $Ppl[$pk].Uke; 氏名 = $Ppl[$pk].Name; 行数 = $n; '60字を超える文' = $long; '11012以降' = ($late -join ',') }
    }
}
# ---- C. 眼底と協会判定 ----
$Ceye = @()
foreach ($pk in $Ppl.Keys) {
    $c = $Ppl[$pk].C
    $g = $(if ($c.ContainsKey('600060')) { $c['600060'].H } else { '' }); $k = $(if ($c.ContainsKey('610100')) { $c['610100'].H } else { '(枠なし)' })
    if ($g -ne '' -and $g -ne $k) { $Ceye += [PSCustomObject][ordered]@{ 受診日 = $Ppl[$pk].Ymd; 受付番号 = $Ppl[$pk].Uke; 氏名 = $Ppl[$pk].Name; 判定眼底 = $g; 協会眼底指導区分 = $k } }
}
# ---- D. 服薬・既往歴・J ----
$KI = @{}; foreach ($r in $kiou) { $KI[[string]$r.KOJIN_ID + '|' + ([string]$r.BC).TrimStart('0') + '|' + ([string]$r.C).TrimStart('0')] = 1 }
$MAP = @( @{ Q = '084001'; Cd = '19'; Kk = '600150'; Nm = '血圧' }, @{ Q = '084002'; Cd = '48'; Kk = '600340'; Nm = '糖' }, @{ Q = '084003'; Cd = '47'; Kk = '600320'; Nm = '脂質' } )
$D = @()
foreach ($pk in $Ppl.Keys) {
    $p = $Ppl[$pk]; $c = $p.C
    foreach ($m in $MAP) {
        $q = $(if ($c.ContainsKey($m.Q)) { $c[$m.Q].V } else { '' })
        $med = $YES -contains $q
        $hasK = $KI.ContainsKey($p.Kojin + '|' + $m.Cd + '|5')
        $isJ = $c.ContainsKey($m.Kk) -and $c[$m.Kk].H -eq 'J'
        if ($med -ne $hasK -or $med -ne $isJ) {
            $D += [PSCustomObject][ordered]@{ 受診日 = $p.Ymd; 受付番号 = $p.Uke; 氏名 = $p.Name; 種類 = $m.Nm; 服薬歴 = $q; 既往歴_薬剤治療中 = $(if ($hasK) { 'あり' } else { 'なし' }); くくり判定 = $(if ($c.ContainsKey($m.Kk)) { $c[$m.Kk].H } else { '' }) }
        }
    }
}

# ---- E. 参考: 2026 年の石屋以外の受診者で、判定眼底 600060 と 協会眼底指導区分 610100 の組み合わせ (読み替えの対応を見る) ----
$conn = Open-Db
try {
    $E = @((Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) AS 判定眼底, LTRIM(RTRIM(ISNULL(b.HANTEI_KIGO,''))) AS 協会眼底指導区分, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA a ON a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '600060'
JOIN T_KENSA b ON b.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(b.KOMOKU_CD)) = '610100'
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND NOT ($ISHIYA) AND LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) <> ''
GROUP BY LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))), LTRIM(RTRIM(ISNULL(b.HANTEI_KIGO,'')))
ORDER BY 1, 2
"@).Rows | ForEach-Object { [PSCustomObject][ordered]@{ 判定眼底 = $_.判定眼底; 協会眼底指導区分 = $_.協会眼底指導区分; 人数 = $_.人数 } })
} finally { $conn.Close() }

Write-Host ''
Write-Host ' 外部点検の指摘の確認 (読むだけ)' -ForegroundColor Cyan
$fa = Save $A 'A_消えた文'; $fb = Save $B 'B_行数'; $fc = Save $Ceye 'C_眼底と協会'; $fd = Save $D 'D_服薬と既往歴'
Write-Host ("  A. 自動判定前にあって今は無い医師指示事項: {0} 件 ({1} 人)  {2}" -f $A.Count, @($A | Group-Object 受付番号).Count, $fa)
Write-Host ("  B. 11 行を超える・11012 以降に文がある・10 行以上で長い文がある人: {0} 人  {1}" -f $B.Count, $fb)
Write-Host ("  C. 判定眼底と協会の眼底指導区分が違う人: {0} 人  {1}" -f $Ceye.Count, $fc)
Write-Host ("  D. 服薬歴・既往歴・J がずれている: {0} 件  {1}" -f $D.Count, $fd)
foreach ($g in ($A | Group-Object 分類 | Sort-Object Name)) { Write-Host ("     A の内訳 {0}: {1} 件" -f $g.Name, $g.Count) }
Write-Host ''
Write-Host '  ★先生確認 (A の分類 3・5。名前は出さない):' -ForegroundColor Yellow
foreach ($x in @($A | Where-Object { $_.分類 -match '^[35] ' })) {
    Write-Host ("   {0} {1}  [{2}]" -f $x.受診日, $x.受付番号, $x.分類)
    Write-Host ("      前にあった文: {0}" -f $x.自動判定前にあった文)
    Write-Host ("      今の文      : {0}" -f $x.今の医師指示事項)
}
Write-Host ''
Write-Host '  E. 参考 (石屋以外の 2026 年): 判定眼底 × 協会眼底指導区分'
$E | Format-Table -AutoSize | Out-String -Width 120 | Write-Host
Write-Host ''
Write-Host '  A は先生に見てもらう一覧 (J の人の旧い文・自動判定が言い回しを変えた文も含む)。CSV は受診者の名前入り: 院内だけで使う。' -ForegroundColor Yellow
