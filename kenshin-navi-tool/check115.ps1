<#
  石屋 3 日分: 印刷前の最終点検 (check115.ps1)。DB は読むだけ。名前は出さない (受付番号だけ)。
   1. 総合判定(編集) 11000 の文字と文章が合っていない人 (例 F なのに「要観察」)
   2. くくり判定 (600xxx) の結果が 1 文字で、判定と違う欄
   3. 総合判定(編集) が、くくり判定のいちばん重い文字と違う人 (J はいちばん重い扱い。先生の判断で戻した 4646・4687 は印を付けて出す)
   4. 総合判定が C 以上なのに、医師指示事項が 1 行も無い人
   5. 枠はあるのに結果が空の主な検査 (身長・体重・血圧・視力・聴力・尿・血液・肝・脂質・糖・腎・心電図・胸部・眼底)
#>
[CmdletBinding()]
param([string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt', [string]$ConnectionString)
$ErrorActionPreference = 'Stop'
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) { $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i); return $src.Substring($i, $j - $i) }
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')

$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$CANON = @{ 'A' = '異常なし'; 'B' = 'ほぼ正常(心配なし)'; 'C' = '要注意'; 'D' = '要観察(要経過観察)'; 'E' = '要再検'; 'F' = '要精検'; 'G' = '要治療'; 'J' = '治療中' }
$RANK = @{ 'A' = 1; 'B' = 2; 'C' = 3; 'D' = 4; 'E' = 5; 'F' = 6; 'G' = 7; 'J' = 8 }
$MANUAL = @('4646','4687')   # 先生の判断で総合判定を下げた人
$MAIN = [ordered]@{ '001211' = '身長'; '001212' = '体重'; '001253' = '最高血圧'; '001254' = '最低血圧'; '069206' = '尿蛋白'; '069207' = '尿糖';
    '069003' = '白血球'; '069005' = '血色素'; '068636' = 'AST'; '068637' = 'ALT'; '068652' = 'γ-GTP'; '068705' = '総コレステロール'; '068708' = 'HDL'; '068709' = 'LDL';
    'R06-0002' = '中性脂肪(未判別)'; '068682' = '血糖(未判別)'; '068689' = 'HbA1c'; '068674' = 'クレアチニン'; '068676' = 'eGFR'; '067112A' = '心電図1'; '077010A' = '胸部X線1'; '600060' = '【判定】眼底' }
$SLOTS = @(1..20 | ForEach-Object { [string](11000 + $_) })

$conn = Open-Db
try {
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD,
       LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600___' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '110__' OR LTRIM(RTRIM(k.KOMOKU_CD)) IN ('$(@($MAIN.Keys) -join "','")'))
"@).Rows
} finally { $conn.Close() }
$Ppl = [ordered]@{}
foreach ($r in $rows) {
    $pk = [string]$r.PK_SEQ
    if (-not $Ppl.Contains($pk)) { $Ppl[$pk] = @{ Uke = [string]$r.UKE; Ymd = [string]$r.YMD; C = @{} } }
    $Ppl[$pk].C[[string]$r.CD] = @{ V = [string]$r.V; H = ([string]$r.H).ToUpper() }
}
$o1 = @(); $o2 = @(); $o3 = @(); $o4 = @(); $o5 = @{}
foreach ($pk in $Ppl.Keys) {
    $p = $Ppl[$pk]; $c = $p.C
    $sh = $(if ($c.ContainsKey('11000')) { $c['11000'].H } else { '' }); $sv = $(if ($c.ContainsKey('11000')) { $c['11000'].V } else { '' })
    if ($sh -eq '') { $o1 += ("{0} {1}: 総合判定(編集) が空" -f $p.Ymd, $p.Uke) }
    elseif ($CANON.ContainsKey($sh) -and $sv -ne $CANON[$sh]) { $o1 += ("{0} {1}: {2} なのに文章が「{3}」" -f $p.Ymd, $p.Uke, $sh, $sv) }
    $max = 0; $maxH = ''
    foreach ($cd in @($c.Keys | Where-Object { $_ -match '^600\d{3}$' })) {
        $h = $c[$cd].H; $v = $c[$cd].V
        if ($h -ne '' -and $v -match '^[A-J]$' -and $v -ne $h) { $o2 += ("{0} {1}: {2} 判定 {3} / 結果 {4}" -f $p.Ymd, $p.Uke, $cd, $h, $v) }
        if ($RANK.ContainsKey($h) -and $RANK[$h] -gt $max) { $max = $RANK[$h]; $maxH = $h }
    }
    if ($sh -ne '' -and $maxH -ne '' -and $sh -ne $maxH) { $o3 += ("{0} {1}: 総合 {2} / くくりの最重 {3}{4}" -f $p.Ymd, $p.Uke, $sh, $maxH, $(if ($MANUAL -contains $p.Uke) { '  (先生の判断で変更済み)' } else { '' })) }
    $nCom = @($SLOTS | Where-Object { $c.ContainsKey($_) -and $c[$_].V -ne '' -and $c[$_].V -ne '#' }).Count
    if ($RANK.ContainsKey($sh) -and $RANK[$sh] -ge 3 -and $nCom -eq 0) { $o4 += ("{0} {1}: 総合 {2} なのに医師指示事項が 0 行" -f $p.Ymd, $p.Uke, $sh) }
    foreach ($cd in $MAIN.Keys) { if ($c.ContainsKey($cd) -and $c[$cd].V -eq '' -and $c[$cd].H -eq '') { if (-not $o5.ContainsKey($cd)) { $o5[$cd] = @() }; $o5[$cd] += $p.Uke } }
}
Write-Host ''
Write-Host (' 石屋 印刷前の最終点検 ({0} 人。名前なし・読むだけ)' -f $Ppl.Count) -ForegroundColor Cyan
function Show($title, $list) { Write-Host ''; Write-Host ("{0}: {1} 件" -f $title, $list.Count) -ForegroundColor $(if ($list.Count -gt 0) { 'Yellow' } else { 'Green' }); foreach ($x in $list) { Write-Host ("   {0}" -f $x) } }
Show '1. 総合判定(編集) の文字と文章が合わない・空' $o1
Show '2. くくり判定の結果の文字が判定と違う' $o2
Show '3. 総合判定(編集) とくくりの最重が違う' $o3
Show '4. 総合 C 以上で医師指示事項が 0 行' $o4
Write-Host ''
Write-Host ("5. 枠はあるのに結果が空の主な検査: {0} 項目" -f $o5.Count) -ForegroundColor $(if ($o5.Count -gt 0) { 'Yellow' } else { 'Green' })
foreach ($cd in $o5.Keys) { Write-Host ("   {0} {1}: {2} 人  ({3})" -f $cd, $MAIN[$cd], $o5[$cd].Count, (($o5[$cd] | Select-Object -First 30) -join ',')) }
