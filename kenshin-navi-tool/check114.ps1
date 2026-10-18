<#
  石屋 9/9・9/10: 業者ファイルでは食後に採血した人 (全員空腹扱いで自動判定した) の、空腹時血糖・空腹時中性脂肪・HbA1c の値と判定 (check114.ps1)
  先生が「空腹時のまま印刷するか、採血時間を入れ直すか」を決めるための一覧。名前は出さない。DB は読むだけ。
  食後の区分は 業者の 123 列ファイル (910_入力画面バックアップ_123列_20261014) の 46 列 (空腹) と 47 列 (採血時間(食後)) から。
#>
[CmdletBinding()]
param([string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt', [string]$ConnectionString)
$ErrorActionPreference = 'Stop'
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) { $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i); return $src.Substring($i, $j - $i) }
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')

$MEAL = [ordered]@{}
foreach ($u in @(4718,4738,4739,4742,4743,4746,4747,4754,4755,4759,4760,4762,4763,4765,4767,4768,4774,4775,4777,4780,4825)) { $MEAL[[string]$u] = '食後3.5時間以上 (随時)' }
foreach ($u in @(4761,4769,4770,4827,4829)) { $MEAL[[string]$u] = '食後1時間 (食直後)' }
foreach ($u in @(4719,4720,4722,4723,4724,4725,4741)) { $MEAL[[string]$u] = '記録なし' }
$CDS = [ordered]@{ 'H30-0011' = '空腹時血糖'; 'H30-0012' = '随時血糖'; '068702' = '空腹時中性脂肪'; 'R06-0003' = '随時中性脂肪'; '068689' = 'HbA1c'; '600340' = '糖代謝'; '600320' = '脂質代謝' }
$conn = Open-Db
try {
    $rows = (Invoke-DbQuery $conn @"
SELECT s.UKE_NO_KENSA AS UKE, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/09','2026/09/10') AND s.UKE_NO_KENSA IN ($(@($MEAL.Keys) -join ','))
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('$(@($CDS.Keys) -join "','")')
"@).Rows
} finally { $conn.Close() }
$V = @{}; foreach ($r in $rows) { $V[[string]$r.UKE + '|' + [string]$r.CD] = ("{0}{1}" -f $r.V, $(if ($r.H) { " (" + $r.H + ")" } else { '' })) }
$out = foreach ($u in $MEAL.Keys) {
    $o = [ordered]@{ 受付番号 = $u; 採血 = $MEAL[$u] }
    foreach ($cd in $CDS.Keys) { $o[$CDS[$cd]] = $(if ($V.ContainsKey($u + '|' + $cd)) { $V[$u + '|' + $cd] } else { '' }) }
    [PSCustomObject]$o
}
Write-Host ''
Write-Host ' 9/9・9/10 食後に採血した人の血糖・中性脂肪 (今は全員 空腹時の欄に入っている。名前なし)' -ForegroundColor Cyan
$out | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
