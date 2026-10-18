<#
.SYNOPSIS
  石屋 3 日分の変更一覧を作る (change_list.ps1)。DB は読むだけ。

.DESCRIPTION
  2026/10/18 に服薬歴から既往歴を入れて自動判定をかけ直し、そのあとツールで判定・医師指示事項を直した。その結果を 1 人 1 行にまとめる。
    ・結果票の総合判定 (11000) と 自動判定の総合 (10000)、違えばその理由
    ・治療中 (J) のくくり (血圧・糖代謝・脂質代謝・眼底)
    ・ツールで直した判定 (backup\kukuri_fix_<日時>.csv、-Since 以降)
    ・ツールで足した/直した/消した/戻した医師指示事項 (backup\kukuri_comment_*.csv、-Since 以降。空き枠を詰めた記録 _pack_ は数えない)
    ・画面で直したもの (下の $MANUAL)
  出力: out\石屋_変更一覧_<日時>.csv (Excel で開ける)。受診者の名前が入るので院内だけで使う。
  例: change_list.ps1 -Since 20261018_1130
#>
[CmdletBinding()]
param(
    [string]$Since = '20261018_1130',
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

$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$JKUKURI = [ordered]@{ '600150' = '血圧'; '600340' = '糖代謝'; '600320' = '脂質代謝'; '600060' = '眼底' }
$SLOT_NAME = @{}; foreach ($i in 1..20) { $SLOT_NAME[[string](11000 + $i)] = 1 }
# 画面で直したもの (2026/10/18 ユーザーが健診ナビの画面で直した分)
$MANUAL = @{
    '4631' = '眼底 J の文「【眼底】治療を継続してください。」を入れ、腎機能の重複した文を消した'
    '4633' = '腎機能の重複した文を消し、眼底 F の文を入れた'
    '4644' = '腎機能の重複した文を消し、胸部X線の文を入れた'
    '4646' = '総合判定(編集) を D→B「ほぼ正常(心配なし)」に戻した (先生の判断。自動判定の D は 10000 に残す)'
    '4666' = '腎機能の重複した文を消した。眼底 E の文あり'
    '4687' = '総合判定(編集) を F→D「要観察(要経過観察)」に戻した (先生の判断)。血圧の文を 1 つにまとめた'
    '4692' = '眼底 E の文「【眼底】判定不能のため、眼科で再検査を受けて下さい。」を入れた'
    '4734' = '医師指示事項の空いた行を詰めた'
    '4766' = '血液一般 D の文「【赤血球数】低値です。経過観察の必要があります。」を入れた'
    '4810' = '眼底 E の文「【眼底】判定不能のため、眼科で再検査を受けて下さい。」を入れた'
}
function Stamp-Of([string]$name) { if ($name -match '(\d{8}_\d{6})') { return $Matches[1] } else { return '' } }
function Short([string]$t) { $t = ($t -replace '\s+', ' ').Trim(); if ($t.Length -gt 40) { return $t.Substring(0, 40) + '…' } else { return $t } }

Write-Host ''
Write-Host ' 石屋 変更一覧を作る (DB は読むだけ)' -ForegroundColor Cyan
Write-Host ("  記録は {0} 以降のものを使う" -f $Since)

# ---- 記録 (backup の CSV) ----
$fixRows = @(); $comRows = @()
foreach ($f in @(Get-ChildItem -Path $BackupDir -Filter 'kukuri_fix_*.csv' -File)) {
    if ((Stamp-Of $f.Name) -lt $Since) { continue }
    $fixRows += @(Import-Csv -Path $f.FullName -Encoding UTF8)
}
foreach ($f in @(Get-ChildItem -Path $BackupDir -Filter 'kukuri_comment_*.csv' -File)) {
    if ($f.Name -like 'kukuri_comment_pack_*') { continue }
    if ((Stamp-Of $f.Name) -lt $Since) { continue }
    $kind = $(if ($f.Name -like 'kukuri_comment_restore_*') { '戻した' } elseif ($f.Name -like 'kukuri_comment_treat_*') { '治療継続' } else { '' })
    foreach ($r in @(Import-Csv -Path $f.FullName -Encoding UTF8)) { $r | Add-Member -NotePropertyName 種類 -NotePropertyValue $kind -Force; $comRows += $r }
}
Write-Host ("  判定の記録 {0} 件 / 医師指示事項の記録 {1} 件" -f $fixRows.Count, $comRows.Count)

$conn = Open-Db
try {
    $rows = (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS H, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS V
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('10000','11000','600150','600340','600320','600060')
WHERE $ISHIYA
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@).Rows
}
finally { $conn.Close() }

$People = [ordered]@{}
foreach ($r in $rows) {
    $pk = [string]$r.PK_SEQ
    if (-not $People.Contains($pk)) { $People[$pk] = @{ Uke = [string]$r.UKE; Ymd = [string]$r.YMD; Name = [string]$r.NAME; C = @{} } }
    $People[$pk].C[[string]$r.CD] = @{ H = ([string]$r.H).ToUpper(); V = [string]$r.V }
}

$out = @()
foreach ($pk in $People.Keys) {
    $p = $People[$pk]; $c = $p.C
    $edit = $(if ($c.ContainsKey('11000')) { $c['11000'] } else { @{ H = ''; V = '' } })
    $auto = $(if ($c.ContainsKey('10000')) { $c['10000'].H } else { '' })
    $why = ''
    if ($auto -ne $edit.H) {
        $eye = $(if ($c.ContainsKey('600060')) { $c['600060'].H } else { '' })
        if ($eye -eq $edit.H) { $why = "眼底 $eye のため上げた" } else { $why = '画面で直した (先生の判断)' }
    }
    $js = @($JKUKURI.Keys | Where-Object { $c.ContainsKey($_) -and $c[$_].H -eq 'J' } | ForEach-Object { $JKUKURI[$_] })
    $fx = @($fixRows | Where-Object { [string]$_.PK_SEQ -eq $pk } | ForEach-Object {
        if ($_.前 -ne $_.後) { "{0} {1}→{2}" -f $_.欄, $_.前, $_.後 } else { "{0} 文字 {1}→{2}" -f $_.欄, $(if ($_.前の結果) { $_.前の結果 } else { '空欄' }), $_.後の結果 } })
    $add = @(); $chg = @(); $del = @(); $back = @()
    foreach ($m in @($comRows | Where-Object { [string]$_.PK_SEQ -eq $pk })) {
        $b = Normalize-Text ([string]$m.前の文章); $a = Normalize-Text ([string]$m.後の文章)
        if ($m.種類 -eq '戻した') { $back += (Short $a); continue }
        if ($b -eq '' -and $a -ne '') { $add += (Short $a) }
        elseif ($b -ne '' -and $a -eq '') { $del += (Short $b) }
        else { $chg += ("{0} → {1}" -f (Short $b), (Short $a)) }
    }
    $man = $(if ($MANUAL.ContainsKey($p.Uke)) { $MANUAL[$p.Uke] } else { '' })
    $out += [PSCustomObject][ordered]@{
        受診日 = $p.Ymd; 受付番号 = $p.Uke; 氏名 = $p.Name
        結果票の総合判定 = ("{0} {1}" -f $edit.H, $edit.V).Trim()
        自動判定の総合 = $auto
        違う理由 = $why
        治療中J = ($js -join '・')
        ツールで直した判定 = ($fx -join ' / ')
        足した文 = ($add -join ' / ')
        直した文 = ($chg -join ' / ')
        消した文 = ($del -join ' / ')
        戻した手入力の文 = ($back -join ' / ')
        画面で直したもの = $man
    }
}
$file = Join-Path $OutDir ("石屋_変更一覧_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
$out | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
$changed = @($out | Where-Object { $_.違う理由 -or $_.治療中J -or $_.ツールで直した判定 -or $_.足した文 -or $_.直した文 -or $_.消した文 -or $_.戻した手入力の文 -or $_.画面で直したもの })
Write-Host ''
foreach ($d in @('2026/09/03','2026/09/09','2026/09/10')) {
    $all = @($out | Where-Object { $_.受診日 -eq $d }); $ch = @($changed | Where-Object { $_.受診日 -eq $d })
    Write-Host ("  {0}: {1} 人 (何か変わった人 {2} 人 / J のある人 {3} 人 / 総合判定が自動判定と違う人 {4} 人)" -f $d, $all.Count, $ch.Count,
        @($all | Where-Object { $_.治療中J }).Count, @($all | Where-Object { $_.違う理由 }).Count)
}
Write-Host ''
Write-Host ("[完了] {0}" -f $file) -ForegroundColor Green
Write-Host '  受診者の名前が入っています。院内だけで使い、外に出さないでください。' -ForegroundColor Yellow
try { Start-Process $file } catch { }
