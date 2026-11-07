<#
.SYNOPSIS
  新島の受診予定者を Excel (.xlsx) に書き出す (niijima_list.ps1)

.DESCRIPTION
  健診ナビの予約のうち、事業所名に「新島」「式根島」が入る事業所で、指定した期間に受診する人 (取消を除く) を
  1 人 1 行で Excel に書き出す。
  - データベースは読むだけ (SELECT のみ)。予約・個人情報の追加・変更・削除はしない
  - 出力先は excel_tools\out だけ。メールやクラウドには送らない (要配慮個人情報)
  - 番号の頭の 0 が消えないよう、全部の欄を文字として書く (Excel が無くても作れる)
  列: 予約ID / 個人ID / 受診日 / 事業所コード / 事業所名 / 氏名 / フリガナ / フリガナ(全角) / 性別 / 生年月日 / 年齢 /
      コースコード / コース名 / オプション / 受付番号 / 住民番号等
  例:
    niijima_list.ps1                                  (2026/11/01〜2026/11/30)
    niijima_list.ps1 -From 2026/11/10 -To 2026/11/12
#>
[CmdletBinding()]
param(
    [string]$From = '2026/11/01',
    [string]$To = '2026/11/30',
    [string]$Words = '新島,式根島',
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'
trap { Write-Host (" エラー: {0}`n   {1} 行目: {2}" -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber, ([string]$_.InvocationInfo.Line).Trim()) -ForegroundColor Red; break }
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$a, [string]$b) { $i = $src.IndexOf($a); $j = $src.IndexOf($b, $i); if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($a)" }; return $src.Substring($i, $j - $i) }
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName Microsoft.VisualBasic

$Inv = [Globalization.CultureInfo]::InvariantCulture
foreach ($d in @($From, $To)) { if ($d -notmatch '^\d{4}/\d{1,2}/\d{1,2}$') { throw "日付は 2026/11/01 の形で: $d" } }
$wl = @(($Words -split '[,、\s]+') | Where-Object { $_ -ne '' })
if ($wl.Count -eq 0) { throw '事業所名で探す言葉を入れてください' }
$likes = ($wl | ForEach-Object { "d.MEISYO1 LIKE N'%" + ($_ -replace "'", "''") + "%'" }) -join ' OR '
$WHO = "s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN @f AND @t AND s.DANTAI_CD1 IN (SELECT d.DANTAI_CD1 FROM M_DANTAI d WHERE $likes)"
$P = @{ f = $From; t = $To }

Write-Host ''
Write-Host (" 新島の受診予定者を Excel に書き出す ({0}〜{1}、事業所名に {2})。データベースは読むだけ" -f $From, $To, ($wl -join '・')) -ForegroundColor Cyan
$conn = Open-Db
try {
    $rows = @((Invoke-DbQuery $conn @"
SELECT s.PK_SEQ AS PK, s.KOJIN_ID AS KID, s.D_KENSIN AS KDATE,
       LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS DCD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS DNAME,
       LTRIM(RTRIM(ISNULL(k.KANJI_SIMEI,''))) AS KANJI, LTRIM(RTRIM(ISNULL(k.KANA_SIMEI,''))) AS KANA,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(5), k.SEIBETU),''))) AS SEX, k.D_SEINEN AS BIRTH,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS CCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS CNAME,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),''))) AS UKE,
       LTRIM(RTRIM(ISNULL(k.KOJIN_NO,''))) AS KNO,
       STUFF((SELECT N'・' + LTRIM(RTRIM(r.KOMOKU_CD)) + N' ' + LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))
              FROM T_RYOUKIN r LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
              WHERE r.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) LIKE 'OP%'
              GROUP BY r.KOMOKU_CD, m.MEISYO1 FOR XML PATH('')), 1, 1, '') AS OPTS
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE $WHO
ORDER BY s.D_KENSIN, s.DANTAI_CD1, k.KANA_SIMEI
"@ $P).Rows)
} finally { $conn.Close() }

if ($rows.Count -eq 0) { Write-Host ' 該当する受診予定者がいません。期間・事業所名を確かめてください。何も書き出していません。' -ForegroundColor Yellow; return }

function Ymd($v) {
    if ($null -eq $v -or $v -is [DBNull]) { return '' }
    if ($v -is [datetime]) { return $v.ToString('yyyy/MM/dd') }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse([string]$v, [ref]$dt)) { return $dt.ToString('yyyy/MM/dd') }
    return ([string]$v).Trim()
}
$HDR = @('予約ID','個人ID','受診日','事業所コード','事業所名','氏名','フリガナ','フリガナ(全角)','性別','生年月日','年齢','コースコード','コース名','オプション','受付番号','住民番号等')
$data = New-Object System.Collections.Generic.List[object]
foreach ($r in $rows) {
    $kd = Ymd $r.KDATE; $bd = Ymd $r.BIRTH; $age = ''
    $kdt = [datetime]::MinValue; $bdt = [datetime]::MinValue
    if ([datetime]::TryParse($kd, [ref]$kdt) -and [datetime]::TryParse($bd, [ref]$bdt)) {
        $a = $kdt.Year - $bdt.Year; if ($kdt.Month -lt $bdt.Month -or ($kdt.Month -eq $bdt.Month -and $kdt.Day -lt $bdt.Day)) { $a-- }
        $age = [string]$a
    }
    $kana = [string]$r.KANA
    $wide = $(if ($kana) { [Microsoft.VisualBasic.Strings]::StrConv($kana, [Microsoft.VisualBasic.VbStrConv]::Wide) } else { '' })
    $sex = switch ([string]$r.SEX) { '1' { '男' } '2' { '女' } default { '' } }
    $uke = $(if ([string]$r.UKE -eq '0') { '' } else { [string]$r.UKE })
    $data.Add(@([string]$r.PK, [string]$r.KID, $kd, [string]$r.DCD, [string]$r.DNAME, [string]$r.KANJI, $kana, $wide, $sex, $bd, $age,
                [string]$r.CCD, [string]$r.CNAME, $(if ($r.OPTS -is [DBNull]) { '' } else { [string]$r.OPTS }), $uke, [string]$r.KNO))
}

# ---- .xlsx を作る (全部の欄を文字で書く) ----
function XEsc([string]$s) {
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}
function ColName([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = ([string][char](65 + $m)) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; return $s }
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
[void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>')
# List どうしを + でつなげると PowerShell 5.1 でエラーになるので、1 行ずつ足す
$all = New-Object System.Collections.Generic.List[object]
$all.Add($HDR)
foreach ($x in $data) { $all.Add($x) }
for ($ri = 0; $ri -lt $all.Count; $ri++) {
    [void]$sb.Append("<row r=`"$($ri + 1)`">")
    $vals = $all[$ri]
    for ($ci = 0; $ci -lt $vals.Count; $ci++) {
        $ref = (ColName ($ci + 1)) + ($ri + 1)
        $st = $(if ($ri -eq 0) { ' s="1"' } else { '' })
        [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"$st><is><t xml:space=`"preserve`">$(XEsc ([string]$vals[$ci]))</t></is></c>")
    }
    [void]$sb.Append('</row>')
}
$lastRef = (ColName $HDR.Count) + $all.Count
[void]$sb.Append("</sheetData><autoFilter ref=`"A1:$lastRef`"/></worksheet>")

$parts = [ordered]@{
    '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="受診予定者" sheetId="1" r:id="rId1"/></sheets></workbook>'
    'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    'xl/styles.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font><font><b/><sz val="11"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="49" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    'xl/worksheets/sheet1.xml' = $sb.ToString()
}
$OutDir = Join-Path $PSScriptRoot 'out'
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
$file = Join-Path $OutDir ("新島_受診予定者_{0}-{1}_{2}.xlsx" -f ($From -replace '/', ''), ($To -replace '/', ''), (Get-Date -Format 'yyyyMMdd_HHmmss'))
$fs = [System.IO.File]::Create($file)
try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) {
            $e = $zip.CreateEntry($name)
            $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding $false))
            try { $w.Write($parts[$name]) } finally { $w.Dispose() }
        }
    } finally { $zip.Dispose() }
} finally { $fs.Dispose() }

# ---- 画面には人数だけ出す (氏名は出さない) ----
Write-Host (" {0} 人を書き出しました: {1}" -f $data.Count, $file) -ForegroundColor Green
Write-Host ''
Write-Host ' 受診日ごとの人数:'
$data | Group-Object { $_[2] } | Sort-Object Name | ForEach-Object { Write-Host ("   {0}  {1} 人" -f $_.Name, $_.Count) }
$noKana = @($data | Where-Object { $_[6] -eq '' }).Count
$noBirth = @($data | Where-Object { $_[9] -eq '' }).Count
$dup = @($data | Group-Object { $_[1] } | Where-Object { $_.Count -gt 1 }).Count
if ($noKana) { Write-Host ("   フリガナが空の人: {0} 人" -f $noKana) -ForegroundColor Yellow }
if ($noBirth) { Write-Host ("   生年月日が空の人: {0} 人" -f $noBirth) -ForegroundColor Yellow }
if ($dup) { Write-Host ("   予約が 2 つ以上ある人: {0} 人 (同じ個人ID が 2 行以上)" -f $dup) -ForegroundColor Yellow }
Write-Host ''
Write-Host ' ★このファイルは要配慮個人情報です。院内の PC の中だけで使い、メール・クラウドに送らないでください。使い終わったら消してください。' -ForegroundColor Yellow
Start-Process $OutDir
