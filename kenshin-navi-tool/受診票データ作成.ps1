<#
.SYNOPSIS
  受診票HTML用のデータを健診ナビから書き出す (受診票データ作成.ps1)

.DESCRIPTION
  指定した受診日の受診者について、次をJSONに書き出す。
    ・受診票に事前印字する情報 (氏名・フリガナ・事業所・受診日・受付NO・コース)
    ・前回値 (同じ人の1つ前の受診日の結果)

  健診ナビは読むだけ。一切変更しない。
  できたJSONを 健康診断受診票.html の「名簿を読込」で開く。

.EXAMPLE
  受診票データ作成.bat をダブルクリックして受診日を入れる

  powershell -ExecutionPolicy Bypass -File 受診票データ作成.ps1 -Ymd 2026/08/21
  powershell -ExecutionPolicy Bypass -File 受診票データ作成.ps1 -Ymd 2026/08/21 -Out C:\Users\User\Desktop
#>
[CmdletBinding()]
param(
    [string]$Ymd,                 # 受診日 (2026/08/21)
    [string]$Dantai,              # 団体コードでしぼる (省略可)
    [string]$Out,                 # 出力先フォルダ (既定: デスクトップ)
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

# form_import.ps1 から接続まわりを借りる (同じ接続先を使うため)
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text' '# .xlsx / .xlsm の読み込み')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Invoke-DbExec')

if (-not $Ymd) { $Ymd = Read-Host '受診日を入れてください (例 2026/08/21)' }
$Ymd = Normalize-Ymd $Ymd
if (-not $Ymd) { throw '受診日を 2026/08/21 のように入れてください。' }
if (-not $Out) { $Out = [Environment]::GetFolderPath('Desktop') }

# 受診票の欄 → 健診ナビの項目コード
#   Alt は、主の項目が空だったときに代わりに使うもの (裸眼が無ければ矯正)
$PREV = @(
    @{ Id = 'shinchou';       Cd = '001211'; Label = '身長' }
    @{ Id = 'taijuu';         Cd = '001212'; Label = '体重' }
    @{ Id = 'fukui';          Cd = '001215'; Label = '腹囲' }
    @{ Id = 'bp1_h';          Cd = '001253'; Label = '最高血圧' }
    @{ Id = 'bp1_l';          Cd = '001254'; Label = '最低血圧' }
    @{ Id = 'bp2_h';          Cd = '001255'; Label = '最高血圧2回目' }
    @{ Id = 'bp2_l';          Cd = '001256'; Label = '最低血圧2回目' }
    @{ Id = 'shiryoku_r_v';   Cd = '067012'; Alt = '067018'; Label = '視力(右)' }
    @{ Id = 'shiryoku_l_v';   Cd = '067013'; Alt = '067019'; Label = '視力(左)' }
    @{ Id = 'chouryoku_r_1k'; Cd = '067132'; Label = '聴力右1000' }
    @{ Id = 'chouryoku_r_4k'; Cd = '067133'; Label = '聴力右4000' }
    @{ Id = 'chouryoku_l_1k'; Cd = '067134'; Label = '聴力左1000' }
    @{ Id = 'chouryoku_l_4k'; Cd = '067135'; Label = '聴力左4000' }
    @{ Id = 'u_tanpaku';      Cd = '069206'; Label = '尿蛋白' }
    @{ Id = 'u_tou';          Cd = '069207'; Label = '尿糖' }
    @{ Id = 'u_senketsu';     Cd = '069211'; Label = '尿潜血' }
)
$codes = @($PREV | ForEach-Object { $_.Cd; if ($_.Alt) { $_.Alt } } | Sort-Object -Unique)
$inList = "'" + ($codes -join "','") + "'"

$conn = Open-Db
try {
    # 個人マスタの列は環境によって違うので、名前から探して使う
    $cols = @(Invoke-DbQuery $conn @"
SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KOJIN1'
"@).Rows | ForEach-Object { [string]$_.COLUMN_NAME }
    function Pick([string[]]$cands) { foreach ($c in $cands) { if ($cols -contains $c) { return $c } } return $null }
    $cBirth = Pick @('D_SEINENGAPPI', 'SEINENGAPPI', 'D_BIRTH', 'BIRTHDAY', 'D_SEIYMD', 'SEIYMD')
    $cSex   = Pick @('SEIBETU', 'SEIBETSU', 'SEX', 'F_SEIBETU')
    $selBirth = if ($cBirth) { "k.$cBirth" } else { "''" }
    $selSex   = if ($cSex)   { "k.$cSex"   } else { "''" }
    $sB = if ($cBirth) { $cBirth } else { '(見つからず)' }
    $sS = if ($cSex)   { $cSex   } else { '(見つからず)' }
    Write-Host ("[個人マスタ] 生年月日={0} / 性別={1}" -f $sB, $sS) -ForegroundColor DarkGray

    $whereDantai = if ($Dantai) { "AND s.DANTAI_CD1 = '$Dantai'" } else { '' }

    # --- 当日の受診者 ---
    $people = @(Invoke-DbQuery $conn @"
SELECT s.PK_SEQ, s.KOJIN_ID,
       ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), '') AS UKENO,
       ISNULL(k.KANJI_SIMEI,'') AS KANJI, ISNULL(k.KANA_SIMEI,'') AS KANA,
       ISNULL(d.MEISYO1,'')     AS DANTAI,
       ISNULL(c.MEISYO,'')      AS COURSE,
       ISNULL(CONVERT(varchar(20), $selBirth), '') AS BIRTH,
       ISNULL(CONVERT(varchar(10), $selSex), '')   AS SEX
FROM T_KENSIN s
LEFT JOIN T_KOJIN1  k ON k.KOJIN_ID  = s.KOJIN_ID
LEFT JOIN M_DANTAI  d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0 $whereDantai
ORDER BY k.KANA_SIMEI
"@ @{ ymd = $Ymd }).Rows
    Write-Host ("[受診者] {0} 人" -f $people.Count) -ForegroundColor Cyan
    if ($people.Count -eq 0) { throw "$Ymd の受診者が見つかりません。受診日を確かめてください。" }

    # --- 前回の受診 (同じ人の、その日より前で一番新しいもの) ---
    $prevVisit = @{}
    foreach ($r in (Invoke-DbQuery $conn @"
SELECT s.PK_SEQ AS CUR_PK, p.PK_SEQ AS PREV_PK, CONVERT(varchar(10), p.D_KENSIN, 111) AS PREV_YMD
FROM T_KENSIN s
JOIN T_KENSIN p ON p.KOJIN_ID = s.KOJIN_ID AND p.F_TORIKESI = 0 AND p.D_KENSIN < @ymd
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0 $whereDantai
  AND p.D_KENSIN = (SELECT MAX(q.D_KENSIN) FROM T_KENSIN q
                     WHERE q.KOJIN_ID = s.KOJIN_ID AND q.F_TORIKESI = 0 AND q.D_KENSIN < @ymd)
"@ @{ ymd = $Ymd }).Rows) {
        $prevVisit[[string]$r.CUR_PK] = @{ Pk = [string]$r.PREV_PK; Ymd = [string]$r.PREV_YMD }
    }
    Write-Host ("[前回あり] {0} 人" -f $prevVisit.Count) -ForegroundColor Cyan

    # --- 前回値の中身 ---
    $vals = @{}
    if ($prevVisit.Count -gt 0) {
        $pks = "'" + (($prevVisit.Values | ForEach-Object { $_.Pk }) -join "','") + "'"
        foreach ($r in (Invoke-DbQuery $conn @"
SELECT k.PK_SEQ, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, ISNULL(k.KEKKA,'') AS V
FROM T_KENSA k
WHERE k.PK_SEQ IN ($pks) AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ($inList)
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('', '#')
"@).Rows) {
            $key = '{0}/{1}' -f $r.PK_SEQ, $r.CD
            $vals[$key] = (Normalize-Text ([string]$r.V))
        }
    }

    # --- 組み立て ---
    $list = @()
    foreach ($p in $people) {
        $inputs = [ordered]@{
            jushinbi     = $Ymd
            uketsuke_no  = (Normalize-Text ([string]$p.UKENO))
            course       = (Normalize-Text ([string]$p.COURSE))
            jigyousho    = (Normalize-Text ([string]$p.DANTAI))
            furigana     = (Normalize-Text ([string]$p.KANA))
            shimei       = (Normalize-Text ([string]$p.KANJI))
            seinengappi  = (Normalize-Text ([string]$p.BIRTH))
            seibetsu     = (Normalize-Text ([string]$p.SEX))
        }
        $prev = [ordered]@{}
        $pv = $prevVisit[[string]$p.PK_SEQ]
        if ($pv) {
            foreach ($m in $PREV) {
                $v = $vals['{0}/{1}' -f $pv.Pk, $m.Cd]
                if ((-not $v) -and $m.Alt) { $v = $vals['{0}/{1}' -f $pv.Pk, $m.Alt] }
                if ($v) { $prev[$m.Id] = $v }
            }
        }
        $list += [ordered]@{
            label     = ('{0}  {1}' -f (Normalize-Text ([string]$p.UKENO)), (Normalize-Text ([string]$p.KANJI))).Trim()
            inputs    = $inputs
            prev      = $prev
            prev_ymd  = $(if ($pv) { $pv.Ymd } else { '' })
        }
    }

    $doc = [ordered]@{
        kind     = 'kenshin_jushinhyo_roster'
        ymd      = $Ymd
        made     = (Get-Date -Format 'yyyy/MM/dd HH:mm')
        people   = $list
    }
    $name = '受診票データ_' + ($Ymd -replace '/', '') + '.json'
    $path = Join-Path $Out $name
    # BOM無しUTF-8で書く (BOMがあるとブラウザの JSON.parse が失敗する)
    [System.IO.File]::WriteAllText($path, ($doc | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))

    $withPrev = @($list | Where-Object { $_.prev.Count -gt 0 }).Count
    Write-Host ''
    Write-Host ("できました: {0}" -f $path) -ForegroundColor Green
    Write-Host ("  {0} 人 / うち前回値あり {1} 人" -f $list.Count, $withPrev)
    Write-Host ''
    Write-Host '健康診断受診票.html を開いて「名簿を読込」でこのファイルを選んでください。'
}
finally { $conn.Close() }
