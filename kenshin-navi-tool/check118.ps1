<#
  新島: 受付アプリに事前登録するための下調べ (check118.ps1)
  健診ナビの受診予定者 (予約) のうち 新島・式根島 の事業所の分について、件数と「項目がどれくらい埋まっているか」だけを出す。
  ★氏名・フリガナ・生年月日などの中身は一切出さない (人数と列名だけ)。そのまま貼ってよい。
  DBは読むだけ (SELECT のみ)。データの追加・変更・削除はしない。
  check118.bat → r_check118.txt
  例: check118.ps1 -From 2026/11/01 -To 2027/03/31
#>
[CmdletBinding()]
param(
    [string]$From,
    [string]$To,
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

$Inv = [Globalization.CultureInfo]::InvariantCulture
if (-not $From) { $From = (Get-Date).ToString('yyyy/MM/dd') }
if (-not $To)   { $To = (Get-Date).AddMonths(6).ToString('yyyy/MM/dd') }
foreach ($d in @($From, $To)) { if ($d -notmatch '^\d{4}/\d{1,2}/\d{1,2}$') { throw "日付は 2026/11/01 の形で: $d" } }
$wl = @(($Words -split '[,、\s]+') | Where-Object { $_ -ne '' })
if ($wl.Count -eq 0) { throw '事業所名で探す言葉を入れてください' }
$likes = ($wl | ForEach-Object { "d.MEISYO1 LIKE N'%" + ($_ -replace "'", "''") + "%'" }) -join ' OR '
$DANTAI = "(SELECT d.DANTAI_CD1 FROM M_DANTAI d WHERE $likes)"
$WHO = "s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN @f AND @t AND s.DANTAI_CD1 IN $DANTAI"
$P = @{ f = $From; t = $To }

$out = Join-Path $PSScriptRoot 'r_check118.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function T($dt) { W (($dt | Format-Table -AutoSize | Out-String -Width 250).TrimEnd()) }
"=== 新島: 受付アプリ事前登録の下調べ ($From〜$To、事業所名に $($wl -join '・')) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W '  ★人数と列名だけ。氏名などの中身は出していない。DBは読むだけ。'

$conn = Open-Db
try {
    W ''; W '--- 1. 対象の事業所 (事業所名に 新島・式根島 が入るもの) と、期間内の予約人数 ---'
    T (Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(d.DANTAI_CD1)) AS 事業所CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所名,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), d.DANTAIGROUP_CD),''))) AS グループCD,
       (SELECT COUNT(*) FROM T_KENSIN s WHERE s.DANTAI_CD1 = d.DANTAI_CD1 AND s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN @f AND @t) AS 期間内の予約
FROM M_DANTAI d WHERE $likes
ORDER BY 4 DESC, 1
"@ $P)

    W ''; W '--- 2. ★受診日 × コース ごとの人数 (取消を除く) ---'
    T (Invoke-DbQuery $conn @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       MAX(LTRIM(RTRIM(ISNULL(c.MEISYO,'')))) AS コース名, COUNT(*) AS 人数
FROM T_KENSIN s
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE $WHO
GROUP BY s.D_KENSIN, s.COURSE_CD ORDER BY 1, 2
"@ $P)

    W ''; W '--- 3. 取消になっている予約の数 (参考。事前登録には入れない) ---'
    T (Invoke-DbQuery $conn @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 取消
FROM T_KENSIN s WHERE s.F_TORIKESI <> 0 AND s.D_KENSIN BETWEEN @f AND @t AND s.DANTAI_CD1 IN $DANTAI
GROUP BY s.D_KENSIN ORDER BY 1
"@ $P)

    W ''; W '--- 4. ★受付アプリに渡す項目が、何人分埋まっているか (中身は出さない) ---'
    T (Invoke-DbQuery $conn @"
SELECT COUNT(*) AS 予約人数,
  SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KANJI_SIMEI,''))) <> '' THEN 1 ELSE 0 END) AS 氏名あり,
  SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KANA_SIMEI,''))) <> '' THEN 1 ELSE 0 END) AS フリガナあり,
  SUM(CASE WHEN k.KANA_SIMEI COLLATE Japanese_BIN LIKE N'%[ｦ-ﾟ]%' THEN 1 ELSE 0 END) AS うち半角カナ,   -- 全角・半角を区別するため BIN で比べる
  SUM(CASE WHEN k.D_SEINEN IS NOT NULL AND LTRIM(RTRIM(CONVERT(varchar(20), k.D_SEINEN))) <> '' THEN 1 ELSE 0 END) AS 生年月日あり,
  SUM(CASE WHEN LTRIM(RTRIM(ISNULL(CONVERT(varchar(5), k.SEIBETU),''))) IN ('1','2') THEN 1 ELSE 0 END) AS 性別あり,
  SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) <> '' THEN 1 ELSE 0 END) AS コースあり,
  SUM(CASE WHEN LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),''))) NOT IN ('','0') THEN 1 ELSE 0 END) AS 受付番号あり,
  SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KOJIN_NO,''))) <> '' THEN 1 ELSE 0 END) AS 個人番号あり,
  COUNT(DISTINCT s.KOJIN_ID) AS 別人の数
FROM T_KENSIN s LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
WHERE $WHO
"@ $P)
    W '  (別人の数 が 予約人数 より少なければ、同じ人に予約が 2 つ以上ある)'

    W ''; W '--- 5. 同じ人に予約が 2 つ以上ある数 (受付アプリで二重登録にならないよう確かめる) ---'
    T (Invoke-DbQuery $conn @"
SELECT COUNT(*) AS 予約が2つ以上の人 FROM (
  SELECT s.KOJIN_ID FROM T_KENSIN s WHERE $WHO GROUP BY s.KOJIN_ID HAVING COUNT(*) > 1) x
"@ $P)

    W ''; W '--- 6. オプション (料金明細の OP〜) の人数 ---'
    T (Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS オプション, COUNT(DISTINCT r.PK_SEQ) AS 人数
FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
WHERE $WHO AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) LIKE 'OP%'
GROUP BY LTRIM(RTRIM(r.KOMOKU_CD)) ORDER BY 3 DESC
"@ $P)

    # 7. 予約の表 (T_KENSIN) の列: 何人分埋まっているかだけ。時間帯・順番の列を探すため
    W ''; W '--- 7. 予約の表 (T_KENSIN) の列ごとに、何人分値が入っているか (時間帯・順番などの列を探すため。中身は出さない) ---'
    $cols = @((Invoke-DbQuery $conn "SELECT COLUMN_NAME AS C, DATA_TYPE AS T FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KENSIN' ORDER BY ORDINAL_POSITION").Rows |
        Where-Object { [string]$_.T -notin @('image','varbinary','binary','timestamp','xml','sql_variant') })
    $parts = foreach ($c in $cols) {
        $n = [string]$c.C
        if ($n -notmatch '^[A-Za-z0-9_]+$') { continue }
        "SUM(CASE WHEN s.[$n] IS NOT NULL AND LTRIM(RTRIM(CONVERT(nvarchar(100), s.[$n]))) NOT IN ('', '0') THEN 1 ELSE 0 END) AS [$n]"
    }
    $row = (Invoke-DbQuery $conn ("SELECT COUNT(*) AS [__ALL], " + ($parts -join ', ') + " FROM T_KENSIN s WHERE $WHO") $P).Rows[0]
    $all = [int]$row['__ALL']
    $list = foreach ($c in $cols) { $n = [string]$c.C; if ($n -match '^[A-Za-z0-9_]+$') { [PSCustomObject]@{ 列 = $n; 型 = [string]$c.T; 値あり = [int]$row[$n]; 全体 = $all } } }
    T $list
    W '  (JIKAN・TIME・JUN・WAKU などの名前で 値あり が多い列が、時間帯・順番の候補。次の段階で中身の形だけ確かめる)'
} finally { $conn.Close() }

W ''
W '=== 読み方 ==='
W '  1 で、新島の事業所が全部入っているか (足りない・余分な会社がないか) を見る。'
W '  2 が受付アプリに入れる人数の内訳。4 で、受付アプリに渡す項目がそろっているか (フリガナ・生年月日が空の人がいないか) を見る。'
W '  ここまでで実際の一覧は作っていない。見本の形が決まったら、一覧を出す道具を作る。'
notepad $out
