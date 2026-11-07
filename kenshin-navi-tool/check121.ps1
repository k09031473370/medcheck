<#
  新島: 当日追加オプションの下調べ その 3 (check121.ps1)
  check120 で時間切れになった所と、新しく出た疑問を確かめる。マスタと人数・金額だけ。氏名は出さない。
  DBは読むだけ (SELECT のみ)。時間がかかる問い合わせがあるので、待ち時間を 15 分にしている。
    1. 画面でオプションを付けた予約に、付けていない予約には無い検査の行が何できているか (孫の行・くくりの行も含めて)
    2. 検体依頼のコード変換 (M_HENKAN_IRAI) にオプションの検査が載っているか (check120 の 2 のやり直し)
    3. オプションの料金明細の金額が、団体のオプション定義 (M_OPTION) どおりか
    4. 料金明細の連番 99 は何か (合計の行か)
    5. 新島の 11 月で料金明細が無い 1 件
    6. オプションの検査の行に入っている列 (親の行・中身の行・ほかの行)
  check121.bat → r_check121.txt
#>
[CmdletBinding()]
param(
    [string]$From = '2026/11/01', [string]$To = '2026/11/30', [string]$Words = '新島,式根島',
    [string]$P1 = '2025/11/01', [string]$P2 = '2026/11/06',
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
$out  = Join-Path $PSScriptRoot 'r_check121.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
"=== 新島: 当日追加オプションの下調べ その 3 (新島 $From〜$To / 比べる期間 $P1〜$P2) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W '  ★マスタと人数・金額だけ。氏名は出していない。DBは読むだけ。'
if (-not (Test-Path $Core)) { W "form_import.ps1 がありません: $PSScriptRoot"; notepad.exe $out; return }
$src = Get-Content $Core -Raw
function Get-Part([string]$a, [string]$b) {
    $i = $src.IndexOf($a); $j = $src.IndexOf($b, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($a)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Read-TextShared')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Resolve-PkSeq')

function Query-Long($conn, [string]$sql) {
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $sql; $cmd.CommandTimeout = 900
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt); $cmd.Dispose()
    return ,$dt
}
function Show($dt) {
    if ($dt.Rows.Count -eq 0) { W '(全 0 行)'; return }
    ($dt | Format-Table -AutoSize -Wrap | Out-String -Width 400).TrimEnd() | Out-File $out -Append -Encoding Default
    W "(全 $($dt.Rows.Count) 行)"
}
function Sec($title, $conn, [string]$sql) {
    W ''; W $title
    Write-Host " $title" -ForegroundColor Cyan
    try { $dt = Query-Long $conn $sql; Show $dt; return ,$dt }
    catch { W ("  エラー: " + $_.Exception.Message); return $null }
}

$wl = @(($Words -split '[,、\s]+') | Where-Object { $_ -ne '' })
$likes = ($wl | ForEach-Object { "d.MEISYO1 LIKE N'%" + ($_ -replace "'", "''") + "%'" }) -join ' OR '
$WHO = "s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND s.DANTAI_CD1 IN (SELECT d.DANTAI_CD1 FROM M_DANTAI d WHERE $likes)"
$OPS = "('OP0001','OP0002','OP0003','OP0004','OP0005','OP0006','OP0007','OP0008','OP0009')"
# 画面でオプションを付けた予約 (比べる期間)
$O = @"
SELECT DISTINCT s.PK_SEQ, s.DANTAI_CD1, s.COURSE_CD, LTRIM(RTRIM(r.KOMOKU_CD)) AS OP
FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$P1' AND '$P2' AND r.KOMOKU_CD IN $OPS
"@

Write-Host ''
Write-Host ' 新島の当日オプションの下調べ その 3 (DBは読むだけ。数分かかることがあります)' -ForegroundColor Cyan
$conn = Open-Db
try {
    # ---- 1 ----
    $fp = Sec '--- 1. ★オプションを付けた予約にだけある検査の行 (同じ団体・コースで付けていない予約と比べる) ---' $conn @"
WITH o AS ($O),
g AS (SELECT DISTINCT OP, DANTAI_CD1, COURSE_CD FROM o),
b AS (SELECT DISTINCT g.OP, s.PK_SEQ FROM g JOIN T_KENSIN s ON s.DANTAI_CD1 = g.DANTAI_CD1 AND s.COURSE_CD = g.COURSE_CD
      WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$P1' AND '$P2'
        AND NOT EXISTS (SELECT 1 FROM o x WHERE x.PK_SEQ = s.PK_SEQ AND x.OP = g.OP)),
na AS (SELECT OP, COUNT(DISTINCT PK_SEQ) AS N FROM o GROUP BY OP),
nb AS (SELECT OP, COUNT(DISTINCT PK_SEQ) AS N FROM b GROUP BY OP),
ia AS (SELECT o.OP, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,''))) AS OYA,
              LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.F_KOMOKU),''))) AS FK, COUNT(DISTINCT k.PK_SEQ) AS N
       FROM o JOIN T_KENSA k ON k.PK_SEQ = o.PK_SEQ
       GROUP BY o.OP, LTRIM(RTRIM(k.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,''))), LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.F_KOMOKU),'')))),
ib AS (SELECT b.OP, LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, COUNT(DISTINCT k.PK_SEQ) AS N
       FROM b JOIN T_KENSA k ON k.PK_SEQ = b.PK_SEQ GROUP BY b.OP, LTRIM(RTRIM(k.KOMOKU_CD)))
SELECT ia.OP AS オプション, ia.OYA AS 親, ia.CD AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 項目名,
       MAX(ISNULL(CONVERT(varchar(10), m.K_SET),'')) AS セット, MAX(LTRIM(RTRIM(ISNULL(m.F_IRAI,'')))) AS 依頼の印, ia.FK AS F_KOMOKU,
       ia.N AS 付けた予約で, MAX(na.N) AS 付けた予約数, MAX(ISNULL(ib.N,0)) AS 付けていない予約で, MAX(ISNULL(nb.N,0)) AS 付けていない予約数
FROM ia JOIN na ON na.OP = ia.OP
LEFT JOIN ib ON ib.OP = ia.OP AND ib.CD = ia.CD
LEFT JOIN nb ON nb.OP = ia.OP
LEFT JOIN T_KOMOKU m ON m.KOMOKU_CD = ia.CD
WHERE ia.N * 2 >= na.N AND (nb.N IS NULL OR ISNULL(ib.N,0) * na.N * 2 < ia.N * nb.N)
GROUP BY ia.OP, ia.OYA, ia.CD, ia.FK, ia.N
ORDER BY 1, 2, 3
"@
    W '  読み方: 付けた予約の半分以上にあり、付けていない予約ではずっと少ない行 = 画面がオプションで作る行。親が空の行は、親を持たずに増える行 (くくりなど)。'

    # ---- 2 ----
    $codes = New-Object System.Collections.Generic.List[string]
    foreach ($c in @('OP0001','OP0002','OP0003','OP0004','OP0005','OP0006','OP0007','OP0008','OP0009',
                     '067818','068690','067812','067823','067802','067804','067809','067841','068205','068206','068211','168611','069245','069246')) { $codes.Add($c) }
    for ($n = 68360; $n -le 68399; $n++) { $codes.Add('0' + $n) }
    if ($fp) { foreach ($r in $fp.Rows) { $cd = [string]$r['項目CD']; if ($cd -match '^[0-9A-Za-z-]+$' -and -not $codes.Contains($cd)) { $codes.Add($cd) } } }
    $IN = "('" + ($codes -join "','") + "')"
    $null = Sec '--- 2. ★検体依頼のコード変換 (M_HENKAN_IRAI) にオプションの検査が載っているか (センターごと) ---' $conn @"
SELECT LTRIM(RTRIM(h.KOMOKU_CD)) AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 項目名, LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), h.CENTER_CD),''))) AS センター,
       LTRIM(RTRIM(ISNULL(h.OUT_CD,''))) AS 依頼コード, LTRIM(RTRIM(ISNULL(h.OUT_CD_OYA,''))) AS 依頼コード親,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), h.KENTAI_KBN),''))) AS 検体区分, LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), h.YOUKI_CD),''))) AS 容器CD
FROM M_HENKAN_IRAI h LEFT JOIN T_KOMOKU m ON m.KOMOKU_CD = h.KOMOKU_CD
WHERE h.KOMOKU_CD IN $IN
GROUP BY h.KOMOKU_CD, h.CENTER_CD, h.OUT_CD, h.OUT_CD_OYA, h.KENTAI_KBN, h.YOUKI_CD ORDER BY 1, 3
"@
    $null = Sec '--- 2b. 上の検査のうち、コード変換が 1 つも無いもの ---' $conn @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名, ISNULL(CONVERT(varchar(10), m.K_SET),'') AS セット, LTRIM(RTRIM(ISNULL(m.F_IRAI,''))) AS 依頼の印
FROM T_KOMOKU m WHERE m.KOMOKU_CD IN $IN AND NOT EXISTS (SELECT 1 FROM M_HENKAN_IRAI h WHERE h.KOMOKU_CD = m.KOMOKU_CD)
ORDER BY 1
"@

    # ---- 3 ----
    $null = Sec '--- 3. ★オプションの料金明細の金額は、団体のオプション定義 (M_OPTION) どおりか (比べる期間・全団体) ---' $conn @"
SELECT x.OP AS オプション, x.結果, COUNT(*) AS 行数, MIN(x.GOUKEI) AS 合計最小, MAX(x.GOUKEI) AS 合計最大
FROM (SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS OP, r.GOUKEI,
             CASE WHEN o.KOMOKU_CD IS NULL THEN N'団体・コースに定義なし'
                  WHEN ISNULL(r.DANTAI_RYOUKIN,0) = ISNULL(o.DANTAI_RYOUKIN,0) AND ISNULL(r.KOJIN_RYOUKIN,0) = ISNULL(o.KOJIN_RYOUKIN,0) THEN N'定義どおり'
                  WHEN ISNULL(r.GOUKEI,0) = 0 THEN N'0円 (定義は有料)'
                  ELSE N'定義と違う' END AS 結果
      FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
      LEFT JOIN M_OPTION o ON o.DANTAI_CD1 = s.DANTAI_CD1 AND o.COURSE_CD = s.COURSE_CD AND o.KOMOKU_CD = r.KOMOKU_CD
      WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$P1' AND '$P2' AND r.KOMOKU_CD IN $OPS) x
GROUP BY x.OP, x.結果 ORDER BY 1, 2
"@
    $null = Sec '--- 3b. 定義どおりでない行の団体 (団体ごと) ---' $conn @"
SELECT LTRIM(RTRIM(s.DANTAI_CD1)) AS 団体CD, MAX(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 団体名, LTRIM(RTRIM(r.KOMOKU_CD)) AS オプション,
       COUNT(*) AS 行数, MAX(r.GOUKEI) AS 明細の合計, MAX(ISNULL(o.DANTAI_RYOUKIN,-1)) AS 定義の団体料金, MAX(ISNULL(o.KOJIN_RYOUKIN,-1)) AS 定義の個人料金
FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN M_OPTION o ON o.DANTAI_CD1 = s.DANTAI_CD1 AND o.COURSE_CD = s.COURSE_CD AND o.KOMOKU_CD = r.KOMOKU_CD
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$P1' AND '$P2' AND r.KOMOKU_CD IN $OPS
  AND NOT (o.KOMOKU_CD IS NOT NULL AND ISNULL(r.DANTAI_RYOUKIN,0) = ISNULL(o.DANTAI_RYOUKIN,0) AND ISNULL(r.KOJIN_RYOUKIN,0) = ISNULL(o.KOJIN_RYOUKIN,0))
GROUP BY s.DANTAI_CD1, r.KOMOKU_CD ORDER BY 4 DESC
"@

    # ---- 4 ----
    $null = Sec '--- 4. ★新島の 11 月: 連番 1 と連番 99 の金額の列 (合計) ---' $conn @"
SELECT r.REN AS 連番, COUNT(*) AS 行数, SUM(ISNULL(r.KENPO_RYOUKIN,0)) AS 健保, SUM(ISNULL(r.DANTAI_RYOUKIN,0)) AS 団体, SUM(ISNULL(r.KOJIN_RYOUKIN,0)) AS 個人,
       SUM(ISNULL(r.DAIKOU_RYOUKIN,0)) AS 代行, SUM(ISNULL(r.SONOTA_RYOUKIN,0)) AS その他, SUM(ISNULL(r.GOUKEI,0)) AS 合計,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(r.TEKIYOU,''))) = '' THEN 0 ELSE 1 END) AS 摘要あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(r.DENPYOU_BANGO,''))) = '' THEN 0 ELSE 1 END) AS 伝票あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(r.D_RYOUSYUU,''))) = '' THEN 0 ELSE 1 END) AS 領収日あり
FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ WHERE $WHO
GROUP BY r.REN ORDER BY 1
"@
    $null = Sec '--- 4b. ★連番 99 の合計と、ほかの連番の合計の関係 (比べる期間。オプションのある予約・無い予約) ---' $conn @"
SELECT x.種類, CASE WHEN x.R99 IS NULL THEN N'連番 99 が無い'
                   WHEN x.R99 = x.RSUM THEN N'99 = ほかの連番の和'
                   WHEN x.R99 = x.R1 THEN N'99 = 連番 1 (オプションを足していない)'
                   ELSE N'どれとも違う' END AS 結果, COUNT(*) AS 予約数
FROM (SELECT s.PK_SEQ,
             CASE WHEN EXISTS (SELECT 1 FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ AND q.KOMOKU_CD IN $OPS) THEN N'新島のオプションあり'
                  WHEN EXISTS (SELECT 1 FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ AND q.REN NOT IN (1, 99)) THEN N'ほかのオプションあり'
                  ELSE N'オプション無し' END AS 種類,
             (SELECT SUM(ISNULL(q.GOUKEI,0)) FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ AND q.REN = 99) AS R99,
             (SELECT SUM(ISNULL(q.GOUKEI,0)) FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ AND q.REN = 1) AS R1,
             (SELECT SUM(ISNULL(q.GOUKEI,0)) FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ AND q.REN <> 99) AS RSUM
      FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$P1' AND '$P2'
        AND EXISTS (SELECT 1 FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ)) x
GROUP BY x.種類, CASE WHEN x.R99 IS NULL THEN N'連番 99 が無い'
                   WHEN x.R99 = x.RSUM THEN N'99 = ほかの連番の和'
                   WHEN x.R99 = x.R1 THEN N'99 = 連番 1 (オプションを足していない)'
                   ELSE N'どれとも違う' END
ORDER BY 1, 2
"@
    $null = Sec '--- 4c. オプションのある予約の連番の並び (比べる期間) ---' $conn @"
SELECT r.REN AS 連番, CASE WHEN LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) = '' THEN '(空)' WHEN r.KOMOKU_CD LIKE 'OP%' THEN 'OP…' ELSE N'そのほかの項目' END AS 項目CD,
       ISNULL(CONVERT(varchar(10), r.F_ADD_DEL),'') AS 追加削除, COUNT(*) AS 行数
FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$P1' AND '$P2'
  AND EXISTS (SELECT 1 FROM T_RYOUKIN q WHERE q.PK_SEQ = s.PK_SEQ AND q.KOMOKU_CD IN $OPS)
GROUP BY r.REN, CASE WHEN LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) = '' THEN '(空)' WHEN r.KOMOKU_CD LIKE 'OP%' THEN 'OP…' ELSE N'そのほかの項目' END, r.F_ADD_DEL
ORDER BY 1, 2
"@

    # ---- 5 ----
    $null = Sec '--- 5. 新島の 11 月で料金明細が無い予約 (コース・検査の行の数だけ) ---' $conn @"
SELECT s.D_KENSIN AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD, (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS 検査の行
FROM T_KENSIN s WHERE $WHO AND NOT EXISTS (SELECT 1 FROM T_RYOUKIN r WHERE r.PK_SEQ = s.PK_SEQ)
"@

    # ---- 6 ----
    W ''; W '--- 6. 検査の行 (T_KENSA) の列の入り方: 新島のオプションの親の行・中身の行・ほかの行 (比べる期間。値は印の列だけ) ---'
    Write-Host ' --- 6. 検査の行の列の入り方 ---' -ForegroundColor Cyan
    try {
        $cols = @((Query-Long $conn "SELECT c.COLUMN_NAME AS C FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_KENSA' ORDER BY c.ORDINAL_POSITION").Rows | ForEach-Object { [string]$_.C })
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($c in $cols) { $parts.Add("SUM(CASE WHEN k.[$c] IS NULL THEN 0 WHEN LTRIM(RTRIM(CONVERT(nvarchar(200), k.[$c]))) = '' THEN 0 ELSE 1 END) AS [$c]") }
        $kind = "CASE WHEN k.KOMOKU_CD IN $OPS THEN N'親' WHEN k.OYA_KOMOKU_CD IN $OPS THEN N'中身' ELSE N'ほか' END"
        $fill = Query-Long $conn @"
SELECT $kind AS 種類, COUNT(*) AS 行数, $($parts -join ', ')
FROM T_KENSA k WHERE k.PK_SEQ IN (SELECT PK_SEQ FROM ($O) o)
GROUP BY $kind
"@
        $flagCols = @($cols | Where-Object { $_ -match '^(F_|K_|D_|T_)' -or $_ -match 'KBN|FLG' })
        $vals = @{}
        foreach ($c in $flagCols) {
            $v = Query-Long $conn @"
SELECT $kind AS 種類, LTRIM(RTRIM(ISNULL(CONVERT(nvarchar(40), k.[$c]), N'(空)'))) AS V, COUNT(*) AS N
FROM T_KENSA k WHERE k.PK_SEQ IN (SELECT PK_SEQ FROM ($O) o) AND ($kind) <> N'ほか'
GROUP BY $kind, LTRIM(RTRIM(ISNULL(CONVERT(nvarchar(40), k.[$c]), N'(空)')))
"@
            foreach ($r in $v.Rows) {
                $key = $c + '|' + [string]$r['種類']
                $txt = ([string]$r['V']); if ($c -match '^D_' -and $txt -ne '(空)') { $txt = '(日付)' }
                if (-not $vals.ContainsKey($key)) { $vals[$key] = @{} }
                $h = $vals[$key]; if ($h.ContainsKey($txt)) { $h[$txt] += [int]$r['N'] } else { $h[$txt] = [int]$r['N'] }
            }
        }
        $byKind = @{}; foreach ($r in $fill.Rows) { $byKind[[string]$r['種類']] = $r }
        W ('  行数: 親 {0} / 中身 {1} / ほか {2}' -f $(if ($byKind['親']) { $byKind['親']['行数'] } else { 0 }), $(if ($byKind['中身']) { $byKind['中身']['行数'] } else { 0 }), $(if ($byKind['ほか']) { $byKind['ほか']['行数'] } else { 0 }))
        W ('  {0,-16} {1,8} {2,8} {3,8}   {4}' -f '列', '親', '中身', 'ほか', '値 (親 / 中身。件数)')
        foreach ($c in $cols) {
            $a = @(); foreach ($kk in @('親','中身','ほか')) { if ($byKind[$kk]) { $a += [string]$byKind[$kk][$c] } else { $a += '-' } }
            $vt = ''
            if ($flagCols -contains $c) {
                $s = @()
                foreach ($kk in @('親','中身')) {
                    $h = $vals[$c + '|' + $kk]
                    if ($h) { $s += ($kk + ': ' + (($h.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' ')) }
                }
                $vt = $s -join ' / '
            }
            W ('  {0,-16} {1,8} {2,8} {3,8}   {4}' -f $c, $a[0], $a[1], $a[2], $vt)
        }
        W '  (数は「値が入っている行」の数。KEKKA など結果の値そのものは出していない)'
    } catch { W ("  エラー: " + $_.Exception.Message) }
}
finally { $conn.Close() }

W ''
W '=== 読み方 ==='
W '  1: 画面がオプションで作る行の全部。親が空の行や、オプションのコード以外を親に持つ行があれば、それも作る必要がある。'
W '  2・2b: コード変換が無い検査は、足しても検体依頼に載らないかもしれない (セットの親は無くてよいことが多い)。'
W '  3・3b: 料金明細の金額を、団体のオプション定義から入れているか。0 円の団体は請求を別にしている団体の可能性。'
W '  4・4b・4c: 連番 99 が合計の行なら、ツールで足すときは 99 も直す必要がある。'
W '  6: 親の行・中身の行に入れる印の値。ツールで作るならこれと同じにする。'
Write-Host ''
Write-Host " 終わりました: $out" -ForegroundColor Green
Start-Process notepad.exe $out
