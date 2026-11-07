<#
  新島: 当日追加の血液オプションを既存の予約に入れられるかの下調べ (check119.ps1)
  健診ナビの画面でオプションを付けたときに何が作られるか (料金明細・検査の行・検体依頼に使う印) を、
  マスタと実際のデータから読む。★人数とマスタだけ。氏名は出さない。DBは読むだけ (SELECT のみ)。
  check119.bat → r_check119.txt
#>
[CmdletBinding()]
param([string]$From = '2026/11/01', [string]$To = '2026/11/30', [string]$Words = '新島,式根島')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check119.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 新島: 当日追加オプションの下調べ ($From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W '  ★マスタと人数だけ。氏名は出していない。DBは読むだけ。'
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$wl = @(($Words -split '[,、\s]+') | Where-Object { $_ -ne '' })
$likes = ($wl | ForEach-Object { "d.MEISYO1 LIKE N'%" + ($_ -replace "'", "''") + "%'" }) -join ' OR '
$DANTAI = "(SELECT d.DANTAI_CD1 FROM M_DANTAI d WHERE $likes)"
$WHO = "s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND s.DANTAI_CD1 IN $DANTAI"
# 新島のオプション = 新島の団体に登録されたもの (M_OPTION) ＋ 新島の予約に実際に付いているもの (料金明細の OP〜)
$OPTS = @"
(SELECT LTRIM(RTRIM(o.KOMOKU_CD)) FROM M_OPTION o WHERE o.DANTAI_CD1 IN $DANTAI
 UNION SELECT LTRIM(RTRIM(r.KOMOKU_CD)) FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
       WHERE $WHO AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) LIKE 'OP%')
"@

Q '--- 1. ★新島の団体に登録されているオプション (M_OPTION。コースごと・料金つき) ---' @"
SELECT LTRIM(RTRIM(o.KOMOKU_CD)) AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(k.MEISYO1,'')))) AS オプション,
       LTRIM(RTRIM(ISNULL(o.COURSE_CD,''))) AS コースCD, o.KOJIN_RYOUKIN AS 個人料金, o.DANTAI_RYOUKIN AS 団体料金,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), o.K_ADD_DEL),''))) AS 追加削除, COUNT(DISTINCT o.DANTAI_CD1) AS 団体数,
       MAX(LTRIM(RTRIM(ISNULL(k.K_SET,'')))) AS セット, MAX(LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.OPTION_GRP),'')))) AS オプション群
FROM M_OPTION o LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(o.KOMOKU_CD))
WHERE o.DANTAI_CD1 IN $DANTAI
GROUP BY LTRIM(RTRIM(o.KOMOKU_CD)), o.COURSE_CD, o.KOJIN_RYOUKIN, o.DANTAI_RYOUKIN, o.K_ADD_DEL
ORDER BY 1, 3
"@ 200

Q '--- 2. ★オプションごとの中身の検査 (実際の予約の T_KENSA から。全団体・全期間) と検体依頼の印 (T_KOMOKU.F_IRAI) ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS オプション, LTRIM(RTRIM(k.KOMOKU_CD)) AS 中身の項目CD,
       MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 項目名, MAX(LTRIM(RTRIM(ISNULL(m.F_IRAI,'')))) AS 検体依頼の印,
       MAX(LTRIM(RTRIM(ISNULL(m.JLAC10,'')))) AS JLAC10, MAX(LTRIM(RTRIM(ISNULL(m.K_KOMOKU,'')))) AS 区分,
       COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSA k LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LTRIM(RTRIM(k.OYA_KOMOKU_CD)) IN $OPTS
GROUP BY LTRIM(RTRIM(k.OYA_KOMOKU_CD)), LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1, 2
"@ 400

Q '--- 3. オプションの親の行 (T_KENSA) にどんな印が入っているか ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,'(空)'))) AS 親,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.F_KOMOKU),'(空)'))) AS F_KOMOKU,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.F_HANTEI),'(空)'))) AS F_HANTEI, COUNT(*) AS 行数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.KOMOKU_CD)) IN $OPTS
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), k.OYA_KOMOKU_CD, k.F_KOMOKU, k.F_HANTEI ORDER BY 1
"@ 80

Q '--- 4. 中身の行の F_KOMOKU (予約のときに付いたか、あとから足したか の印のはず) ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS オプション, LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.F_KOMOKU),'(空)'))) AS F_KOMOKU, COUNT(*) AS 行数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.OYA_KOMOKU_CD)) IN $OPTS
GROUP BY LTRIM(RTRIM(k.OYA_KOMOKU_CD)), k.F_KOMOKU ORDER BY 1, 2
"@ 80

Q '--- 5. 料金明細 (T_RYOUKIN) の列 ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_RYOUKIN' ORDER BY c.ORDINAL_POSITION
"@ 60

Q '--- 6. ★新島のオプションの料金明細の中身の形 (オプションごと。氏名なし) ---' @"
SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS 項目CD, ISNULL(CONVERT(varchar(10), r.F_ADD_DEL),'') AS 追加削除,
       LTRIM(RTRIM(ISNULL(r.TEKIYOU,''))) AS 摘要, COUNT(*) AS 行数, MIN(r.REN) AS 連番最小, MAX(r.REN) AS 連番最大
FROM T_RYOUKIN r WHERE LTRIM(RTRIM(r.KOMOKU_CD)) IN $OPTS
GROUP BY LTRIM(RTRIM(r.KOMOKU_CD)), r.F_ADD_DEL, r.TEKIYOU ORDER BY 1
"@ 80

Q '--- 7. 料金明細の行の例 (新島のオプション 3 行。金額などの列の入り方を見る) ---' @"
SELECT TOP 3 r.* FROM T_RYOUKIN r WHERE LTRIM(RTRIM(r.KOMOKU_CD)) IN $OPTS ORDER BY r.PK_SEQ DESC, r.REN
"@ 5

Q '--- 8. 予約 (T_KENSIN) で金額・依頼・受付にかかわりそうな列 (列名だけ) ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME = 'T_KENSIN' AND (c.COLUMN_NAME LIKE '%KIN%' OR c.COLUMN_NAME LIKE '%GAKU%' OR c.COLUMN_NAME LIKE '%RYO%'
  OR c.COLUMN_NAME LIKE '%IRAI%' OR c.COLUMN_NAME LIKE '%OPT%' OR c.COLUMN_NAME LIKE '%SEQ%' OR c.COLUMN_NAME LIKE '%UKE%')
ORDER BY c.ORDINAL_POSITION
"@ 60

Q '--- 9. 検体依頼にかかわりそうな表 (名前に IRAI・KENTAI・LAB が入る表と列) ---' @"
SELECT c.TABLE_NAME AS 表, c.COLUMN_NAME AS 列 FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE c.TABLE_NAME LIKE '%IRAI%' OR c.TABLE_NAME LIKE '%KENTAI%' OR c.TABLE_NAME LIKE '%LAB%' OR c.COLUMN_NAME LIKE '%IRAI%'
ORDER BY 1, 2
"@ 120

Q '--- 10. 検体依頼の印 (T_KOMOKU.F_IRAI) の付き方: 基本の血液項目と比べる ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名, LTRIM(RTRIM(ISNULL(F_IRAI,''))) AS 検体依頼の印
FROM T_KOMOKU WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068689','068709','069005','068682','R06-0002','069245')
ORDER BY 1
"@ 20

Q '--- 11. ★重複の心配: 新島の予約で、オプションの中身の検査がもう枠にある人 (基本コースに入っている・すでにオプションがある) ---' @"
SELECT LTRIM(RTRIM(c.OYA)) AS オプション, LTRIM(RTRIM(c.CD)) AS 中身の項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 項目名,
       COUNT(DISTINCT k.PK_SEQ) AS もう枠がある人,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,''))) = '' THEN 1 ELSE 0 END) AS うち基本コースの枠
FROM (SELECT DISTINCT LTRIM(RTRIM(x.OYA_KOMOKU_CD)) AS OYA, LTRIM(RTRIM(x.KOMOKU_CD)) AS CD FROM T_KENSA x WHERE LTRIM(RTRIM(x.OYA_KOMOKU_CD)) IN $OPTS) c
JOIN T_KENSA k ON LTRIM(RTRIM(k.KOMOKU_CD)) = c.CD
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = c.CD
WHERE $WHO
GROUP BY c.OYA, c.CD ORDER BY 1, 2
"@ 300

Q '--- 12. 新島の予約で、オプションごとにもう付いている人数 (料金明細 / 検査の親の行) ---' @"
SELECT x.CD AS オプション,
       (SELECT COUNT(DISTINCT r.PK_SEQ) FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ WHERE $WHO AND LTRIM(RTRIM(r.KOMOKU_CD)) = x.CD) AS 料金明細あり,
       (SELECT COUNT(DISTINCT k.PK_SEQ) FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ WHERE $WHO AND LTRIM(RTRIM(k.KOMOKU_CD)) = x.CD) AS 検査の親の行あり
FROM (SELECT DISTINCT LTRIM(RTRIM(o.KOMOKU_CD)) AS CD FROM M_OPTION o WHERE o.DANTAI_CD1 IN $DANTAI) x
ORDER BY 1
"@ 60

W ''
W '=== 読み方 ==='
W '  1 が新島で使えるオプション (8 種類のはず) と料金。2 がそれぞれの中身の検査と、検体依頼の印。'
W '  3・4・6・7 が、画面でオプションを付けたときに健診ナビが入れている印・料金の形。ツールで同じものを作れるかの判断材料。'
W '  9・10 で、検体依頼書が何を見て作られるか (項目の印か、別の表か) の手がかりを見る。'
W '  11・12 で、基本コースやすでに付いているオプションと重ならないかを見る。'
notepad $out
