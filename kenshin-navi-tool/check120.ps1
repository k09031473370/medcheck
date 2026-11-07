<#
  新島: 当日追加オプションの下調べ その 2 (check120.ps1)
  check119 で分かったことの続き。マスタと人数だけ。氏名は出さない。DBは読むだけ (SELECT のみ)。
    1. セットの中にさらにセットがあるもの (脳梗塞リスク・ABC・アレルギー) の孫の行
    2. 検体依頼のコード変換 (M_HENKAN_IRAI) と容器 (M_KENTAI_YOUKI) にオプションの検査が載っているか
    3. 料金明細の連番 1 は何か・オプションの行のほかの列の入り方
    4. 予約の合計 (TOKUTEI_GOUKEI) が料金明細の合計と合っているか (画面が合計も書き換えているか)
    5. 予約の 受付・検体依頼 の印の入り方
  check120.bat → r_check120.txt
#>
[CmdletBinding()]
param([string]$From = '2026/11/01', [string]$To = '2026/11/30', [string]$Words = '新島,式根島')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check120.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 新島: 当日追加オプションの下調べ その 2 ($From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W '  ★マスタと人数だけ。氏名は出していない。DBは読むだけ。'
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$wl = @(($Words -split '[,、\s]+') | Where-Object { $_ -ne '' })
$likes = ($wl | ForEach-Object { "d.MEISYO1 LIKE N'%" + ($_ -replace "'", "''") + "%'" }) -join ' OR '
$WHO = "s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND s.DANTAI_CD1 IN (SELECT d.DANTAI_CD1 FROM M_DANTAI d WHERE $likes)"
$OPS = "('OP0001','OP0002','OP0003','OP0004','OP0005','OP0006','OP0007','OP0008','OP0009')"
# オプションの子と孫 (子がセットのとき、その下の行)
$KIDS = "(SELECT DISTINCT LTRIM(RTRIM(KOMOKU_CD)) FROM T_KENSA WHERE LTRIM(RTRIM(OYA_KOMOKU_CD)) IN $OPS)"
$ALL = "(SELECT DISTINCT LTRIM(RTRIM(KOMOKU_CD)) FROM T_KENSA WHERE LTRIM(RTRIM(OYA_KOMOKU_CD)) IN $OPS OR LTRIM(RTRIM(OYA_KOMOKU_CD)) IN $KIDS)"

Q '--- 1. ★オプションの子がセットのとき、その下の行 (孫) ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS 親 , LTRIM(RTRIM(k.KOMOKU_CD)) AS 孫の項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 項目名,
       MAX(LTRIM(RTRIM(ISNULL(m.F_IRAI,'')))) AS 検体依頼の印, MAX(LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.F_KOMOKU),'')))) AS F_KOMOKU, COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSA k LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LTRIM(RTRIM(k.OYA_KOMOKU_CD)) IN $KIDS
GROUP BY LTRIM(RTRIM(k.OYA_KOMOKU_CD)), LTRIM(RTRIM(k.KOMOKU_CD)) ORDER BY 1, 2
"@ 200

Q '--- 2. ★検体依頼のコード変換 (M_HENKAN_IRAI) にオプションの検査が載っているか (検査センターごと) ---' @"
SELECT LTRIM(RTRIM(h.KOMOKU_CD)) AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 項目名, LTRIM(RTRIM(ISNULL(h.CENTER_CD,''))) AS センター,
       LTRIM(RTRIM(ISNULL(h.OUT_CD,''))) AS 依頼コード, LTRIM(RTRIM(ISNULL(h.OUT_CD_OYA,''))) AS 依頼コード親,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), h.KENTAI_KBN),''))) AS 検体区分, LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), h.YOUKI_CD),''))) AS 容器CD,
       MAX(LTRIM(RTRIM(ISNULL(y.YOUKI_MEISYO,'')))) AS 容器
FROM M_HENKAN_IRAI h
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(h.KOMOKU_CD))
LEFT JOIN M_KENTAI_YOUKI y ON LTRIM(RTRIM(CONVERT(varchar(20), y.YOUKI_CD))) = LTRIM(RTRIM(CONVERT(varchar(20), h.YOUKI_CD)))
WHERE LTRIM(RTRIM(h.KOMOKU_CD)) IN $OPS OR LTRIM(RTRIM(h.KOMOKU_CD)) IN $ALL
GROUP BY h.KOMOKU_CD, h.CENTER_CD, h.OUT_CD, h.OUT_CD_OYA, h.KENTAI_KBN, h.YOUKI_CD ORDER BY 1, 3
"@ 300

Q '--- 2b. 比べる: 基本の血液項目の検体依頼のコード変換 ---' @"
SELECT LTRIM(RTRIM(h.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(h.CENTER_CD,''))) AS センター, LTRIM(RTRIM(ISNULL(h.OUT_CD,''))) AS 依頼コード,
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), h.KENTAI_KBN),''))) AS 検体区分, LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), h.YOUKI_CD),''))) AS 容器CD
FROM M_HENKAN_IRAI h WHERE LTRIM(RTRIM(h.KOMOKU_CD)) IN ('068689','068709','069005','068682')
ORDER BY 1, 2
"@ 40

Q '--- 2c. 検体容器の一覧 (M_KENTAI_YOUKI) ---' @"
SELECT LTRIM(RTRIM(CONVERT(varchar(20), YOUKI_CD))) AS 容器CD, LTRIM(RTRIM(ISNULL(YOUKI_MEISYO,''))) AS 容器, LTRIM(RTRIM(ISNULL(ZAIRYO,''))) AS 材料,
       MAISUU AS 枚数, LTRIM(RTRIM(ISNULL(KANREN_KOMOKU,''))) AS 関連項目
FROM M_KENTAI_YOUKI ORDER BY 1
"@ 60

Q '--- 3. ★料金明細の連番ごとに何が入っているか (新島の 11 月の予約) ---' @"
SELECT r.REN AS 連番, CASE WHEN LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) = '' THEN '(空)' WHEN r.KOMOKU_CD LIKE 'OP%' THEN 'OP…' ELSE LTRIM(RTRIM(r.KOMOKU_CD)) END AS 項目CD,
       ISNULL(CONVERT(varchar(10), r.F_ADD_DEL),'') AS 追加削除, COUNT(*) AS 行数, MIN(r.GOUKEI) AS 合計最小, MAX(r.GOUKEI) AS 合計最大
FROM T_RYOUKIN r JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ WHERE $WHO
GROUP BY r.REN, CASE WHEN LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) = '' THEN '(空)' WHEN r.KOMOKU_CD LIKE 'OP%' THEN 'OP…' ELSE LTRIM(RTRIM(r.KOMOKU_CD)) END, r.F_ADD_DEL
ORDER BY 1, 2
"@ 60

Q '--- 3b. オプションの料金明細の行の残りの列 (例 3 行。金額・伝票・領収の列) ---' @"
SELECT TOP 3 r.REN, r.KOMOKU_CD, r.F_ADD_DEL, r.KENPO_RYOUKIN, r.DANTAI_RYOUKIN, r.KOJIN_RYOUKIN, r.GOUKEI, r.DAIKOU_RYOUKIN, r.SONOTA_RYOUKIN,
       '[' + ISNULL(r.DENPYOU_BANGO,'') + ']' AS 伝票番号, '[' + ISNULL(r.D_RYOUSYUU,'') + ']' AS 領収日, LTRIM(RTRIM(ISNULL(r.DANTAI_CD1,''))) AS 団体CD, LTRIM(RTRIM(ISNULL(r.COURSE_CD,''))) AS コースCD
FROM T_RYOUKIN r WHERE LTRIM(RTRIM(r.KOMOKU_CD)) IN ('OP0001','OP0002','OP0003','OP0004','OP0005','OP0006','OP0007','OP0008')
ORDER BY r.PK_SEQ DESC
"@ 5

Q '--- 4. ★予約の合計 (T_KENSIN.TOKUTEI_GOUKEI) と料金明細の合計 (GOUKEI の和) が合っているか (オプションのある予約) ---' @"
SELECT CASE WHEN x.T = x.R THEN N'合っている' WHEN x.T IS NULL THEN N'予約の合計が空' ELSE N'違う' END AS 結果, COUNT(*) AS 予約数
FROM (SELECT s.PK_SEQ, s.TOKUTEI_GOUKEI AS T, (SELECT SUM(r.GOUKEI) FROM T_RYOUKIN r WHERE r.PK_SEQ = s.PK_SEQ) AS R
      FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND EXISTS (SELECT 1 FROM T_RYOUKIN r WHERE r.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(r.KOMOKU_CD)) LIKE 'OP%')) x
GROUP BY CASE WHEN x.T = x.R THEN N'合っている' WHEN x.T IS NULL THEN N'予約の合計が空' ELSE N'違う' END
"@ 10

Q '--- 5. ★新島の 11 月の予約の 受付・検体依頼 の印 ---' @"
SELECT LTRIM(RTRIM(ISNULL(s.F_UKETUKE,'(空)'))) AS 受付の印, LTRIM(RTRIM(ISNULL(s.F_IRAI,'(空)'))) AS 検体依頼の印,
       LTRIM(RTRIM(ISNULL(s.F_IRAI_YUKI,'(空)'))) AS 依頼_YUKI, COUNT(*) AS 予約数
FROM T_KENSIN s WHERE $WHO
GROUP BY s.F_UKETUKE, s.F_IRAI, s.F_IRAI_YUKI ORDER BY 4 DESC
"@ 20

Q '--- 5b. 比べる: 結果が出ている過去の予約 (2026/10) の 受付・検体依頼 の印 ---' @"
SELECT LTRIM(RTRIM(ISNULL(s.F_UKETUKE,'(空)'))) AS 受付の印, LTRIM(RTRIM(ISNULL(s.F_IRAI,'(空)'))) AS 検体依頼の印,
       LTRIM(RTRIM(ISNULL(s.F_IRAI_YUKI,'(空)'))) AS 依頼_YUKI, COUNT(*) AS 予約数
FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/10/01' AND '2026/10/31'
GROUP BY s.F_UKETUKE, s.F_IRAI, s.F_IRAI_YUKI ORDER BY 4 DESC
"@ 20

Q '--- 6. 後日の表 (T_GOJITU) の列 (便の後日提出などに使うか) ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_GOJITU' ORDER BY c.ORDINAL_POSITION
"@ 40

W ''
W '=== 読み方 ==='
W '  1: オプションを足すとき、孫の行まで作る必要があるか。'
W '  2: オプションの検査に検体依頼のコード変換があるか (無ければ、足しても検体依頼書に載らない)。容器 (採血管) も分かる。'
W '  3・3b: 料金明細の連番の採り方と、オプションの行に入る値。'
W '  4: 画面が予約の合計も書き換えているか (合っていれば、ツールで足すなら合計も直す必要がある)。'
W '  5: 新島の予約で検体依頼がもう出ているか (出たあとに足すと、追加の依頼が要る)。'
notepad $out
