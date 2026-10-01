<#
  オプションごとの子項目の一覧を取る (check76.ps1)
  check76.bat をダブルクリックすると実行され、結果 r_check76.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  これが最後の下調べ。
  check75 で、オプションを付けると T_KENSA に
     ・親行 1行 (例 OPJ014)
     ・子行 N行 (OYA_KOMOKU_CD = OPJ014 の行。例 017210, 017210A/B/C)
  ができると分かった。T_SET2 にはオプションの中身が入っていないので、
  実際のデータから「どのオプションに何の子項目が付くか」を拾う。
  福生だけだと件数が少ないオプションがあるので、全団体・全期間から集める。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check76.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
# 福生で使うオプション (O 溶接ヒュームは無視)
$OPTS = "'OPJ004','OPJ009','OPJ010','OPJ011','OPJ014','OPJ015','OPJ017','OPJ018','OPJ019','OPJ020','OP0008'"
"=== オプションごとの子項目 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★★福生で使うオプションの子項目 (全団体・全期間から) ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS オプション,
       LTRIM(RTRIM(ISNULL(o.MEISYO1,''))) AS オプション名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 子項目CD,
       LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 子項目名,
       COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
LEFT JOIN T_KOMOKU o ON LTRIM(RTRIM(o.KOMOKU_CD)) = LTRIM(RTRIM(k.OYA_KOMOKU_CD))
WHERE LTRIM(RTRIM(k.OYA_KOMOKU_CD)) IN ($OPTS)
GROUP BY k.OYA_KOMOKU_CD, o.MEISYO1, k.KOMOKU_CD, m.MEISYO1
ORDER BY 1, 3
"@ 200

Q '--- 2. そのオプションを受けた人数 (子項目が全員に付いているかの確認用) ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS オプション, COUNT(DISTINCT k.PK_SEQ) AS 人数,
       COUNT(DISTINCT LTRIM(RTRIM(k.KOMOKU_CD))) AS 子項目の種類, COUNT(*) AS 行数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.OYA_KOMOKU_CD)) IN ($OPTS)
GROUP BY k.OYA_KOMOKU_CD ORDER BY 1
"@ 40

Q '--- 3. 親行 (オプションのコードそのもの) が T_KENSA にどう入っているか ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 名前,
       LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,'(空)'))) AS 親項目CD,
       LTRIM(RTRIM(ISNULL(k.F_KOMOKU,'(空)'))) AS F_KOMOKU,
       LTRIM(RTRIM(ISNULL(k.F_HANTEI,'(空)'))) AS F_HANTEI,
       COUNT(*) AS 行数
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) IN ($OPTS)
GROUP BY k.KOMOKU_CD, m.MEISYO1, k.OYA_KOMOKU_CD, k.F_KOMOKU, k.F_HANTEI
ORDER BY 1
"@ 60

Q '--- 4. 子行の F_KOMOKU / F_HANTEI / HANTEI_CD に何が入っているか ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS オプション, LTRIM(RTRIM(k.KOMOKU_CD)) AS 子項目CD,
       LTRIM(RTRIM(ISNULL(k.F_KOMOKU,'(空)'))) AS F_KOMOKU,
       LTRIM(RTRIM(ISNULL(k.F_HANTEI,'(空)'))) AS F_HANTEI,
       LTRIM(RTRIM(ISNULL(k.HANTEI_CD,'(空)'))) AS HANTEI_CD,
       LTRIM(RTRIM(ISNULL(k.KEKKA_CD,'(空)'))) AS KEKKA_CD,
       COUNT(*) AS 行数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.OYA_KOMOKU_CD)) IN ($OPTS)
GROUP BY k.OYA_KOMOKU_CD, k.KOMOKU_CD, k.F_KOMOKU, k.F_HANTEI, k.HANTEI_CD, k.KEKKA_CD
ORDER BY 1, 2
"@ 200

Q '--- 5. オプションを受けた人に、くくり判定の行も増えるか (乳腺・腹部超音波など) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 判定項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 名前,
       COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) IN ('600260','600210','600282','600200','600480','600481','600491','600470','600380','600390')
  AND k.PK_SEQ IN (SELECT DISTINCT PK_SEQ FROM T_KENSA WHERE LTRIM(RTRIM(OYA_KOMOKU_CD)) IN ($OPTS))
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY 1
"@ 40

Q '--- 6. T_RYOUKIN の REN (連番) の採り方 ---' @"
SELECT r.PK_SEQ, COUNT(*) AS 行数, MIN(r.REN) AS 連番最小, MAX(r.REN) AS 連番最大
FROM T_RYOUKIN r
JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/09/01' AND '2026/09/30' AND d.MEISYO1 LIKE N'%福生%'
GROUP BY r.PK_SEQ HAVING COUNT(*) > 1 ORDER BY COUNT(*) DESC
"@ 40

Q '--- 7. T_RYOUKIN の中身の実物 (オプションのある人を1人) ---' @"
SELECT TOP 10 * FROM T_RYOUKIN
WHERE PK_SEQ IN (SELECT TOP 1 PK_SEQ FROM T_RYOUKIN WHERE LTRIM(RTRIM(KOMOKU_CD)) = 'OPJ014')
ORDER BY REN
"@ 20

W ''
W '=== 読み方 ==='
W '  1 が本命。オプションごとに作るべき子行の一覧。'
W '    「人数」がそのオプションを受けた人数と同じなら、全員に付く子項目。'
W '    一部の人にしかない子項目があれば、条件付き (性別など) かもしれない。'
W '  3・4 で、行を作るときに F_KOMOKU / F_HANTEI などに何を入れるかが分かる。'
W '  5 で、くくり判定の行も増やす必要があるかが分かる。'
W '  6・7 で T_RYOUKIN の連番の採り方と、入れる値が分かる。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
