<#
  オプション・便本数・検査の追加中止を健診ナビのどこに入れているか調べる (check73.ps1)
  check73.bat をダブルクリックすると実行され、結果 r_check73.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  やりたいこと
    会場受付アプリの当日変更を健診ナビに反映したい。反映先が未確定なのは次の5つ。
      ① オプション (D〜O) を受けた人
      ② 便の提出本数 (0/1/2)
      ③ 採血・胃・尿の当日の追加/中止
      ④ 受診日が予定と変わった人
      ⑤ キャンセル
    9月の福生3社 (嵯峨野9/14・福生コンクリート9/17・関東照明9/24) は同じ巡回健診なので、
    そこで実際にどう記録されているかを見れば、書き先が決まる。
#>
[CmdletBinding()]
param([string]$Name = '福生', [string]$From = '2026/09/01', [string]$To = '2026/09/30')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check73.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
$SCOPE = "s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND d.MEISYO1 LIKE N'%$Name%'"
"=== 当日変更の書き先をさがす ($Name $From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★同じコースなのに検査項目の数が違う人がいるか (= 人ごとに項目を足し引きしている) ---' @"
SELECT LTRIM(RTRIM(s.COURSE_CD)) AS コースCD,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = s.DANTAI_CD1 AND x.COURSE_CD = s.COURSE_CD) AS コースの項目数,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS この人の検査行数,
       COUNT(*) AS 人数
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE $SCOPE
GROUP BY s.COURSE_CD, s.DANTAI_CD1,
         (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = s.DANTAI_CD1 AND x.COURSE_CD = s.COURSE_CD),
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ)
ORDER BY 1, 3
"@ 60

Q '--- 2. ★T_RYOUKIN (受付で入力した料金明細) に何が入っているか ---' @"
SELECT LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) AS 項目CD, LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目,
       COUNT(*) AS 件数, MIN(r.GOUKEI) AS 金額最小, MAX(r.GOUKEI) AS 金額最大
FROM T_RYOUKIN r
JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
WHERE $SCOPE
GROUP BY r.KOMOKU_CD, k.MEISYO1 ORDER BY COUNT(*) DESC
"@ 60

Q '--- 3. T_RYOUKIN の列 (何を入れる表なのか) ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_RYOUKIN') ORDER BY c.column_id
"@ 40

Q '--- 4. ★便潜血 (069245/069246) が人ごとにどう入っているか ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 検査項目,
       COUNT(*) AS 行数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) = '' THEN 1 ELSE 0 END) AS 結果が空,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 結果あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $SCOPE AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('069245','069246')
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY 1
"@ 30

Q '--- 5. ★福生の団体に登録されているオプション (M_OPTION) ---' @"
SELECT TOP 60 LTRIM(RTRIM(o.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       LTRIM(RTRIM(o.COURSE_CD)) AS コースCD, LTRIM(RTRIM(o.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(o.K_ADD_DEL,''))) AS 追加削除,
       o.KOJIN_RYOUKIN AS 個人料金, o.DANTAI_RYOUKIN AS 団体料金
FROM M_OPTION o
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = o.DANTAI_CD1
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(o.KOMOKU_CD))
WHERE d.MEISYO1 LIKE N'%$Name%'
ORDER BY o.DANTAI_CD1, o.COURSE_CD, o.KOMOKU_CD
"@ 80

Q '--- 6. 9月の福生で、コースに無い検査が入っている人 (= 当日追加したもの) ---' @"
SELECT TOP 60 LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 検査項目,
       COUNT(DISTINCT s.PK_SEQ) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $SCOPE
  AND NOT EXISTS (SELECT 1 FROM T_COURSE2 c2
                  WHERE c2.DANTAI_CD1 = s.DANTAI_CD1 AND c2.COURSE_CD = s.COURSE_CD
                    AND LTRIM(RTRIM(c2.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)))
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY COUNT(DISTINCT s.PK_SEQ) DESC
"@ 80

Q '--- 7. 逆に、コースにあるのに検査行が無いもの (= 当日中止したもの) ---' @"
SELECT TOP 40 LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 検査項目,
       COUNT(*) AS 人数
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
JOIN T_COURSE2 c2 ON c2.DANTAI_CD1 = s.DANTAI_CD1 AND c2.COURSE_CD = s.COURSE_CD
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE $SCOPE
  AND NOT EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
                    AND LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD)))
GROUP BY c2.KOMOKU_CD, m.MEISYO1 ORDER BY COUNT(*) DESC
"@ 60

Q '--- 8. キャンセル (F_TORIKESI) の使われ方 ---' @"
SELECT s.F_TORIKESI AS 取消, COUNT(*) AS 件数,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) THEN 1 ELSE 0 END) AS 検査行あり
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN BETWEEN '$From' AND '$To' AND d.MEISYO1 LIKE N'%$Name%'
GROUP BY s.F_TORIKESI
"@ 20

Q '--- 9. 受診日を後から動かした跡があるか (更新日時の列があれば) ---' @"
SELECT c.name AS 列, ty.name AS 型
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KENSIN')
  AND (c.name LIKE '%UPD%' OR c.name LIKE '%KOSIN%' OR c.name LIKE '%TOUROKU%' OR c.name LIKE '%D_%')
ORDER BY c.column_id
"@ 60

W ''
W '=== 読み方 ==='
W '  1 で「コースの項目数」と「この人の検査行数」が同じなら、'
W '    検査行はコースから機械的に作られていて、人ごとの足し引きはしていない。'
W '    違う人がいれば、人ごとに項目を足し引きしている。'
W '  2 の T_RYOUKIN に行があれば、オプションは「受付で料金を入力する」運用。'
W '  6 に項目が出れば、当日追加は T_KENSA に行を足す形。'
W '  7 に項目が出れば、当日中止は T_KENSA の行を作らない形。'
W '  4 で便潜血が 1回目だけある人がいるかが分かる。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
