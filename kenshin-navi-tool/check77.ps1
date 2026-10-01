<#
  受診日を動かすと何が食い違うのかを調べる (check77.ps1)
  check77.bat をダブルクリックすると実行され、結果 r_check77.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  なぜ調べるか
    6日予定の人が2日に来た場合、健診ナビの予約の受診日を動かす必要がある。
    ところが T_KENSA の SEQ1 / SEQ2 は「受診日(YYYYMMDD) + 4桁連番」で作られている。
    受診日を動かしたとき、健診ナビが SEQ1/SEQ2 まで作り直しているのか、
    それとも作ったときのまま残しているのかが分からない。

    過去に受診日を動かした人がいれば、D_KENSIN と SEQ1 の頭8桁がズレて残っている。
    それを数えれば「受診日を動かすのが安全な操作か」が分かる。

  見たいこと
    1. D_KENSIN と SEQ1 の頭8桁が食い違っている予約が、過去にどれだけあるか
    2. 食い違っている人は、結果票や請求がちゃんと出ているのか (T_KENSA の中身)
    3. 受診日を持っている列が、ほかにどのテーブルにあるか
       (受診日を動かすとき一緒に直さないといけない場所の洗い出し)
    4. 10月の福生で、同じ人が複数の日に予約を持っていないか
#>
[CmdletBinding()]
param([string]$Name = '福生')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check77.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受診日を動かすと何が食い違うか $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★D_KENSIN と SEQ1 の頭8桁が食い違う予約の数 (過去2年) ---' @"
SELECT CASE WHEN LEFT(LTRIM(RTRIM(k.SEQ1)), 8) = REPLACE(CONVERT(varchar(10), s.D_KENSIN, 111), '/', '')
            THEN N'一致' ELSE N'食い違い' END AS 状態,
       COUNT(DISTINCT s.PK_SEQ) AS 予約数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2024/10/01'
GROUP BY CASE WHEN LEFT(LTRIM(RTRIM(k.SEQ1)), 8) = REPLACE(CONVERT(varchar(10), s.D_KENSIN, 111), '/', '')
              THEN N'一致' ELSE N'食い違い' END
"@ 20

Q '--- 2. ★食い違っている人の実物 (受診日と SEQ1 の頭) ---' @"
SELECT TOP 40 s.PK_SEQ, CONVERT(varchar(10), s.D_KENSIN, 111) AS 健診ナビの受診日,
       LEFT(LTRIM(RTRIM(MIN(k.SEQ1))), 8) AS SEQ1の頭8桁,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
       s.UKE_NO_KENSA AS 受付番号,
       COUNT(*) AS 検査行数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 結果が入っている行数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2024/10/01'
  AND LEFT(LTRIM(RTRIM(k.SEQ1)), 8) <> REPLACE(CONVERT(varchar(10), s.D_KENSIN, 111), '/', '')
GROUP BY s.PK_SEQ, s.D_KENSIN, d.MEISYO1, s.COURSE_CD, s.UKE_NO_KENSA
ORDER BY s.D_KENSIN DESC
"@ 50

Q '--- 3. 受診日らしい列を持っているテーブル (一緒に直す場所の洗い出し) ---' @"
SELECT o.name AS テーブル, c.name AS 列, ty.name AS 型
FROM sys.columns c
JOIN sys.objects o ON o.object_id = c.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE o.type = 'U'
  AND (c.name LIKE '%KENSIN%' OR c.name LIKE '%KEN_YMD%' OR c.name LIKE 'D_KEN%'
       OR c.name LIKE '%JUSIN%' OR c.name = 'SEQ1' OR c.name = 'SEQ2')
ORDER BY o.name, c.column_id
"@ 120

Q '--- 4. SEQ1 / SEQ2 を持っているテーブルの件数 ---' @"
SELECT o.name AS テーブル, COUNT(*) AS SEQ列の数
FROM sys.columns c JOIN sys.objects o ON o.object_id = c.object_id
WHERE o.type = 'U' AND c.name IN ('SEQ1','SEQ2')
GROUP BY o.name ORDER BY o.name
"@ 60

Q '--- 5. ★10月の福生で、同じ人が2つ以上の日に予約を持っていないか ---' @"
SELECT LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名, COUNT(*) AS 予約数,
       MIN(CONVERT(varchar(10), s.D_KENSIN, 111)) AS 受診日1,
       MAX(CONVERT(varchar(10), s.D_KENSIN, 111)) AS 受診日2,
       MIN(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 事業所
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/10/01' AND '2026/10/31'
  AND d.MEISYO1 LIKE N'%$Name%'
GROUP BY g.KANJI_SIMEI HAVING COUNT(*) > 1
ORDER BY COUNT(*) DESC, 1
"@ 60

Q '--- 6. ★10月の福生で、同姓同名が同じ日にいないか (氏名で突き合わせる弱点の確認) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名, COUNT(*) AS 人数
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/10/01' AND '2026/10/31'
  AND d.MEISYO1 LIKE N'%$Name%'
GROUP BY s.D_KENSIN, g.KANJI_SIMEI HAVING COUNT(*) > 1
ORDER BY 1, 2
"@ 60

Q '--- 7. 10月の福生の予約の状態 (日付ごと。受付番号が入っているか) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 予約数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0 THEN 1 ELSE 0 END) AS 受付番号なし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) = '' THEN 1 ELSE 0 END) AS カナなし,
       COUNT(DISTINCT LTRIM(RTRIM(ISNULL(s.COURSE_CD,'')))) AS コースの種類
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/10/01' AND '2026/10/31'
  AND d.MEISYO1 LIKE N'%$Name%'
GROUP BY s.D_KENSIN ORDER BY 1
"@ 40

Q '--- 9. ★受付番号は同じ日の中で重複しないのか (過去2年) ---' @"
SELECT CASE WHEN c.人数 = 1 THEN N'その日に1人だけ' ELSE N'同じ日に同じ番号が複数' END AS 状態,
       COUNT(*) AS 組み合わせ数
FROM (SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, s.UKE_NO_KENSA AS UKE, COUNT(*) AS 人数
      FROM T_KENSIN s
      WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2024/10/01'
        AND s.UKE_NO_KENSA IS NOT NULL AND s.UKE_NO_KENSA <> 0
      GROUP BY s.D_KENSIN, s.UKE_NO_KENSA) c
GROUP BY CASE WHEN c.人数 = 1 THEN N'その日に1人だけ' ELSE N'同じ日に同じ番号が複数' END
"@ 20

Q '--- 10. ★T_KANJA_G (受診日と番号を持つテーブル) の中身 ---' @"
SELECT TOP 20 * FROM T_KANJA_G ORDER BY 1 DESC
"@ 30

Q '--- 11. ★10月の福生の予約に、既に埋まっている受付番号があるか (移動先で衝突しないか) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       MIN(s.UKE_NO_KENSA) AS 番号の最小, MAX(s.UKE_NO_KENSA) AS 番号の最大,
       COUNT(*) AS 番号が入っている人数
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/10/01' AND '2026/10/31'
  AND d.MEISYO1 LIKE N'%$Name%'
  AND s.UKE_NO_KENSA IS NOT NULL AND s.UKE_NO_KENSA <> 0
GROUP BY s.D_KENSIN ORDER BY 1
"@ 40

Q '--- 12. ★受診日を変えたとき一緒に動くはずのテーブルに、PK_SEQ 以外の受診日があるか ---' @"
SELECT o.name AS テーブル, c.name AS 列, ty.name AS 型
FROM sys.columns c
JOIN sys.objects o ON o.object_id = c.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE o.type = 'U'
  AND EXISTS (SELECT 1 FROM sys.columns p WHERE p.object_id = o.object_id AND p.name = 'PK_SEQ')
  AND (ty.name IN ('date','datetime','smalldatetime') OR c.name LIKE '%YMD%' OR c.name LIKE '%SEQ%')
ORDER BY o.name, c.column_id
"@ 150

Q '--- 8. T_KENSIN_LOG に受診日を変えた履歴が残っているか ---' @"
SELECT TOP 20 c.name AS 列, ty.name AS 型
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KENSIN_LOG') ORDER BY c.column_id
"@ 30

W ''
W '=== 読み方 ==='
W '  1 が本命。'
W '    「食い違い」が 0 なら、健診ナビは受診日を動かすときに SEQ1/SEQ2 も作り直している。'
W '      → 受診日の変更は健診ナビの画面でやれば安全。私のツールは触らない。'
W '    「食い違い」が沢山あるなら、健診ナビは SEQ1/SEQ2 を作ったままにしている。'
W '      → SEQ1 は過去の日付のまま残る。それで結果票が出ているなら実害は無い、と判断できる。'
W '  2 で、食い違っている人に結果がちゃんと入っているかを見る。入っていれば実害なし。'
W '  3・4 で、受診日を動かすとき一緒に直さないといけない場所が分かる。'
W '    ここが T_KENSIN だけなら話は簡単。他にもあるなら健診ナビの画面に任せるべき。'
W '  5 が 0 件なら、同じ人が2つの日に予約を持っている問題は無い。'
W '  6 が 0 件なら、氏名で突き合わせても取り違えは起きない。'
W '  7 で今の10月の予約の状態 (受付番号・カナの入り具合) が分かる。'
W ''
W '  ここから下は「予約の受診日を動かす」案のための確認。'
W '  9  「同じ日に同じ番号が複数」が 0 なら、受付番号は日の中で一意。'
W '      → 日付を動かすときは、移動先の日で番号が衝突しないかを必ず見ないといけない。'
W '      0 でないなら、健診ナビは一意性を気にしていない。'
W '  10 T_KANJA_G に受診日 (KEN_YMD) が入っているなら、ここも一緒に直す対象。'
W '  11 10月の福生で、既に受付番号が入っている予約があるか。'
W '      あれば、それが今日入れる番号とぶつからないかを見る必要がある。'
W '  12 PK_SEQ を持つテーブルの日付列・SEQ列の一覧。'
W '      受診日を動かすときに取り残される場所の洗い出し。'
W '      ここが T_KENSIN の D_KENSIN だけなら、ツールで動かしてよい。'
W '      T_KENSA の SEQ1/SEQ2 や T_KANJA_G の KEN_YMD にも日付が埋まっているなら、'
W '      1列だけ書き換えるのは不整合になるので、健診ナビの画面に任せるべき。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
