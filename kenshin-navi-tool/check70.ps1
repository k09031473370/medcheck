<#
  マイクロンのコース状況と、福生のコース・オプションを確かめる (check70.ps1)
  check70.bat をダブルクリックすると実行され、結果 r_check70.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  分かっていること
    ・マイクロンは 0000000477 が正 (0000000475 は使わない)
    ・0000000477 に 10/6 91人・10/7 81人・10/14・10/15 の予約が既に入っている
      → 何かのコースが既に割り当たっているはず。それを見てから コース作成 の要否を決める
    ・福生商工会 325人 (10/2・10/6・10/7) は健診ナビ未取込
    ・振り分け表のコースは A / B / C。健診ナビ側は FA 福生A / FB 福生B まで確認済み
#>
[CmdletBinding()]
param(
    [string]$Old = '0000000475',   # 使わない方のマイクロン
    [string]$New = '0000000477'    # 正しい方のマイクロン
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check70.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== マイクロンのコース状況 / 福生のコース・オプション $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. マイクロン 2つの団体の中身くらべ ($Old / $New) ---" @"
SELECT LTRIM(RTRIM(d.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       (SELECT COUNT(*) FROM T_COURSE1 x WHERE x.DANTAI_CD1 = d.DANTAI_CD1) AS コース数,
       (SELECT COUNT(*) FROM T_KENSIN x WHERE x.DANTAI_CD1 = d.DANTAI_CD1 AND x.F_TORIKESI = 0) AS 受診のべ人数,
       (SELECT COUNT(*) FROM M_OPTION x WHERE x.DANTAI_CD1 = d.DANTAI_CD1) AS オプション定義数
FROM T_DANTAI1 d WHERE d.DANTAI_CD1 IN ('$Old','$New')
"@ 20

Q "--- 2. ★$New の予約が使っているコース ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,'(空)'))) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(*) AS 人数,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = s.DANTAI_CD1 AND x.COURSE_CD = s.COURSE_CD) AS 項目数
FROM T_KENSIN s
LEFT JOIN T_COURSE1 c ON c.DANTAI_CD1 = s.DANTAI_CD1 AND c.COURSE_CD = s.COURSE_CD
WHERE s.DANTAI_CD1 = '$New' AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN, s.COURSE_CD, c.MEISYO, s.DANTAI_CD1
ORDER BY s.D_KENSIN, COUNT(*) DESC
"@ 40

Q "--- 3. $New に登録されているコース一覧 ---" @"
SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       LTRIM(RTRIM(ISNULL(c.KIJUN_CD,''))) AS 基準値CD, LTRIM(RTRIM(ISNULL(c.TYOHYO,''))) AS 帳票,
       LTRIM(RTRIM(ISNULL(c.F_TOSINKYO,''))) AS 東振協,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数,
       (SELECT TOP 1 CONVERT(varchar(20), y.DANTAI_RYOUKIN) FROM T_COURSE3 y WHERE y.DANTAI_CD1 = c.DANTAI_CD1 AND y.COURSE_CD = c.COURSE_CD) AS 団体料金
FROM T_COURSE1 c WHERE c.DANTAI_CD1 = '$New' ORDER BY c.COURSE_CD
"@ 40

Q '--- 4. ★福生のコース (コース名に「福生」が入るもの) ---' @"
SELECT LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       LTRIM(RTRIM(ISNULL(c.TYOHYO,''))) AS 帳票,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE c.MEISYO LIKE N'%福生%' OR LTRIM(RTRIM(c.COURSE_CD)) IN ('FA','FB','FC')
ORDER BY c.COURSE_CD, c.DANTAI_CD1
"@ 80

Q '--- 5. 福生の各コースの検査項目数と、ミニドック(C)に当たるものがあるか ---' @"
SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(DISTINCT c.DANTAI_CD1) AS この名前を持つ団体数,
       MIN((SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD)) AS 項目数最小,
       MAX((SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD)) AS 項目数最大
FROM T_COURSE1 c
WHERE c.MEISYO LIKE N'%福生%'
GROUP BY c.COURSE_CD, c.MEISYO ORDER BY c.COURSE_CD
"@ 40

Q '--- 6. ★福生の団体に登録されているオプション (M_OPTION) ---' @"
SELECT TOP 60 LTRIM(RTRIM(o.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(o.COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(o.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目,
       o.KOJIN_RYOUKIN AS 個人料金, o.DANTAI_RYOUKIN AS 団体料金, LTRIM(RTRIM(ISNULL(o.K_ADD_DEL,''))) AS 追加削除
FROM M_OPTION o
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = o.DANTAI_CD1
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(o.KOMOKU_CD))
WHERE d.MEISYO1 LIKE N'%福生%'
ORDER BY o.DANTAI_CD1, o.COURSE_CD, o.KOMOKU_CD
"@ 80

Q '--- 7. 振り分け表のオプションに当たりそうな項目 (名前でさがす) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 区分, OPTION_GRP AS オプション群
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%乳腺%' OR MEISYO1 LIKE N'%腹部超音波%' OR MEISYO1 LIKE N'%PSA%'
   OR MEISYO1 LIKE N'%BNP%' OR MEISYO1 LIKE N'%CA15%' OR MEISYO1 LIKE N'%CA125%'
   OR MEISYO1 LIKE N'%CEA%' OR MEISYO1 LIKE N'%CA19%' OR MEISYO1 LIKE N'%AFP%'
   OR MEISYO1 LIKE N'%ABC%' OR MEISYO1 LIKE N'%脳梗塞%' OR MEISYO1 LIKE N'%ｱﾚﾙｷﾞｰ%'
   OR MEISYO1 LIKE N'%アレルギ%' OR MEISYO1 LIKE N'%子宮%' OR MEISYO1 LIKE N'%大腸%'
   OR MEISYO1 LIKE N'%じん肺%' OR MEISYO1 LIKE N'%塵肺%' OR MEISYO1 LIKE N'%ﾋｭｰﾑ%'
   OR MEISYO1 LIKE N'%ヒューム%' OR MEISYO1 LIKE N'%溶接%'
ORDER BY KOMOKU_CD
"@ 200

W ''
W '=== 読み方 ==='
W '  2 が本命。0000000477 の予約が「空」や見覚えのないコースなら、コースを作って付け替える。'
W '  ちゃんとしたコースが既に付いているなら、コース作成は不要。'
W '  4・5 で 福生A/B/C (FA/FB/FC) が揃っているかを見る。Cが無ければ作る必要がある。'
W '  6・7 でオプションの受け皿 (D〜O に当たる項目CD) をさがす。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
