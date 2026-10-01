<#
  福生商工会の予約取込の結果を確かめる (check71.ps1)
  check71.bat をダブルクリックすると実行され、結果 r_check71.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  振り分け表(10/1版)の中身 ─ これと合っているかを見る
    合計 325人 / 77社
    10/02 103人 (A 82 / B 13 / C 8)
    10/06 115人 (A 91 / B  2 / C 21 / コースなし 1)
    10/07 107人 (A 76 / B 10 / C 20 / コースなし 1)
    コース A=定期健康診断A B=定期健康診断B C=ミニドック

  見たいこと
    ・何人入ったか、日付と団体は合っているか
    ・コースが何に割り当たったか (FA/FB/FC があるか)
    ・生年月日・カナが入っているか
    ・同じ人が二重に個人登録されていないか
    ・ついでにマイクロン(0000000477)のコース状況
#>
[CmdletBinding()]
param(
    [string]$Micron = '0000000477'
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check71.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
# 福生の受診をまとめる条件
$SCOPE = @"
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND d.MEISYO1 LIKE N'%福生%'
  AND s.D_KENSIN IN ('2026/10/02','2026/10/06','2026/10/07')
"@
"=== 福生の予約取込の結果 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★日付ごとの人数 (振り分け表: 10/2=103 10/6=115 10/7=107 計325) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日, COUNT(*) AS 人数,
       COUNT(DISTINCT s.DANTAI_CD1) AS 事業所数
$SCOPE
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN
"@ 30

Q '--- 2. ★日付×コース (振り分け表: 10/2 A82 B13 C8 / 10/6 A91 B2 C21 / 10/7 A76 B10 C20) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,'(空)'))) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名, COUNT(*) AS 人数,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = s.DANTAI_CD1 AND x.COURSE_CD = s.COURSE_CD) AS 項目数
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.DANTAI_CD1 = s.DANTAI_CD1 AND c.COURSE_CD = s.COURSE_CD
WHERE s.F_TORIKESI = 0 AND d.MEISYO1 LIKE N'%福生%'
  AND s.D_KENSIN IN ('2026/10/02','2026/10/06','2026/10/07')
GROUP BY s.D_KENSIN, s.COURSE_CD, c.MEISYO, s.DANTAI_CD1
ORDER BY s.D_KENSIN, COUNT(*) DESC
"@ 60

Q '--- 3. ★生年月日・カナが入っているか ---' @"
SELECT COUNT(*) AS 人数,
       SUM(CASE WHEN g.D_BIRTH IS NULL THEN 1 ELSE 0 END) AS 生年月日なし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) = '' THEN 1 ELSE 0 END) AS カナなし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.SEIBETU,''))) = '' THEN 1 ELSE 0 END) AS 性別なし
$SCOPE
"@ 20

Q '--- 4. ★★同じ氏名で個人IDが2つ以上ある人 (名寄せ失敗・二重登録の疑い) ---' @"
SELECT REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))),' ',''),N'　','') AS 氏名,
       COUNT(DISTINCT s.KOJIN_ID) AS 個人ID数, COUNT(*) AS 予約数
$SCOPE
GROUP BY REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))),' ',''),N'　','')
HAVING COUNT(DISTINCT s.KOJIN_ID) > 1 OR COUNT(*) > 1
ORDER BY COUNT(*) DESC
"@ 60

Q '--- 5. 事業所ごとの人数 (振り分け表と突き合わせる) ---' @"
SELECT LTRIM(RTRIM(s.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所, COUNT(*) AS 人数
$SCOPE
GROUP BY s.DANTAI_CD1, d.MEISYO1 ORDER BY COUNT(*) DESC
"@ 120

Q '--- 6. 受診者 先頭30人 (中身の確認) ---' @"
SELECT TOP 30 CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ,
       CONVERT(varchar(10), g.D_BIRTH) AS 生年月日,
       LTRIM(RTRIM(ISNULL(g.SEIBETU,''))) AS 性別,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       s.UKE_NO_KENSA AS 受付番号
$SCOPE
ORDER BY s.D_KENSIN, d.MEISYO1, g.KANA_SIMEI
"@ 40

Q '--- 7. 取消 (F_TORIKESI<>0) が増えていないか ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日, COUNT(*) AS 取消件数
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI <> 0 AND d.MEISYO1 LIKE N'%福生%'
  AND s.D_KENSIN IN ('2026/10/02','2026/10/06','2026/10/07')
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN
"@ 30

Q '--- 8. 福生のコース一覧 (FA/FB/FC が揃っているか) ---' @"
SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(DISTINCT c.DANTAI_CD1) AS 団体数,
       MIN((SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD)) AS 項目数最小,
       MAX((SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD)) AS 項目数最大
FROM T_COURSE1 c LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE d.MEISYO1 LIKE N'%福生%'
GROUP BY c.COURSE_CD, c.MEISYO ORDER BY c.COURSE_CD
"@ 40

Q "--- 9. ついでに: マイクロン $Micron の予約が使っているコース ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,'(空)'))) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名, COUNT(*) AS 人数,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = s.DANTAI_CD1 AND x.COURSE_CD = s.COURSE_CD) AS 項目数
FROM T_KENSIN s
LEFT JOIN T_COURSE1 c ON c.DANTAI_CD1 = s.DANTAI_CD1 AND c.COURSE_CD = s.COURSE_CD
WHERE s.DANTAI_CD1 = '$Micron' AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN, s.COURSE_CD, c.MEISYO, s.DANTAI_CD1
ORDER BY s.D_KENSIN, COUNT(*) DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  1・2 の人数が振り分け表と合っているかが第一。'
W '  2 でコースが「(空)」や項目数0なら、コースが付いていないので要対応。'
W '  3 の生年月日なしが多ければ、あとで一括で埋める必要がある (ツールを作れます)。'
W '  4 に行が出たら、同じ人が二重に登録されたか、別人の同姓同名。'
W '  9 でマイクロンのコースが分かる。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
