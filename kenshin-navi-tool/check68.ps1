<#
  会場受付アプリの受付番号を健診ナビに入れられるか確かめる (check68.ps1)
  check68.bat をダブルクリックすると実行され、結果 r_check68.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check67 からの変更
    予約日に来るとは限らない (10/2の人が10/6・10/7に回ることがある) ので、
    1日ではなく福生市商工会の全日程をまとめて見る。
    同姓同名のチェックも全日程まとめてでないと意味がない。

  背景
    受付番号は当日会場の受付アプリが発行する確定番号。健診ナビは受け取る側。
    受付アプリは 受診者ID(P0001形式)・会社名・氏名 しか持たず、
    生年月日・カナ・健診ナビの予約IDを持っていない。
    → 突き合わせは 氏名 (+ 事業所名) になる。それが成立するかを見る。
#>
[CmdletBinding()]
param(
    [string]$Name = '福生',          # 団体名にこれが入っているものを福生の案件とみなす
    [string]$From = '2026/09/01',    # 見る期間
    [string]$To   = '2026/12/31'
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check68.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
# 福生の受診を1つにまとめる条件 (団体名に「福生」が入る)
$SCOPE = @"
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'
  AND d.MEISYO1 LIKE N'%$Name%'
"@
"=== 受付番号を入れられるか ($Name / $From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★「$Name」の日程と人数 (全日程をまとめて扱う) ---" @"
SELECT s.D_KENSIN AS 受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) <> '' THEN 1 ELSE 0 END) AS 受付番号あり,
       COUNT(DISTINCT s.DANTAI_CD1) AS 事業所数
$SCOPE
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN
"@ 40

Q "--- 2. ★★全日程まとめて同姓同名がいないか (空白を詰めて比べる) ---" @"
SELECT REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))), ' ', ''), N'　', '') AS 氏名,
       COUNT(*) AS 件数,
       MIN(CONVERT(varchar(10), s.D_KENSIN)) AS 日1, MAX(CONVERT(varchar(10), s.D_KENSIN)) AS 日2,
       MIN(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 事業所1, MAX(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 事業所2
$SCOPE
GROUP BY REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))), ' ', ''), N'　', '')
HAVING COUNT(*) > 1
ORDER BY COUNT(*) DESC
"@ 60

Q "--- 3. 氏名+事業所 でも重なるか (2 で重複があったとき用) ---" @"
SELECT REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))), ' ', ''), N'　', '') AS 氏名,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所, COUNT(*) AS 件数
$SCOPE
GROUP BY REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))), ' ', ''), N'　', ''), d.MEISYO1
HAVING COUNT(*) > 1 ORDER BY COUNT(*) DESC
"@ 40

Q "--- 4. 事業所の一覧 (受付アプリCSVの「会社名」と見比べる) ---" @"
SELECT LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所名, COUNT(*) AS 人数,
       MIN(CONVERT(varchar(10), s.D_KENSIN)) AS 最初の日, MAX(CONVERT(varchar(10), s.D_KENSIN)) AS 最後の日
$SCOPE
GROUP BY s.DANTAI_CD1, d.MEISYO1 ORDER BY COUNT(*) DESC
"@ 120

Q "--- 5. コースの内訳 (受付アプリの A/B/C と対応がつくか) ---" @"
SELECT LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名, COUNT(*) AS 人数
$SCOPE
GROUP BY s.COURSE_CD, c.MEISYO ORDER BY COUNT(*) DESC
"@ 40

Q "--- 6. 受診者 先頭40人 (氏名・カナ・事業所・受診日・コース・今の受付番号) ---" @"
SELECT TOP 40 LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       CONVERT(varchar(10), s.D_KENSIN) AS 受診日,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       '[' + LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) + ']' AS 受付番号, s.PK_SEQ
$SCOPE
ORDER BY d.MEISYO1, g.KANA_SIMEI
"@ 50

Q "--- 7. 取消になっている人 (F_TORIKESI<>0。日程変更の跡が残っていないか) ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN) AS 受診日, COUNT(*) AS 取消件数
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI <> 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND d.MEISYO1 LIKE N'%$Name%'
GROUP BY s.D_KENSIN ORDER BY s.D_KENSIN
"@ 30

Q "--- 8. 便潜血の項目がコースに入っているか (便本数の受け皿) ---" @"
SELECT DISTINCT LTRIM(RTRIM(s.COURSE_CD)) AS コースCD, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
JOIN T_COURSE2 c2 ON c2.DANTAI_CD1 = s.DANTAI_CD1 AND c2.COURSE_CD = s.COURSE_CD
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'
  AND d.MEISYO1 LIKE N'%$Name%' AND k.MEISYO1 LIKE N'%便潜血%'
ORDER BY 1, 2
"@ 40

W ''
W '=== 読み方 ==='
W '  1 で日程が全部そろっているか (10/2・10/6・10/7)。'
W '  2 が本命。0件なら「氏名」だけで全日程まとめて突き合わせできる。'
W '    1件でもあれば 3 を見て、氏名+事業所 で分けられるかを確認する。'
W '  4 の事業所名が、受付アプリCSVの「会社名」と同じ書き方かを見比べてほしい。'
W '  7 に取消があれば、日程変更で予約を作り直した跡かもしれない。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
