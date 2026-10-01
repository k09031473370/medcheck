<#
  受付アプリの番号を健診ナビに入れられるか確かめる (check67.ps1)
  check67.bat をダブルクリックすると実行され、結果 r_check67.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  背景
    受付番号は当日会場の受付アプリが発行する確定番号。健診ナビはそれを受け取るだけ。
    受付アプリ側は 受診者ID(P0001形式)・会社名・氏名 しか持っておらず、
    生年月日・カナ・健診ナビの予約IDを持っていない。
    → 突き合わせは 受診日 + 氏名 (+ 事業所名) になる。それが成立するかを見る。

  ここで見たいこと
    ・10/2 の受診者が健診ナビに何人入っているか
    ・同姓同名がいないか (いると氏名で突き合わせできない)
    ・事業所名が受付アプリ側の「会社名」と一致しそうか
    ・今の受付番号 (UKE_NO_KENSA) が空か
#>
[CmdletBinding()]
param([string]$Ymd = '2026/10/02')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check67.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受付番号を入れられるか ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. $Ymd の受診者数と、団体・コースの内訳 ---" @"
SELECT LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) <> '' THEN 1 ELSE 0 END) AS 受付番号あり
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.DANTAI_CD1 = s.DANTAI_CD1 AND c.COURSE_CD = s.COURSE_CD
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
GROUP BY s.DANTAI_CD1, d.MEISYO1, s.COURSE_CD, c.MEISYO
ORDER BY COUNT(*) DESC
"@ 60

Q "--- 2. ★同姓同名がいないか ($Ymd の中で) ---" @"
SELECT LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名, COUNT(*) AS 件数,
       MIN(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 事業所1,
       MAX(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 事業所2
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
GROUP BY g.KANJI_SIMEI HAVING COUNT(*) > 1
ORDER BY COUNT(*) DESC
"@ 40

Q "--- 3. ★空白を詰めた氏名でも同姓同名がいないか (突き合わせは空白を無視するため) ---" @"
SELECT REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))), ' ', ''), N'　', '') AS 氏名詰め,
       COUNT(*) AS 件数
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
GROUP BY REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))), ' ', ''), N'　', '')
HAVING COUNT(*) > 1 ORDER BY COUNT(*) DESC
"@ 40

Q "--- 4. $Ymd の受診者 先頭40人 (氏名・事業所・コース・今の受付番号) ---" @"
SELECT TOP 40 LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD,
       '[' + LTRIM(RTRIM(ISNULL(s.UKE_NO_KENSA,''))) + ']' AS 受付番号,
       s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
ORDER BY d.MEISYO1, g.KANA_SIMEI
"@ 50

Q "--- 5. $Ymd に関わる事業所の一覧 (受付アプリの「会社名」と見比べる) ---" @"
SELECT LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所名, COUNT(*) AS 人数
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
GROUP BY s.DANTAI_CD1, d.MEISYO1 ORDER BY COUNT(*) DESC
"@ 80

Q "--- 6. 取消になっている人 (F_TORIKESI=1) ---" @"
SELECT COUNT(*) AS 取消件数 FROM T_KENSIN WHERE D_KENSIN = '$Ymd' AND F_TORIKESI <> 0
"@ 10

Q "--- 7. $Ymd のコースに入っている便潜血の項目 (便本数の受け皿) ---" @"
SELECT DISTINCT LTRIM(RTRIM(s.COURSE_CD)) AS コースCD, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目
FROM T_KENSIN s
JOIN T_COURSE2 c2 ON c2.DANTAI_CD1 = s.DANTAI_CD1 AND c2.COURSE_CD = s.COURSE_CD
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND k.MEISYO1 LIKE N'%便潜血%'
ORDER BY 1, 2
"@ 40

W ''
W '=== 読み方 ==='
W '  2・3 が本命。0件なら 受診日+氏名 で突き合わせできる。'
W '  1件でもあれば、その人だけ事業所名も使うか、手で指定することになる。'
W '  5 の事業所名が、受付アプリのCSVの「会社名」と同じ書き方かを見比べてほしい。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
