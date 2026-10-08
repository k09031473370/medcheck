<#
  受付番号がまだ入っていない人を調べる (check81.ps1)
  check81.bat をダブルクリックすると実行され、結果 r_check81.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  SRL ファイルで「受付番号を設定」したあと、該当者なしになった人
  (9/9 の 4693 ｱｻﾊﾞ ｷﾖﾔｽ / 4701 ﾌｼﾞｲ ｴｲｲﾁﾛｳ / 4702 ﾉﾑﾗ ﾌﾐｵ) が
  健診ナビ側でどう登録されているか (カナ違い・別の日・未登録) を見るためのもの。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check81.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受付番号が入っていない人 (9/3・9/9・9/10) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 日ごとの人数 (予約数 / 受付番号あり / なし) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 予約数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0 THEN 0 ELSE 1 END) AS 受付番号あり,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0 THEN 1 ELSE 0 END) AS 受付番号なし
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
GROUP BY s.D_KENSIN ORDER BY 1
"@ 10

Q '--- 2. ★受付番号がまだ入っていない人 (カナ・漢字・団体) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.PK_SEQ,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 漢字,
       CONVERT(varchar(10), g.D_BIRTH, 111) AS 生年月日,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND (s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0)
ORDER BY 1, 3
"@ 50

Q '--- 3. 該当者なしだった3人をカナの一部で探す (日付を問わず・取消も含む) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.F_TORIKESI AS 取消, s.UKE_NO_KENSA AS 受付番号,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 漢字,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN >= '2026/08/01'
  AND (g.KANA_SIMEI LIKE N'%ｱｻﾊ%' OR g.KANA_SIMEI LIKE N'%アサバ%' OR g.KANA_SIMEI LIKE N'%アサハ%'
    OR g.KANA_SIMEI LIKE N'%ﾌｼﾞｲ%' OR g.KANA_SIMEI LIKE N'%ﾌｼｲ%' OR g.KANA_SIMEI LIKE N'%フジイ%'
    OR g.KANA_SIMEI LIKE N'%ﾉﾑﾗ%' OR g.KANA_SIMEI LIKE N'%ノムラ%'
    OR g.KANJI_SIMEI LIKE N'%浅葉%' OR g.KANJI_SIMEI LIKE N'%浅羽%' OR g.KANJI_SIMEI LIKE N'%麻場%'
    OR g.KANJI_SIMEI LIKE N'%藤井%' OR g.KANJI_SIMEI LIKE N'%野村%')
ORDER BY 1, 4
"@ 40

Q '--- 4. 9/9 で受付番号が重なっていないか (0 行が正常) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, COUNT(*) AS 人数
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND NOT (s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0)
GROUP BY s.D_KENSIN, s.UKE_NO_KENSA HAVING COUNT(*) > 1 ORDER BY 1, 2
"@ 20

W ''
W '=== 読み方 ==='
W '  1: 受付番号なしが 9/3=0, 9/9=3, 9/10=0 なら、SRL の設定は想定どおり。'
W '  2: 9/9 の 3 人が、SRL の ｱｻﾊﾞ ｷﾖﾔｽ(4693) / ﾌｼﾞｲ ｴｲｲﾁﾛｳ(4701) / ﾉﾑﾗ ﾌﾐｵ(4702) と'
W '     同じ人ならカナの書き方が違うだけ。健診ナビの予約画面でその人に受付番号を手で入れる。'
W '  3: 2 に出てこないときは、別の日に予約されているか、予約が無い。ここに出た行で判断する。'
W '  4: 0 行であること。'
notepad $out
