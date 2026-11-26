<#
  受診票 OCR v5 の裏面 (家族歴・自覚症状) を健診ナビに入れるための下調べ (check122.ps1)
  マスタと件数だけ。氏名は出さない。DBは読むだけ (SELECT のみ)。
    1. 既往歴の表 (T_KIOU) の列
    2. 続柄の使われ方 (本人 / 父親・母親 など)
    3. 家族の行に入っている病名コードと病名 (受診票の B01〜B17 と同じ番号か)
    4. 家族の行の年齢・治療の入り方、連番の振り方
    5. 自覚症状の選択肢 (受診票の J01〜J34 と同じ番号か)
    6. 病名・続柄のマスタらしい表
    7. 会期 (既定 2026/11/22〜24) の予約に、自覚症状の枠 (017107A〜C) があるか
  check122.bat → r_check122.txt
#>
[CmdletBinding()]
param([string]$From = '2026/11/22', [string]$To = '2026/11/24')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check122.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受診票 OCR v5 の裏面を入れるための下調べ ($From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W '  ★マスタと件数だけ。氏名は出していない。DBは読むだけ。'
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; Start-Process notepad.exe $out; return }
$SELF = "(LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.ZOKUGARA),''))) IN ('', '0'))"

Q '--- 1. 既往歴の表 (T_KIOU) の列 ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型, c.CHARACTER_MAXIMUM_LENGTH AS 長さ, c.IS_NULLABLE AS 空可 FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_KIOU' ORDER BY c.ORDINAL_POSITION
"@ 40

Q '--- 2. ★続柄の使われ方 (コード・名前ごとの行数) ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.ZOKUGARA),''))) + ']' AS 続柄CD, '[' + LTRIM(RTRIM(ISNULL(k.ZOKUGARAMEI,''))) + ']' AS 続柄名, COUNT(*) AS 行数, COUNT(DISTINCT k.KOJIN_ID) AS 人数
FROM T_KIOU k GROUP BY LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.ZOKUGARA),''))), LTRIM(RTRIM(ISNULL(k.ZOKUGARAMEI,''))) ORDER BY 3 DESC
"@ 40

Q '--- 3. ★家族 (本人でない) の行の病名コードと病名 (行数の多い順) ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) + ']' AS 病名CD, LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) AS 病名, COUNT(*) AS 行数
FROM T_KIOU k WHERE NOT $SELF
GROUP BY LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))), LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) ORDER BY 3 DESC
"@ 80

Q '--- 3b. 比べる: 本人の行の病名コード 1〜18 (家族と同じ番号が本人にも使われているか) ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) + ']' AS 病名CD, LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) AS 病名, COUNT(*) AS 行数
FROM T_KIOU k WHERE $SELF AND LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) IN ('1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','01','02','03','04','05','06','07','08','09')
GROUP BY LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))), LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) ORDER BY 1, 3 DESC
"@ 80

Q '--- 4. 家族の行の年齢・治療・備考の入り方 ---' @"
SELECT CASE WHEN $SELF THEN N'本人' ELSE N'家族' END AS 種類, COUNT(*) AS 行数,
       SUM(CASE WHEN k.AGE IS NULL THEN 0 ELSE 1 END) AS 年齢あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.CHIRYO),''))) IN ('', '0') THEN 0 ELSE 1 END) AS 治療CDあり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) = '' THEN 0 ELSE 1 END) AS 治療名あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.BIKOU,''))) = '' THEN 0 ELSE 1 END) AS 備考あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) = '' THEN 0 ELSE 1 END) AS 病名CDあり
FROM T_KIOU k GROUP BY CASE WHEN $SELF THEN N'本人' ELSE N'家族' END
"@ 10

Q '--- 4b. 連番の振り方: 本人と家族で連番が続き番号か (家族の行がある人。連番の最小・最大) ---' @"
SELECT x.種類, MIN(x.R) AS 連番最小, MAX(x.R) AS 連番最大, COUNT(*) AS 行数
FROM (SELECT CASE WHEN $SELF THEN N'本人' ELSE N'家族' END AS 種類, k.RENBAN AS R FROM T_KIOU k
      WHERE k.KOJIN_ID IN (SELECT k2.KOJIN_ID FROM T_KIOU k2 WHERE NOT (LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k2.ZOKUGARA),''))) IN ('', '0')))) x
GROUP BY x.種類
"@ 10

Q '--- 4c. 同じ人・同じ連番の行が 2 つ以上ある組 (0 のはず。主キーの確認) ---' @"
SELECT COUNT(*) AS 組の数 FROM (SELECT k.KOJIN_ID, k.RENBAN FROM T_KIOU k GROUP BY k.KOJIN_ID, k.RENBAN HAVING COUNT(*) > 1) x
"@ 5

Q '--- 4d. 家族の行で、同じ病名に続柄が 2 つ以上ある例があるか (1 病名 1 行か、続柄ごとに 1 行か) ---' @"
SELECT x.N AS 同じ人の同じ病名の家族の行数, COUNT(*) AS 組の数
FROM (SELECT k.KOJIN_ID, LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) AS B, COUNT(*) AS N FROM T_KIOU k WHERE NOT $SELF GROUP BY k.KOJIN_ID, LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) ) x
GROUP BY x.N ORDER BY 1
"@ 10

Q '--- 5. ★自覚症状の選択肢 (017107A の選択肢。受診票の J01〜J34 と同じ番号か) ---' @"
SELECT LTRIM(RTRIM(s.SYOKEN_CD)) AS 選択肢の表, LTRIM(RTRIM(s.KEKKA_CD)) AS 番号, LTRIM(RTRIM(ISNULL(s.SYOKEN,''))) AS 表示, LTRIM(RTRIM(ISNULL(s.HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 s WHERE s.SYOKEN_CD IN (SELECT m.SYOKEN_CD FROM T_KOMOKU m WHERE m.KOMOKU_CD = '017107A')
ORDER BY CASE WHEN ISNUMERIC(s.KEKKA_CD) = 1 THEN CONVERT(int, s.KEKKA_CD) ELSE 9999 END, s.KEKKA_CD
"@ 60

Q '--- 6. 病名・続柄・家族のマスタらしい表 (名前に BYOMEI・ZOKU・KAZOKU・KIOU が入る表と列) ---' @"
SELECT c.TABLE_NAME AS 表, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS c
WHERE (c.TABLE_NAME LIKE '%BYOMEI%' OR c.TABLE_NAME LIKE '%ZOKU%' OR c.TABLE_NAME LIKE '%KAZOKU%' OR c.TABLE_NAME LIKE '%KIOU%'
    OR c.COLUMN_NAME LIKE '%ZOKUGARA%' OR c.COLUMN_NAME LIKE '%KAZOKU%') AND c.TABLE_NAME <> 'T_KIOU'
ORDER BY 1, c.ORDINAL_POSITION
"@ 80

Q '--- 7. 会期の予約に自覚症状の枠 (017107A〜C) があるか (コースごと) ---' @"
SELECT s.D_KENSIN AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コースCD, COUNT(*) AS 予約数,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND k.KOMOKU_CD = '017107A') THEN 1 ELSE 0 END) AS 枠A,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND k.KOMOKU_CD = '017107C') THEN 1 ELSE 0 END) AS 枠C,
       SUM(CASE WHEN ISNULL(s.UKE_NO_KENSA, 0) <> 0 THEN 1 ELSE 0 END) AS 受付番号あり
FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'
GROUP BY s.D_KENSIN, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) ORDER BY 1, 2
"@ 80

Q '--- 7b. 会期の受付番号は 3 日で通し番号か (日ごとの最小・最大。重なっていなければ通し) ---' @"
SELECT s.D_KENSIN AS 受診日, COUNT(*) AS 受付番号あり, MIN(s.UKE_NO_KENSA) AS 最小, MAX(s.UKE_NO_KENSA) AS 最大
FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND ISNULL(s.UKE_NO_KENSA, 0) <> 0
GROUP BY s.D_KENSIN ORDER BY 1
"@ 10

W ''
W '=== 読み方 ==='
W '  2・3: 家族歴は T_KIOU に「続柄つきの行」で入っているはず。続柄のコードと名前、病名コードが受診票の Z-1〜Z-4・B01〜B17 と合うかを見る。'
W '  4〜4d: 家族の行に入れる列と、連番の振り方 (本人の行と通しか)。'
W '  5: 自覚症状の番号が受診票の J01〜J34 と同じなら、いまの対応表 (JCODE) で正しく入る。'
W '  7: 枠 A が 0 のコースは、自覚症状が「枠なし」になって入らない。7b: 受付番号が日ごとに重なっていれば、ファイルを受診日ごとに分ける必要がある。'
Start-Process notepad.exe $out
