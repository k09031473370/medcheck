<#
  受診票の住所を健診ナビに入れるための下調べ (check123.ps1)
  列の作りと件数だけ。氏名・住所そのものは出さない。DBは読むだけ (SELECT のみ)。
    1. 個人マスタ2 (T_KOJIN2) の列 (住所・郵便番号・更新日時の欄を探す)
    2. T_KOJIN2 は 1 人 1 行か
    3. 会期 (既定 2026/11/22〜24) の受診者の、住所・郵便番号の入り方 (件数だけ)
    4. 郵便番号の書き方 (ハイフンあり / なし)
    5. 住所1・2・3 の分け方 (文字数と、都道府県・市区町村の先頭だけ)
    6. 住所らしい列がほかの表にもあるか
    7. 住所テーブル (T_JYUSYO) の作り
    8. T_KOJIN2 のトリガ
  check123.bat → r_check123.txt
#>
[CmdletBinding()]
param([string]$From = '2026/11/22', [string]$To = '2026/11/24')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check123.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受診票の住所を入れるための下調べ ($From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W '  ★列の作りと件数だけ。氏名・住所そのものは出していない。DBは読むだけ。'
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; Start-Process notepad.exe $out; return }

# 会期の受診者 (取消でない予約の人)
$WHO = "SELECT DISTINCT s.KOJIN_ID FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'"
function Filled([string]$col) { return "SUM(CASE WHEN LTRIM(RTRIM(ISNULL(j.$col,''))) = '' THEN 0 ELSE 1 END)" }

Q '--- 1. 個人マスタ2 (T_KOJIN2) の列 ---' @"
SELECT c.ORDINAL_POSITION AS 順, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型, c.CHARACTER_MAXIMUM_LENGTH AS 長さ, c.IS_NULLABLE AS 空可
FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_KOJIN2' ORDER BY c.ORDINAL_POSITION
"@ 200

Q '--- 1b. T_KOJIN2 の主キー ---' @"
SELECT kc.CONSTRAINT_NAME AS 制約, kc.COLUMN_NAME AS 列, kc.ORDINAL_POSITION AS 順
FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kc ON kc.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
WHERE tc.TABLE_NAME = 'T_KOJIN2' AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY' ORDER BY kc.ORDINAL_POSITION
"@ 10

Q '--- 2. T_KOJIN2 は 1 人 1 行か (行数と人数が同じなら 1 人 1 行) ---' @"
SELECT COUNT(*) AS 行数, COUNT(DISTINCT j.KOJIN_ID) AS 人数 FROM T_KOJIN2 j
"@ 5

Q '--- 2b. 会期の受診者のうち、T_KOJIN2 に行がある人・無い人 ---' @"
SELECT COUNT(*) AS 会期の人数,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KOJIN2 j WHERE j.KOJIN_ID = w.KOJIN_ID) THEN 1 ELSE 0 END) AS KOJIN2あり,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KOJIN2 j WHERE j.KOJIN_ID = w.KOJIN_ID) THEN 0 ELSE 1 END) AS KOJIN2なし
FROM ($WHO) w
"@ 5

Q '--- 3. ★会期の受診者の住所・郵便番号の入り方 (入っている人数) ---' @"
SELECT COUNT(*) AS 人数,
       $(Filled 'YUUBIN') AS 郵便番号あり, $(Filled 'JYUSYO1') AS 住所1あり, $(Filled 'JYUSYO2') AS 住所2あり, $(Filled 'JYUSYO3') AS 住所3あり,
       $(Filled 'TEL') AS 電話あり,
       $(Filled 'SOUFU_YUUBIN') AS 送付先郵便あり, $(Filled 'SOUFU_JYUSYO1') AS 送付先住所1あり, $(Filled 'SOUFU_JYUSYO2') AS 送付先住所2あり
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO)
"@ 5

Q '--- 3b. 比べる: 全員 (T_KOJIN2 全体) の入り方 ---' @"
SELECT COUNT(*) AS 人数,
       $(Filled 'YUUBIN') AS 郵便番号あり, $(Filled 'JYUSYO1') AS 住所1あり, $(Filled 'JYUSYO2') AS 住所2あり, $(Filled 'JYUSYO3') AS 住所3あり,
       $(Filled 'SOUFU_YUUBIN') AS 送付先郵便あり, $(Filled 'SOUFU_JYUSYO1') AS 送付先住所1あり
FROM T_KOJIN2 j
"@ 5

Q '--- 3c. 会期の受診者: 住所が入っている人・いない人 (受診日ごと) ---' @"
SELECT s.D_KENSIN AS 受診日, COUNT(*) AS 予約数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(j.JYUSYO1,''))) + LTRIM(RTRIM(ISNULL(j.JYUSYO2,''))) = '' THEN 0 ELSE 1 END) AS 住所あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(j.JYUSYO1,''))) + LTRIM(RTRIM(ISNULL(j.JYUSYO2,''))) = '' THEN 1 ELSE 0 END) AS 住所なし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(j.YUUBIN,''))) = '' THEN 1 ELSE 0 END) AS 郵便番号なし
FROM T_KENSIN s LEFT JOIN T_KOJIN2 j ON j.KOJIN_ID = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' GROUP BY s.D_KENSIN ORDER BY 1
"@ 10

Q '--- 4. ★郵便番号の書き方 (全員。形ごとの人数) ---' @"
SELECT x.形, COUNT(*) AS 人数 FROM (
  SELECT CASE WHEN LTRIM(RTRIM(ISNULL(j.YUUBIN,''))) = '' THEN N'(空)'
              WHEN LTRIM(RTRIM(j.YUUBIN)) LIKE '[0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]' THEN N'123-4567 (半角・ハイフンあり)'
              WHEN LTRIM(RTRIM(j.YUUBIN)) LIKE '[0-9][0-9][0-9][0-9][0-9][0-9][0-9]' THEN N'1234567 (半角・ハイフンなし)'
              ELSE N'その他 (全角・桁違いなど) 長さ ' + CONVERT(varchar(5), LEN(LTRIM(RTRIM(j.YUUBIN)))) END AS 形
  FROM T_KOJIN2 j) x GROUP BY x.形 ORDER BY 2 DESC
"@ 20

Q '--- 5. ★住所1・2・3 の文字数 (会期の受診者。入っている人だけ) ---' @"
SELECT N'住所1' AS 欄, COUNT(*) AS 人数, MIN(LEN(LTRIM(RTRIM(j.JYUSYO1)))) AS 最短, AVG(LEN(LTRIM(RTRIM(j.JYUSYO1)))) AS 平均, MAX(LEN(LTRIM(RTRIM(j.JYUSYO1)))) AS 最長
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO1,''))) <> ''
UNION ALL
SELECT N'住所2', COUNT(*), MIN(LEN(LTRIM(RTRIM(j.JYUSYO2)))), AVG(LEN(LTRIM(RTRIM(j.JYUSYO2)))), MAX(LEN(LTRIM(RTRIM(j.JYUSYO2))))
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO2,''))) <> ''
UNION ALL
SELECT N'住所3', COUNT(*), MIN(LEN(LTRIM(RTRIM(j.JYUSYO3)))), AVG(LEN(LTRIM(RTRIM(j.JYUSYO3)))), MAX(LEN(LTRIM(RTRIM(j.JYUSYO3))))
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO3,''))) <> ''
"@ 5

Q '--- 5b. 住所1 の先頭 4 文字 (会期の受診者。都道府県から始まっているか。多い順) ---' @"
SELECT TOP 15 LEFT(LTRIM(RTRIM(j.JYUSYO1)), 4) AS 住所1の先頭4文字, COUNT(*) AS 人数
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO1,''))) <> ''
GROUP BY LEFT(LTRIM(RTRIM(j.JYUSYO1)), 4) HAVING COUNT(*) >= 3 ORDER BY 2 DESC
"@ 15

Q '--- 5c. 住所1 が都道府県だけの人・市区町村まで入っている人 (会期の受診者) ---' @"
SELECT x.形, COUNT(*) AS 人数 FROM (
  SELECT CASE WHEN LTRIM(RTRIM(j.JYUSYO1)) LIKE N'%[都道府県]' AND LEN(LTRIM(RTRIM(j.JYUSYO1))) <= 4 THEN N'都道府県だけ'
              WHEN LTRIM(RTRIM(j.JYUSYO1)) LIKE N'東京都%' OR LTRIM(RTRIM(j.JYUSYO1)) LIKE N'北海道%' OR LTRIM(RTRIM(j.JYUSYO1)) LIKE N'%[府県]%' THEN N'都道府県＋その先'
              ELSE N'都道府県なしで始まる' END AS 形
  FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO1,''))) <> '') x GROUP BY x.形 ORDER BY 2 DESC
"@ 10

Q '--- 5d. 住所2 の先頭 3 文字 (会期の受診者。市区町村から始まっているか。3 人以上のものだけ) ---' @"
SELECT TOP 15 LEFT(LTRIM(RTRIM(j.JYUSYO2)), 3) AS 住所2の先頭3文字, COUNT(*) AS 人数
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO2,''))) <> ''
GROUP BY LEFT(LTRIM(RTRIM(j.JYUSYO2)), 3) HAVING COUNT(*) >= 3 ORDER BY 2 DESC
"@ 15

Q '--- 5e. 住所の数字は全角か半角か (会期の受診者。住所2 に入っている数字) ---' @"
SELECT SUM(CASE WHEN j.JYUSYO2 LIKE '%[0-9]%' THEN 1 ELSE 0 END) AS 半角数字あり,
       SUM(CASE WHEN j.JYUSYO2 LIKE N'%[０-９]%' THEN 1 ELSE 0 END) AS 全角数字あり,
       SUM(CASE WHEN j.JYUSYO2 LIKE N'%丁目%' THEN 1 ELSE 0 END) AS 丁目と書いてある,
       SUM(CASE WHEN j.JYUSYO2 LIKE '%-%' OR j.JYUSYO2 LIKE N'%－%' OR j.JYUSYO2 LIKE N'%ー%' THEN 1 ELSE 0 END) AS ハイフンあり
FROM T_KOJIN2 j WHERE j.KOJIN_ID IN ($WHO) AND LTRIM(RTRIM(ISNULL(j.JYUSYO2,''))) <> ''
"@ 5

Q '--- 6. 住所・郵便番号らしい列がある表 (T_KOJIN2 以外も) ---' @"
SELECT c.TABLE_NAME AS 表, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型, c.CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.COLUMN_NAME LIKE '%JYUSYO%' OR c.COLUMN_NAME LIKE '%JUSYO%' OR c.COLUMN_NAME LIKE '%JUSHO%' OR c.COLUMN_NAME LIKE '%YUUBIN%' OR c.COLUMN_NAME LIKE '%YUBIN%' OR c.COLUMN_NAME LIKE '%ADDR%'
ORDER BY 1, c.ORDINAL_POSITION
"@ 120

Q '--- 7. 住所テーブル (T_JYUSYO) の列 ---' @"
SELECT c.ORDINAL_POSITION AS 順, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型, c.CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_JYUSYO' ORDER BY c.ORDINAL_POSITION
"@ 40

Q '--- 7b. T_JYUSYO の行数 (郵便番号の辞書なら十数万行。受診者の住所なら人数くらい) ---' @"
SELECT COUNT(*) AS 行数 FROM T_JYUSYO
"@ 5

Q '--- 8. T_KOJIN2 に付いているトリガ (あれば、書いたときに別の表も動く) ---' @"
SELECT t.name AS トリガ, t.is_disabled AS 無効, t.is_instead_of_trigger AS INSTEAD_OF FROM sys.triggers t WHERE t.parent_id = OBJECT_ID('T_KOJIN2')
"@ 10

W ''
W '=== 読み方 ==='
W '  1: YUUBIN・JYUSYO1〜3 の型と長さ (受診票の住所が入りきるか)。更新日時・更新者の欄があれば、書くときに合わせる。'
W '  2: 1 人 1 行なら、KOJIN_ID で 1 行を書き換えればよい。2b で「なし」の人がいると、行を足す必要がある。'
W '  3: 会期の人に、もう住所が入っているか。入っている人が多いなら、受診票の住所で上書きしてよいかを決める必要がある。'
W '  4・5: 郵便番号と住所の書き方 (ハイフン・全角半角・1/2/3 の分け方)。健診ナビの画面で入れたものと同じ形にそろえる。'
W '  6: 予約 (T_KENSIN) などにも住所の写しがあれば、そちらも合わせる必要があるかを見る。'
W '  ★氏名・住所そのものは出していません。このまま貼り付けて大丈夫です。'
Start-Process notepad.exe $out
