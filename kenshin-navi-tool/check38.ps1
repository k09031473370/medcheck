<#
  受付番号まわりが今どうなっているかを調べる (check38.ps1)
  check38.bat をダブルクリックすると実行され、結果 r_check38.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  結果入力の一覧で受付Noが「未登録」と出るが、自動判定は通っている。
  受付処理は済んでいるはずなのに、どこが空なのかを確かめます。

  くらべる相手として、普通に表示できている院内の日も一緒に出します。
#>
[CmdletBinding()]
param(
    [string]$Ymd  = '2026/08/21',   # 調べたい日 (巡回・城西学園)
    [string]$Ymd2 = '2026/08/20'    # くらべる日 (院内)
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check38.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 受付番号まわりの調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KENSIN の列 (予約Noがどの列かを見る) ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KENSIN'
ORDER BY ORDINAL_POSITION
"@ 90

Q '--- 2. ★8/21 の受付まわり (先頭30人) ---' @"
SELECT TOP 30
       g.KANJI_SIMEI AS 氏名,
       '[' + ISNULL(CONVERT(varchar(20), s.YOYAKU_NO), '') + ']'    AS 予約No,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), '') + ']' AS 受付No,
       '[' + ISNULL(s.SEQ1, '') + ']'                              AS SEQ1,
       s.F_UKETUKE AS 受付F,
       (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付行数,
       (SELECT TOP 1 '[' + ISNULL(q.KEN_NO,'') + ']' FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付KEN_NO,
       '[' + ISNULL(g.KOJIN_NO, '') + ']' AS カルテNo,
       s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
ORDER BY s.YOYAKU_NO
"@ 35

Q '--- 3. ★くらべる: 院内の日 ($Ymd2) の同じ並び (先頭15人) ---' @"
SELECT TOP 15
       g.KANJI_SIMEI AS 氏名,
       '[' + ISNULL(CONVERT(varchar(20), s.YOYAKU_NO), '') + ']'    AS 予約No,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), '') + ']' AS 受付No,
       '[' + ISNULL(s.SEQ1, '') + ']'                              AS SEQ1,
       s.F_UKETUKE AS 受付F,
       (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付行数,
       (SELECT TOP 1 '[' + ISNULL(q.KEN_NO,'') + ']' FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付KEN_NO,
       '[' + ISNULL(g.KOJIN_NO, '') + ']' AS カルテNo,
       s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd2' AND s.F_TORIKESI = 0
ORDER BY s.UKE_NO_KENSA
"@ 20

Q '--- 4. ★まとめ: 空になっているのはどれか ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       COUNT(*) AS 人数,
       SUM(CASE WHEN s.YOYAKU_NO    IS NULL                                  THEN 1 ELSE 0 END) AS 予約No空,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),''))) = '' THEN 1 ELSE 0 END) AS 受付No空,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.SEQ1,''))) = ''                    THEN 1 ELSE 0 END) AS SEQ1空,
       SUM(CASE WHEN NOT EXISTS (SELECT 1 FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) THEN 1 ELSE 0 END) AS 受付行なし
FROM T_KENSIN s
WHERE s.D_KENSIN IN ('$Ymd', '$Ymd2') AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN
ORDER BY 1
"@ 10

Q '--- 5. 予約Noと受付Noが違う人はいるか (院内の直近の日で見る) ---' @"
SELECT TOP 20 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       CONVERT(varchar(20), s.YOYAKU_NO) AS 予約No,
       CONVERT(varchar(20), s.UKE_NO_KENSA) AS 受付No,
       COUNT(*) AS 件数
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/08/01'
  AND s.UKE_NO_KENSA IS NOT NULL
  AND CONVERT(varchar(20), s.YOYAKU_NO) <> CONVERT(varchar(20), s.UKE_NO_KENSA)
GROUP BY s.D_KENSIN, s.YOYAKU_NO, s.UKE_NO_KENSA
ORDER BY 1 DESC, 2
"@ 25

W ''
W '=== 読み方 ==='
W '  2 の「受付No」が [] で、「受付行数」が 1 なら、'
W '     健診ナビの受付処理は済んでいるが T_KENSIN の受付番号だけが空、ということです。'
W '     その場合は予約Noを入れても筋が通ります。'
W ''
W '  2 の「受付KEN_NO」に値が入っているなら、健診ナビはそちらを本物として持っています。'
W '     T_KENSIN だけ書き換えると食い違うので、両方そろえる必要があります。'
W ''
W '  3 の院内の日とくらべて、どの列が埋まっていれば「正常」なのかを見ます。'
W ''
W '  5 に行が出れば、予約Noと受付Noは別物として運用されているということです。'
W '     何も出なければ、この健診ナビでは実質同じ番号を使っています。'

notepad $out
