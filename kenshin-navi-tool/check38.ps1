<#
  受付番号まわりが今どうなっているかを調べる (check38.ps1)
  check38.bat をダブルクリックすると実行され、結果 r_check38.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  結果入力の一覧で受付Noが「未登録」と出るが、自動判定は通っている。
  受付処理は済んでいるはずなのに、どこが空なのかを確かめます。

  「予約No」がどの列かが分かっていないので、候補をまとめて並べます。
  画面の予約No (85,86,87…) と同じ数字が並んでいる列が正解です。
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

# 予約Noの候補になりそうな列をまとめて出す
$cand = @"
       '[' + ISNULL(s.SEQ1, '') + ']'                              AS SEQ1,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), '') + ']' AS 受付No,
       '[' + ISNULL(s.S_UKE_NO, '') + ']'                          AS S_UKE_NO,
       '[' + ISNULL(CONVERT(varchar(20), s.WAKU_NO), '') + ']'      AS WAKU_NO,
       '[' + ISNULL(CONVERT(varchar(20), s.BAR_CODE), '') + ']'     AS BAR_CODE,
       '[' + ISNULL(s.JUSIN_KEN_NO, '') + ']'                      AS JUSIN_KEN_NO,
       '[' + ISNULL(s.OCR_CODE, '') + ']'                          AS OCR_CODE,
       s.F_UKETUKE AS 受付F,
       (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付行数,
       (SELECT TOP 1 '[' + ISNULL(q.KEN_NO,'') + ']' FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付KEN_NO,
       '[' + ISNULL(g.KOJIN_NO, '') + ']' AS カルテNo,
       s.PK_SEQ
"@

"=== 受付番号まわりの調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★8/21 ($Ymd) 先頭30人。画面の予約No(1,2,3…)と同じ列を探す ---" @"
SELECT TOP 30 g.KANJI_SIMEI AS 氏名,
$cand
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
ORDER BY s.SEQ1
"@ 35

Q "--- 2. ★くらべる: 院内の日 ($Ymd2) 先頭15人 ---" @"
SELECT TOP 15 g.KANJI_SIMEI AS 氏名,
$cand
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd2' AND s.F_TORIKESI = 0
ORDER BY s.SEQ1
"@ 20

Q '--- 3. ★まとめ: どの列が空か (2日分) ---' @"
SELECT x.受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN x.受付No = ''  THEN 1 ELSE 0 END) AS 受付No空,
       SUM(CASE WHEN x.SEQ1 = ''    THEN 1 ELSE 0 END) AS SEQ1空,
       SUM(CASE WHEN x.SUKE = ''    THEN 1 ELSE 0 END) AS S_UKE_NO空,
       SUM(CASE WHEN x.WAKU = ''    THEN 1 ELSE 0 END) AS WAKU_NO空,
       SUM(CASE WHEN x.受付行数 = 0 THEN 1 ELSE 0 END) AS 受付行なし,
       SUM(CASE WHEN x.受付KEN = '' THEN 1 ELSE 0 END) AS 受付KEN_NO空
FROM (
  SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
         LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), ''))) AS 受付No,
         LTRIM(RTRIM(ISNULL(s.SEQ1, '')))                              AS SEQ1,
         LTRIM(RTRIM(ISNULL(s.S_UKE_NO, '')))                          AS SUKE,
         LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.WAKU_NO), '')))      AS WAKU,
         (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ)   AS 受付行数,
         LTRIM(RTRIM(ISNULL((SELECT TOP 1 q.KEN_NO FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ), ''))) AS 受付KEN
  FROM T_KENSIN s
  WHERE s.D_KENSIN IN ('$Ymd', '$Ymd2') AND s.F_TORIKESI = 0
) x
GROUP BY x.受診日
ORDER BY 1
"@ 10

Q '--- 4. 受付済みの人の T_KANJA_G の中身 (8/21 先頭10人) ---' @"
SELECT TOP 10 g.KANJI_SIMEI AS 氏名,
       '[' + ISNULL(q.KEN_NO, '') + ']'  AS KEN_NO,
       '[' + ISNULL(q.KEN_YMD, '') + ']' AS KEN_YMD,
       q.PK_SEQ
FROM T_KANJA_G q
JOIN T_KENSIN s   ON s.PK_SEQ = q.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
ORDER BY q.KEN_NO
"@ 15

W ''
W '=== 読み方 ==='
W '  1 で、画面の予約No (85,86,87…) と同じ数字が並んでいる列が「予約No」です。'
W '     SEQ1 は 受診日8桁+予約No4桁 なので、SEQ1 の末尾4桁でも見当がつきます。'
W ''
W '  「受付No」が [] で「受付行数」が 1 なら、'
W '     受付処理は済んでいるが T_KENSIN の受付番号だけが空、ということです。'
W '     その場合は予約Noを入れても筋が通ります。'
W ''
W '  「受付KEN_NO」に値が入っているなら、健診ナビはそちらを本物として持っています。'
W '     T_KENSIN だけ書き換えると食い違うので、両方そろえる必要があります。'
W ''
W '  2 の院内の日とくらべて、どの列が埋まっていれば「正常」なのかを見ます。'

notepad $out
