<#
  SEQ1 が空なのは誰か / SEQ1 は何で決まるか (check25.ps1)
  check25.bat をダブルクリックすると実行され、結果 r_check25.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check24 で分かったこと:
    秋葉さんの行は SEQ1 が空。実データは SEQ1 = 受診日8桁 + 受付番号4桁 が入っている。
    自動判定はこの SEQ1 を Substring で切り出すため、空だとあのエラーになる。
    文字色の違いも、報告書画面の0件も、これで説明がつく。

  ここで確かめる決定的なこと:
    SEQ1 が空なのは「秋葉さんだけ」か「08/21の6人全員」か。

      6人全員が空  → まだ受付処理をしていないだけ。ツールの不具合ではない。
                     本番でも受付さえすれば直る。ツールは今のままでよい。
      秋葉さんだけ空 → ツールの受付番号設定が中途半端な状態を作った。
                     ツール側を直す必要がある。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check25.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== SEQ1 が空なのは誰か $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★決定的: 08/21 の6人の SEQ1 / SEQ2 / 受付番号 / 受付フラグ ---' @"
SELECT s.PK_SEQ, s.KOJIN_ID,
       '[' + ISNULL(s.SEQ1,'NULL') + ']' AS SEQ1,
       '[' + ISNULL(s.SEQ2,'NULL') + ']' AS SEQ2,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),'NULL') + ']' AS 受付番号,
       s.F_UKETUKE AS 受付フラグ
FROM T_KENSIN s
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
ORDER BY s.PK_SEQ
"@ 30

Q '--- 2. SEQ1 は 受診日8桁+受付番号4桁 か (実データで確かめる) ---' @"
SELECT TOP 20 s.PK_SEQ, s.D_KENSIN AS 受診日,
       s.SEQ1, s.SEQ2, s.UKE_NO_KENSA AS 受付番号,
       REPLACE(CONVERT(varchar(10), s.D_KENSIN, 111), '/', '')
         + RIGHT('0000' + CONVERT(varchar(10), s.UKE_NO_KENSA), 4) AS 予想したSEQ1,
       CASE WHEN LTRIM(RTRIM(ISNULL(s.SEQ1,''))) =
                 REPLACE(CONVERT(varchar(10), s.D_KENSIN, 111), '/', '')
                 + RIGHT('0000' + CONVERT(varchar(10), s.UKE_NO_KENSA), 4)
            THEN N'一致' ELSE N'不一致' END AS 判定
FROM T_KENSIN s
WHERE s.D_KENSIN = '2026/08/20' AND s.F_TORIKESI = 0
ORDER BY s.PK_SEQ
"@ 30

Q '--- 3. SEQ1 が空の受診はどれくらいあるか (全体の傾向) ---' @"
SELECT CASE WHEN LTRIM(RTRIM(ISNULL(SEQ1,''))) = '' THEN N'SEQ1が空' ELSE N'SEQ1あり' END AS 区分,
       COUNT(*) AS 件数,
       MIN(D_KENSIN) AS 最古, MAX(D_KENSIN) AS 最新
FROM T_KENSIN
WHERE F_TORIKESI = 0
GROUP BY CASE WHEN LTRIM(RTRIM(ISNULL(SEQ1,''))) = '' THEN N'SEQ1が空' ELSE N'SEQ1あり' END
"@ 10

Q '--- 4. ★重要: SEQ1 が空なのに判定記号が付いている受診はあるか (あれば自動判定はSEQ1無しでも通る) ---' @"
SELECT TOP 20 s.PK_SEQ, s.D_KENSIN AS 受診日, '[' + ISNULL(s.SEQ1,'NULL') + ']' AS SEQ1,
       COUNT(k.KOMOKU_CD) AS 判定付き項目数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(s.SEQ1,''))) = ''
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
GROUP BY s.PK_SEQ, s.D_KENSIN, s.SEQ1
ORDER BY s.D_KENSIN DESC
"@ 30

Q '--- 5. 未来日の予約は SEQ1 が入っているか (受付前は空が普通、という裏付け) ---' @"
SELECT TOP 20 s.D_KENSIN AS 受診日,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.SEQ1,''))) = '' THEN 1 ELSE 0 END) AS SEQ1が空,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(s.SEQ1,''))) <> '' THEN 1 ELSE 0 END) AS SEQ1あり,
       COUNT(*) AS 合計
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/08/01'
GROUP BY s.D_KENSIN
ORDER BY s.D_KENSIN
"@ 40

Q '--- 6. T_KENSA 側も同じか (秋葉さん vs 08/20の人) ---' @"
SELECT N'★秋葉' AS 区分,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(SEQ1,''))) = '' THEN 1 ELSE 0 END) AS SEQ1が空,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(SEQ1,''))) <> '' THEN 1 ELSE 0 END) AS SEQ1あり,
       COUNT(*) AS 行数
FROM T_KENSA WHERE PK_SEQ = 2005094
UNION ALL
SELECT N'08/20平野',
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(SEQ1,''))) = '' THEN 1 ELSE 0 END),
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(SEQ1,''))) <> '' THEN 1 ELSE 0 END),
       COUNT(*)
FROM T_KENSA WHERE PK_SEQ = 2004870
"@ 10

W ''
W '=== 読み方 ==='
W '  1 が答えです。'
W '    6人とも SEQ1 が [] (空) → まだ受付処理をしていないだけ。ツールは正常。'
W '        本番でも、当日 健診ナビの受付画面で受付すれば SEQ1 が入り、自動判定は通ります。'
W '        今回のテストは「受付前に結果だけ入れた」状態なので、自動判定が落ちて当然でした。'
W '    秋葉さんだけ空 → ツールの受付番号設定が中途半端。ツール側を直します。'
W ''
W '  2 で「一致」が並べば、SEQ1 = 受診日8桁 + 受付番号4桁 で確定です。'
W '  4 に行が出れば、SEQ1 が空でも自動判定が通った例があるということなので、'
W '     原因は別にあることになります。0件なら SEQ1 が原因で確定。'
W '  5 で 08/21 以降の予約だけ SEQ1 が空なら、「受付でSEQ1が入る」の裏付けになります。'

notepad $out
