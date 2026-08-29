<#
  自動判定エラーの原因さがし その4: 崩れた基準行を探す (check23.ps1)
  check23.bat をダブルクリックすると実行され、結果 r_check23.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check22 で分かったこと:
    ・YSA1 は標準の基準セット KIJUN_CD=1 を使う
    ・数値の範囲(SAIDAI/SAISYO)も有効期間(D_YUUKOO_F/T)も「文字列」で入っている
    → 自動判定はこれらの文字列を Substring で切り出す。
      空・桁足らず・形式違いの行が1つでもあると、そこで落ちる。

  ここでやること: 崩れた行を全部の基準テーブルからスキャンする。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check23.txt'

# テスト受診者 秋葉達也 (男性・54才)
$PK = 2005094

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 崩れた基準行さがし $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KIJUN1 で範囲の文字列(SAIDAI/SAISYO)が空か数値でない行 ---' @"
SELECT j.KIJUN_CD, LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       j.RENBAN, j.SEIBETU, j.HANTEI_KIGO,
       '[' + ISNULL(j.SAIDAI,'NULL') + ']' AS SAIDAI,
       '[' + ISNULL(j.SAISYO,'NULL') + ']' AS SAISYO,
       j.NENREI_F, j.NENREI_T
FROM T_KIJUN1 j
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE LTRIM(RTRIM(ISNULL(j.SAIDAI,''))) = ''
   OR LTRIM(RTRIM(ISNULL(j.SAISYO,''))) = ''
   OR ISNUMERIC(LTRIM(RTRIM(j.SAIDAI))) = 0
   OR ISNUMERIC(LTRIM(RTRIM(j.SAISYO))) = 0
ORDER BY j.KOMOKU_CD, j.RENBAN
"@ 100

Q '--- 2. T_KIJUN1 で有効期間(D_YUUKOO_F/T)の形が 9999/99/99 型でない行 ---' @"
SELECT j.KIJUN_CD, LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       j.RENBAN, j.SEIBETU,
       '[' + ISNULL(j.D_YUUKOO_F,'NULL') + ']' AS 有効F,
       '[' + ISNULL(j.D_YUUKOO_T,'NULL') + ']' AS 有効T
FROM T_KIJUN1 j
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE (LTRIM(RTRIM(ISNULL(j.D_YUUKOO_F,''))) <> '' AND LTRIM(RTRIM(j.D_YUUKOO_F)) NOT LIKE '[12][09][0-9][0-9]/[01][0-9]/[0-3][0-9]')
   OR (LTRIM(RTRIM(ISNULL(j.D_YUUKOO_T,''))) <> '' AND LTRIM(RTRIM(j.D_YUUKOO_T)) NOT LIKE '[12][09][0-9][0-9]/[01][0-9]/[0-3][0-9]')
   OR LTRIM(RTRIM(ISNULL(j.D_YUUKOO_F,''))) = ''
   OR LTRIM(RTRIM(ISNULL(j.D_YUUKOO_T,''))) = ''
ORDER BY j.KOMOKU_CD, j.RENBAN
"@ 120

Q '--- 3. T_KIJUN2 で定性値(TEISEI)が空の行、有効期間が崩れている行 ---' @"
SELECT j.KIJUN_CD, LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       j.RENBAN, j.SEIBETU, j.HANTEI_KIGO,
       '[' + ISNULL(j.TEISEI,'NULL') + ']' AS TEISEI,
       '[' + ISNULL(j.D_YUUKOO_F,'NULL') + ']' AS 有効F,
       '[' + ISNULL(j.D_YUUKOO_T,'NULL') + ']' AS 有効T
FROM T_KIJUN2 j
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE LTRIM(RTRIM(ISNULL(j.TEISEI,''))) = ''
   OR LTRIM(RTRIM(ISNULL(j.D_YUUKOO_F,''))) = ''
   OR LTRIM(RTRIM(ISNULL(j.D_YUUKOO_T,''))) = ''
   OR (LTRIM(RTRIM(ISNULL(j.D_YUUKOO_F,''))) <> '' AND LTRIM(RTRIM(j.D_YUUKOO_F)) NOT LIKE '[12][09][0-9][0-9]/[01][0-9]/[0-3][0-9]')
   OR (LTRIM(RTRIM(ISNULL(j.D_YUUKOO_T,''))) <> '' AND LTRIM(RTRIM(j.D_YUUKOO_T)) NOT LIKE '[12][09][0-9][0-9]/[01][0-9]/[0-3][0-9]')
ORDER BY j.KOMOKU_CD, j.RENBAN
"@ 120

Q '--- 4. メタボ判定基準 M_METAKIJUN_R6 の全行 (基準値の文字列に空・崩れが無いか) ---' @"
SELECT METAKIJUN_NO AS NO, METAKIJUN_MEISYO AS 名称,
       '[' + ISNULL(METAKIJUN_HIKAKU,'NULL') + ']' AS 比較,
       '[' + ISNULL(METAKIJUN_KIJUNTI,'NULL') + ']' AS 基準値,
       LTRIM(RTRIM(ISNULL(KOMOKU_CD,''))) AS 項目CD
FROM M_METAKIJUN_R6
ORDER BY METAKIJUN_NO
"@ 80

Q '--- 5. 秋葉さんの値入り項目ごとに、使える基準行が何行あるか (男・54才で0行の項目が怪しい) ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, a.KEKKA AS テストの値,
       (SELECT COUNT(*) FROM T_KIJUN1 j
         WHERE LTRIM(RTRIM(j.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
           AND LTRIM(RTRIM(j.KIJUN_CD)) = '1'
           AND j.SEIBETU IN ('0','1')) AS 数値基準行,
       (SELECT COUNT(*) FROM T_KIJUN2 j
         WHERE LTRIM(RTRIM(j.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
           AND LTRIM(RTRIM(j.KIJUN_CD)) = '1'
           AND j.SEIBETU IN ('0','1')) AS 定性基準行
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
ORDER BY 4 ASC, 5 ASC, 1
"@ 200

Q '--- 6. 参考: 直近で判定記号が付いた受診はいつか (自動判定が最近も動いている証拠) ---' @"
SELECT TOP 15 s.D_KENSIN AS 受診日, COUNT(DISTINCT s.PK_SEQ) AS 判定付き人数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN
ORDER BY s.D_KENSIN DESC
"@ 20

W ''
W '=== 読み方 ==='
W '  1〜3 に行が出たら、それが自動判定を落としている候補です。'
W '     [] で囲って表示しているので、空欄・空白だけ・NULL がひと目で分かります。'
W '     とくに 秋葉さんの値入り項目(身長・体重・血圧・肝機能・脂質・血糖など)と'
W '     同じ項目CDの行が出ていたら、それが本命です。'
W '  4 のメタボ基準に空・崩れがあれば、定期健診の自動判定はメタボ判定も通るのでそこで落ちます。'
W '  5 で「数値基準行 0 かつ 定性基準行 0」の項目は、基準がまったく無いのに値だけある状態。'
W '     判定処理がその項目の基準を探して空振りし、Substring で落ちる形です。'
W '  6 で最近の受診日にも判定が付いていれば、他コースでは自動判定が正常に動いています。'
W '     = 壊れているのは全体ではなく、特定の項目・基準行だけということです。'
W ''
W '  この結果で犯人の行が特定できます。直すのは基準マスタの1行なので、'
W '  健診ナビの「基準値マスタ」画面から手で直せる可能性が高いです(システム改造は不要)。'

notepad $out
