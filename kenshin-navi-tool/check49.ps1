<#
  既往歴 (T_KIOU) の実データを見る (check49.ps1)
  check49.bat をダブルクリックすると実行され、結果 r_check49.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  既往歴は受診ごと(T_KENSA)ではなく、個人ごと(T_KIOU / KOJIN_ID)に入っている。
  どの列に何を入れるのが正しいかを、既に入っている人の実例から読み取る。

  特に確かめたいこと
    ・ZOKUGARA (続柄) … 本人と家族をどう分けているか
    ・BYOMEI_CD        … 病名コードが要るのか、文字だけでよいのか
    ・RENBAN           … 1から振るのか、通し番号なのか
    ・T_KIOU_HANTEI が病名を文字で照合しているか (判定への影響)
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check49.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 既往歴(T_KIOU)の実データ $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★T_KIOU の行数と、続柄の内訳 ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.ZOKUGARA,''))) + ']'    AS 続柄CD,
       '[' + LTRIM(RTRIM(ISNULL(k.ZOKUGARAMEI,''))) + ']' AS 続柄名,
       COUNT(*) AS 行数, COUNT(DISTINCT k.KOJIN_ID) AS 人数
FROM T_KIOU k
GROUP BY k.ZOKUGARA, k.ZOKUGARAMEI
ORDER BY 3 DESC
"@ 30

Q '--- 2. ★★実例: 既往歴を持っている人 (先頭30行) ---' @"
SELECT TOP 30 g.KANJI_SIMEI AS 氏名, k.KOJIN_ID, k.RENBAN,
       '[' + LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) + ']'      AS 病名,
       '[' + LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) + ']'   AS 病名CD,
       '[' + LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.AGE),''))) + ']' AS 年齢,
       '[' + LTRIM(RTRIM(ISNULL(k.CHIRYO,''))) + ']'      AS 治療CD,
       '[' + LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) + ']'   AS 治療名,
       '[' + LTRIM(RTRIM(ISNULL(k.ZOKUGARA,''))) + ']'    AS 続柄CD,
       '[' + LTRIM(RTRIM(ISNULL(k.ZOKUGARAMEI,''))) + ']' AS 続柄名,
       '[' + LTRIM(RTRIM(ISNULL(k.BIKOU,''))) + ']'       AS 備考
FROM T_KIOU k
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = k.KOJIN_ID
ORDER BY k.KOJIN_ID, k.RENBAN
"@ 35

Q '--- 3. ★病名の一覧 (どんな文字で入っているか) ---' @"
SELECT TOP 60 '[' + LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) + ']' AS 病名,
       '[' + LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) + ']' AS 病名CD,
       COUNT(*) AS 件数
FROM T_KIOU k
GROUP BY k.BYOMEI, k.BYOMEI_CD
ORDER BY 3 DESC
"@ 65

Q '--- 4. ★治療の区分 (CHIRYO / CHIRYOMEI) ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.CHIRYO,''))) + ']'    AS 治療CD,
       '[' + LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) + ']' AS 治療名,
       COUNT(*) AS 件数
FROM T_KIOU k
GROUP BY k.CHIRYO, k.CHIRYOMEI
ORDER BY 3 DESC
"@ 25

Q "--- 5. ★$Ymd の88人が既に既往歴を持っているか ---" @"
SELECT CASE WHEN x.件数 > 0 THEN N'既往歴あり' ELSE N'なし' END AS 区分,
       COUNT(*) AS 人数
FROM (
  SELECT s.KOJIN_ID, (SELECT COUNT(*) FROM T_KIOU k WHERE k.KOJIN_ID = s.KOJIN_ID) AS 件数
  FROM T_KENSIN s WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
) x
GROUP BY CASE WHEN x.件数 > 0 THEN N'既往歴あり' ELSE N'なし' END
"@ 10

Q '--- 6. ★判定への影響: T_KIOU_HANTEI の病名と、実際の既往歴の病名が一致するか ---' @"
SELECT h.BYOMEI AS 判定表の病名, h.KOMOKU_CD AS 影響する項目, h.HANTEI_KIGO AS つく判定,
       (SELECT COUNT(*) FROM T_KIOU k WHERE LTRIM(RTRIM(k.BYOMEI)) = LTRIM(RTRIM(h.BYOMEI))) AS 実データ件数
FROM T_KIOU_HANTEI h
GROUP BY h.BYOMEI, h.KOMOKU_CD, h.HANTEI_KIGO
ORDER BY 4 DESC, 1
"@ 80

Q '--- 7. 既往歴が結果報告書に出ている人の例 (帳票の**既往歴_状況N に何が出るか) ---' @"
SELECT TOP 10 g.KANJI_SIMEI AS 氏名, k.RENBAN,
       LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) + ' / ' +
       LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) + ' / ' +
       LTRIM(RTRIM(ISNULL(CONVERT(varchar(10), k.AGE),''))) AS 病名_治療_年齢
FROM T_KIOU k
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = k.KOJIN_ID
WHERE LTRIM(RTRIM(ISNULL(k.ZOKUGARA,''))) IN ('', '0', '1')
ORDER BY k.KOJIN_ID, k.RENBAN
"@ 15

W ''
W '=== 読み方 ==='
W '  1 と 2 で、本人の既往歴と家族歴をどう分けているかが分かります。'
W '     ZOKUGARA が「本人」を表す値を確かめないと、家族歴として登録してしまいます。'
W ''
W '  3 で病名が「文字だけ」なのか「コード付き」なのかが分かります。'
W '     BYOMEI_CD が全部空なら、CSVの文言をそのまま入れられます。'
W ''
W '  4 の治療区分は、CSVには無い情報です。空のままでよいかを判断します。'
W ''
W '  5 で 8/21 の人に既に既往歴があるかが分かります。'
W '     あるなら、二重登録にならないよう気をつける必要があります。'
W ''
W '  6 が重要です。「実データ件数」が多い病名は、既往歴を入れると'
W '     自動判定でその項目が指定の判定に変わります。'
W '     CSVの文言 (高血圧・糖尿病など) が一致すれば、同じことが起きます。'

notepad $out
