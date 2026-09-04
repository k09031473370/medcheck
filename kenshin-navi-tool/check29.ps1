<#
  入らない血液項目を調べる (check29.ps1)
  check29.bat をダブルクリックすると実行され、結果 r_check29.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  調べること:
    ① 項目コードが分かっていない5項目のコードを探す
         尿素窒素 / 総蛋白 / 総ビリルビン / ALB(アルブミン) / LAP
    ② そのコードの枠が、城西学園(8/21)の受診者にあるか
         コードが分かっても、枠が無ければ「枠なし」で入らない
    ③ ALPの基準値がIFCC法かJSCC法か
         このファイルはIFCC法(値は60前後)。健診ナビの基準値がJSCC法(100〜330)
         のままだと、正常値を異常と判定してしまう
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check29.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 入らない血液項目の調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# ---------------------------------------------------------------- ①
Q '--- ①-1 項目コードを探す (尿素窒素・総蛋白・総ビリルビン・ALB・LAP・ALP) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名,
       LTRIM(RTRIM(ISNULL(TAN_I,''))) AS 単位
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%尿素窒素%' OR MEISYO1 LIKE N'%BUN%' OR MEISYO1 LIKE N'%ﾌﾞﾄﾞｳ%'
   OR MEISYO1 LIKE N'%総蛋白%'   OR MEISYO1 LIKE N'%総ﾀﾝﾊﾟｸ%' OR MEISYO1 LIKE N'%TP%'
   OR MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%' OR MEISYO1 LIKE N'%ビリルビン%'
   OR MEISYO1 LIKE N'%ｱﾙﾌﾞﾐﾝ%'  OR MEISYO1 LIKE N'%アルブミン%' OR MEISYO1 LIKE N'%ALB%'
   OR MEISYO1 LIKE N'%LAP%'     OR MEISYO1 LIKE N'%ﾛｲｼﾝ%'
   OR MEISYO1 LIKE N'%ALP%'     OR MEISYO1 LIKE N'%ｱﾙｶﾘ%'
ORDER BY KOMOKU_CD
"@ 150

Q '--- ①-2 単位が見当たらないとき用: T_KOMOKU の列名を確認 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KOMOKU'
ORDER BY ORDINAL_POSITION
"@ 60

# ---------------------------------------------------------------- ②
Q '--- ②-1 ★本命: 城西学園(8/21)の受診者が持っている血液の枠 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(DISTINCT k.PK_SEQ) AS 枠のある人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値が入っている人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
  AND (m.MEISYO1 LIKE N'%尿素窒素%' OR m.MEISYO1 LIKE N'%BUN%'
    OR m.MEISYO1 LIKE N'%総蛋白%'   OR m.MEISYO1 LIKE N'%総ﾀﾝﾊﾟｸ%'
    OR m.MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%' OR m.MEISYO1 LIKE N'%ビリルビン%'
    OR m.MEISYO1 LIKE N'%ｱﾙﾌﾞﾐﾝ%'  OR m.MEISYO1 LIKE N'%アルブミン%'
    OR m.MEISYO1 LIKE N'%LAP%'     OR m.MEISYO1 LIKE N'%ALP%' OR m.MEISYO1 LIKE N'%ｱﾙｶﾘ%'
    OR m.MEISYO1 LIKE N'%PSA%'     OR m.MEISYO1 LIKE N'%ﾍﾟﾌﾟｼ%' OR m.MEISYO1 LIKE N'%ﾋﾟﾛﾘ%'
    OR m.MEISYO1 LIKE N'%CEA%'     OR m.MEISYO1 LIKE N'%CA19%'  OR m.MEISYO1 LIKE N'%AFP%'
    OR m.MEISYO1 LIKE N'%CA125%'   OR m.MEISYO1 LIKE N'%CA15%'  OR m.MEISYO1 LIKE N'%SCC%'
    OR m.MEISYO1 LIKE N'%FSH%'     OR m.MEISYO1 LIKE N'%ｴｽﾄﾗ%')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY COUNT(DISTINCT k.PK_SEQ) DESC
"@ 100

Q '--- ②-2 城西学園(8/21)の受診者が持っている枠 全部 (何が入れられるかの全体像) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(DISTINCT k.PK_SEQ) AS 枠のある人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY LTRIM(RTRIM(k.KOMOKU_CD))
"@ 300

# ---------------------------------------------------------------- ③
Q '--- ③-1 ★ALPの基準値 (IFCC法かJSCC法かの判定) ---' @"
SELECT LTRIM(RTRIM(KIJUN_CD)) AS 基準CD, LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD,
       RENBAN AS 連番, SEIBETU AS 性別, HANTEI_KIGO AS 判定,
       SAISYO AS 下限, SAIDAI AS 上限, HYOJI_YO AS 表示, HL
FROM T_KIJUN1
WHERE LTRIM(RTRIM(KOMOKU_CD)) = '068645'
ORDER BY KIJUN_CD, SEIBETU, RENBAN
"@ 60

Q '--- ③-2 参考: 実データのALPの分布 (IFCCなら60前後、JSCCなら200前後) ---' @"
SELECT TOP 20 k.KEKKA AS 値, COUNT(*) AS 件数
FROM T_KENSA k
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = '068645'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY k.KEKKA
ORDER BY COUNT(*) DESC
"@ 30

Q '--- ③-3 ALPの実データの範囲 (数値として集計) ---' @"
SELECT COUNT(*) AS 件数,
       MIN(CAST(k.KEKKA AS float)) AS 最小,
       MAX(CAST(k.KEKKA AS float)) AS 最大,
       AVG(CAST(k.KEKKA AS float)) AS 平均
FROM T_KENSA k
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = '068645'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
  AND ISNUMERIC(k.KEKKA) = 1
"@ 10

W ''
W '=== 読み方 ==='
W '  ①-1 で項目CDが見つかれば、それを対応表に書き足すだけで入るようになります。'
W '       見つからなければ、健診ナビにその検査項目自体が無いということです。'
W ''
W '  ②-1 が本命です。「枠のある人数」が 88 なら、そのまま入ります。'
W '       0 なら、コードがあっても城西のコースに枠が無いので入りません。'
W '       その場合は健診ナビ側でコースに項目を足す必要があります(メーカー領域)。'
W '  ②-2 は城西の人が持っている枠の全部です。この一覧に無い項目は入れられません。'
W ''
W '  ③-1 ALPの上限が 110 前後なら IFCC法 → そのまま取り込んで大丈夫です。'
W '       上限が 330 前後なら JSCC法 → 取り込むと正常値が異常判定になります。'
W '  ③-3 の実データの平均でも判断できます。60前後ならIFCC、200前後ならJSCCです。'

notepad $out
