<#
  アルブミン・総ビリルビンの項目コードと枠の有無を調べる (check83.ps1)
  check83.bat をダブルクリックすると実行され、結果 r_check83.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  SRL の血液ファイル(53列)にはアルブミン(7列目)と総ビリルビン(46列目)の値が入っているが、
  form\mapping_srl53.csv に健診ナビ側の項目コード(KOMOKU_CD)が書いてないので
  プレビューで「項目CD未設定」になり、全員まだ書き込んでいない。
    1) 健診ナビの項目マスタに アルブミン / ビリルビン の項目があるか、そのコードは何か
    2) 9/3・9/9・9/10 の人のコースにその枠があるか (枠が無ければ入れようがない)
  を見て、mapping_srl53.csv にコードを書くかどうかを決めるためのもの。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check83.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== アルブミン・総ビリルビンの項目コードと枠 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 項目マスタで名前に アルブミン / ビリルビン / ALB / BIL が付く項目 ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名,
       LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分, LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見CD
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%アルブミン%' OR MEISYO1 LIKE N'%ｱﾙﾌﾞﾐﾝ%' OR MEISYO1 LIKE N'%ビリルビン%' OR MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%'
   OR MEISYO1 LIKE N'%ALB%' OR MEISYO1 LIKE N'%BIL%' OR MEISYO1 LIKE N'%T-Bil%' OR MEISYO1 LIKE N'%A/G%'
ORDER BY KOMOKU_CD
"@ 40

Q '--- 2. 9/3・9/9・9/10 の人で、その項目の枠を持っている人数 (0 ならコースに枠が無い) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名,
       COUNT(DISTINCT s.PK_SEQ) AS 枠を持つ人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値が入っている数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND (m.MEISYO1 LIKE N'%アルブミン%' OR m.MEISYO1 LIKE N'%ｱﾙﾌﾞﾐﾝ%' OR m.MEISYO1 LIKE N'%ビリルビン%' OR m.MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%'
    OR m.MEISYO1 LIKE N'%ALB%' OR m.MEISYO1 LIKE N'%BIL%')
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY 1
"@ 20

Q '--- 3. 参考: 8月以降の受診者全体で、その項目の枠を持っている人数 (他のコースで使っているか) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名,
       COUNT(DISTINCT s.PK_SEQ) AS 枠を持つ人数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/08/01'
  AND (m.MEISYO1 LIKE N'%アルブミン%' OR m.MEISYO1 LIKE N'%ｱﾙﾌﾞﾐﾝ%' OR m.MEISYO1 LIKE N'%ビリルビン%' OR m.MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%'
    OR m.MEISYO1 LIKE N'%ALB%' OR m.MEISYO1 LIKE N'%BIL%')
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY 1
"@ 20

W ''
W '=== 読み方 ==='
W '  1: 健診ナビにアルブミン・総ビリルビンの項目があるか。ここに出なければ健診ナビ側に項目が無い。'
W '  2: 石屋の人のコースにその枠があるか。「枠を持つ人数」が 0 なら、コースに枠が無いので'
W '     値を持っていても入れる場所が無い (結果票にも載らない)。この場合は何もしなくてよい。'
W '     人数が出ていれば、その 項目CD を mapping_srl53.csv の 7 行目(アルブミン)・46 行目(総ビリルビン)に'
W '     書けば次回の SRL 取込で入る。書く前に私に項目CDを教えてください。'
W '  3: 他のコース (福生など) で使っている項目かの参考。'
notepad $out
