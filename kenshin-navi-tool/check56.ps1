<#
  視力の判定基準がどこに設定されているかを探す (check56.ps1)
  check56.bat をダブルクリックすると実行され、結果 r_check56.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check47 の結果、視力4項目 (067012/067013/067018/067019) は
  基準値マスタ T_KIJUN2 に行が無かった。
  実際には 1.0以上=A / 0.7〜0.9=B / 0.6以下=F と判定されているので、
  どこかに設定があるはず。それを探す。

  あわせて、ほかの項目 (ALT・LDL・総コレステロール) の基準値マスタを出して、
  「人間ドック学会の C を 健診ナビの D に割り当てている」ことを確かめる。
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check56.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 視力の判定基準はどこにあるか $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 名前に KIJUN / HANTEI を含むテーブル一覧 ---' @"
SELECT t.TABLE_NAME AS テーブル, COUNT(*) AS 列数
FROM INFORMATION_SCHEMA.TABLES t
JOIN INFORMATION_SCHEMA.COLUMNS c ON c.TABLE_NAME = t.TABLE_NAME
WHERE t.TABLE_TYPE = 'BASE TABLE'
  AND (t.TABLE_NAME LIKE '%KIJUN%' OR t.TABLE_NAME LIKE '%HANTEI%')
GROUP BY t.TABLE_NAME ORDER BY 1
"@ 40

Q '--- 2. ★KOMOKU_CD 列を持つ表で、視力4項目の行がどこにあるか ---' @"
SELECT c.TABLE_NAME AS テーブル
FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE c.COLUMN_NAME = 'KOMOKU_CD'
ORDER BY 1
"@ 80

Q '--- 3. ★★T_KIJUN2 に視力まわりの行があるか (067012〜067019 / 600040 / 067%) ---' @"
SELECT KIJUN_CD, LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, RENBAN, SEIBETU, HANTEI_KIGO,
       TEISEI, HL, HYOJI_YO, KIJUN_MEISYO
FROM T_KIJUN2
WHERE LTRIM(RTRIM(KOMOKU_CD)) LIKE '067%' OR LTRIM(RTRIM(KOMOKU_CD)) = '600040'
ORDER BY KOMOKU_CD, KIJUN_CD, RENBAN
"@ 80

Q '--- 4. T_KIJUN2 の KIJUN_CD の種類 (どの基準セットがあるか) ---' @"
SELECT KIJUN_CD, MAX(KIJUN_MEISYO) AS 基準名, COUNT(*) AS 行数,
       COUNT(DISTINCT LTRIM(RTRIM(KOMOKU_CD))) AS 項目数
FROM T_KIJUN2
GROUP BY KIJUN_CD ORDER BY 1
"@ 40

Q '--- 5. ★確認: ALT(GPT) 068637 の基準値マスタ (学会C→健診ナビD かどうか) ---' @"
SELECT KIJUN_CD, RENBAN, SEIBETU, HANTEI_KIGO, TEISEI, HL, HYOJI_YO, KIJUN_MEISYO
FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '068637'
ORDER BY KIJUN_CD, RENBAN
"@ 60

Q '--- 6. ★確認: LDL 068709 の基準値マスタ ---' @"
SELECT KIJUN_CD, RENBAN, SEIBETU, HANTEI_KIGO, TEISEI, HL, HYOJI_YO, KIJUN_MEISYO
FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '068709'
ORDER BY KIJUN_CD, RENBAN
"@ 60

Q '--- 7. T_KOMOKU の視力4項目 (計算式・基準値の列に何か入っていないか) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1,
       '[' + ISNULL(K_KIJUNTI,'') + ']' AS K_KIJUNTI,
       '[' + ISNULL(K_SEIBETU,'') + ']' AS K_SEIBETU,
       '[' + ISNULL(CALC_FORM,'') + ']' AS CALC_FORM,
       CALC_ORDER, S_1, S_2, S_3, SYOUSUU,
       '[' + ISNULL(SYOKITI,'') + ']' AS SYOKITI
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('067012','067013','067018','067019','600040')
ORDER BY 1
"@ 20

Q '--- 8. 参考: コースが使う基準セット (T_COURSE1 に KIJUN の列があるか) ---' @"
SELECT c.TABLE_NAME AS テーブル, c.COLUMN_NAME AS 列
FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE c.COLUMN_NAME LIKE '%KIJUN%'
ORDER BY 1, 2
"@ 60

W ''
W '=== 読み方 ==='
W '  3 が0行なら、視力の判定は基準値マスタではなく健診ナビの内部の決まりで出ています。'
W '     その場合、こちらでは直せないので江東微研に問い合わせることになります。'
W '  3 に行があれば、そこが設定場所です。0.7〜0.9 の行の HANTEI_KIGO を見てください。'
W '  4 で、どの基準セット (1=学会準拠 / 10=東振協 など) があるかが分かります。'
W '  5 と 6 で、ALT や LDL が「学会のC → 健診ナビのD」で設定されていることを確かめます。'
W '     これが確認できれば、視力の 0.7〜0.9 も D であるべき、という根拠になります。'
notepad $out
