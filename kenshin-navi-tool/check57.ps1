<#
  数値項目の基準値マスタ T_KIJUN1 を見る (check57.ps1)
  check57.bat をダブルクリックすると実行され、結果 r_check57.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check56 で分かったこと
    T_KIJUN2 … 定性項目 ((-)(+)、所見なし/あり など) の基準値
    T_KIJUN1 … 数値項目の基準値 (範囲)
  視力・ALT・LDL は数値なので T_KIJUN1 にある。そこを見る。

  確かめたいこと
    ・視力 067012/067013/067018/067019 の 0.7〜0.9 が何判定に設定されているか
    ・ALT・LDL・総コレステロールが「学会のC → 健診ナビのD」で設定されているか
    ・城西コース(YSA1)がどの基準セット(1=基準値1 / 10=東振協)を使っているか
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check57.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 数値項目の基準値マスタ T_KIJUN1 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KIJUN1 の列 ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型, c.CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_KIJUN1' ORDER BY c.ORDINAL_POSITION
"@ 25

Q '--- 2. T_KIJUN1 の基準セットの種類 ---' @"
SELECT KIJUN_CD, MAX(KIJUN_MEISYO) AS 基準名, COUNT(*) AS 行数,
       COUNT(DISTINCT LTRIM(RTRIM(KOMOKU_CD))) AS 項目数
FROM T_KIJUN1 GROUP BY KIJUN_CD ORDER BY 1
"@ 20

Q '--- 3. ★★視力4項目の基準値 (0.7〜0.9 が何判定か) ---' @"
SELECT * FROM T_KIJUN1
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('067012','067013','067018','067019','600040')
ORDER BY KOMOKU_CD, KIJUN_CD, RENBAN
"@ 80

Q '--- 4. ★確認: ALT(GPT) 068637 ---' @"
SELECT * FROM T_KIJUN1 WHERE LTRIM(RTRIM(KOMOKU_CD)) = '068637' ORDER BY KIJUN_CD, RENBAN
"@ 40

Q '--- 5. ★確認: LDL 068709 / 総ｺﾚｽﾃﾛｰﾙ 068705 ---' @"
SELECT * FROM T_KIJUN1 WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068709','068705') ORDER BY KOMOKU_CD, KIJUN_CD, RENBAN
"@ 60

Q '--- 6. ★城西コースが使う基準セット ---' @"
SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, c.MEISYO AS コース名, c.KIJUN_CD,
       (SELECT TOP 1 k.KIJUN_MEISYO FROM T_KIJUN1 k WHERE k.KIJUN_CD = c.KIJUN_CD) AS 基準名
FROM T_COURSE1 c
WHERE LTRIM(RTRIM(c.COURSE_CD)) LIKE 'YSA%' OR c.MEISYO LIKE N'%城西%'
ORDER BY 1
"@ 30

Q '--- 7. 参考: 全コースの基準セットの使われ方 ---' @"
SELECT c.KIJUN_CD, COUNT(*) AS コース数
FROM T_COURSE1 c GROUP BY c.KIJUN_CD ORDER BY 1
"@ 20

W ''
W '=== 読み方 ==='
W '  3 が本体です。視力の行の「判定記号」と「範囲」を見てください。'
W '     0.7〜0.9 が B になっていれば、そこを D に直せば学会基準にそろいます。'
W '  4 と 5 で、ALT・LDL・総コレステロールが'
W '     学会B→健診ナビB / 学会C→健診ナビD / 学会D→健診ナビF'
W '     で設定されていることを確かめます。視力だけ違うことの根拠になります。'
W '  6 で城西コースがどちらの基準セットを使っているかが分かります。'
notepad $out
