<#
  肺機能の値が健診ナビにどんな形で入っているかを見る (check86.ps1)
  check86.bat をダブルクリックすると実行され、結果 r_check86.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  読影業者の 9/3 ファイルの肺機能 (予測肺活量・努力肺活量・%1秒量=0.897 のような小数) を
  健診ナビに入れるとき、%1秒量 を 89.7 (%) にするのか 0.897 のままかを決めるために、
  すでに入っている人 (院内など) の値の形を見る。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check86.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 肺機能の値の形 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★2026 年に肺機能の値が入っている人の例 (新しい順に 8 人分・氏名は出さない) ---' @"
SELECT TOP 56 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.PK_SEQ,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名,
       LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS 値, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS 判定,
       LTRIM(RTRIM(ISNULL(m.TANI,''))) AS 単位
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) IN ('067150','067151','067152','067153','067154','067155','067156')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' AND s.D_KENSIN >= '2026/01/01'
  AND s.PK_SEQ IN (SELECT TOP 8 s2.PK_SEQ FROM T_KENSIN s2 JOIN T_KENSA k2 ON k2.PK_SEQ = s2.PK_SEQ
                   WHERE LTRIM(RTRIM(k2.KOMOKU_CD)) = '067156' AND LTRIM(RTRIM(ISNULL(k2.KEKKA,''))) <> '' AND s2.D_KENSIN >= '2026/01/01'
                   ORDER BY s2.D_KENSIN DESC)
ORDER BY s.D_KENSIN DESC, s.PK_SEQ, k.KOMOKU_CD
"@ 60

Q '--- 2. 項目マスタの肺機能 7 項目 (単位・桁の設定があれば) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名, LTRIM(RTRIM(ISNULL(TANI,''))) AS 単位
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('067149','067150','067151','067152','067153','067154','067155','067156')
ORDER BY 1
"@ 10

W ''
W '=== 読み方 ==='
W '  1 で %1秒量(067156)・1秒率(067153)・%肺活量(067155) が 89.7 のような数字なら % で入れる (業者の 0.897 は 100 倍する)。'
W '    0.897 のような小数なら、そのまま入れる。'
W '  「TANI という列はありません」と出たら、単位の列名が違うだけなので 1 の 値 の列だけ見ればよい。'
notepad $out
