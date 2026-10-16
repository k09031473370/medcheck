<#
  石屋: 枠はあるのに値が空の項目を全部出す (check94.ps1)
  check94.bat をダブルクリックすると実行され、結果 r_check94.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  受診者 199 人 (受付番号 4631〜4831) について、健診ナビに枠がある全項目を項目ごとに数え、
  「枠はあるが値が空」の人がいる項目だけを、空の人数が多い順に出す。
  セットの親 (【…】) と 判定の枠 (600…/610…) は別の表に分ける。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check94.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 枠はあるが空の項目 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$base = @"
  SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS NM, s.D_KENSIN,
         CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' OR LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) <> '' OR LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END AS F
  FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
  LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
  WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
"@

Q '--- 1. ★検査項目: 枠はあるが値が空の人がいる項目 (空が多い順) ---' @"
SELECT b.CD AS 項目CD, b.NM AS 項目名, COUNT(*) AS 枠あり, SUM(b.F) AS 値あり, COUNT(*) - SUM(b.F) AS 空,
       SUM(CASE WHEN b.D_KENSIN = '2026/09/03' AND b.F = 0 THEN 1 ELSE 0 END) AS 空_0903,
       SUM(CASE WHEN b.D_KENSIN = '2026/09/09' AND b.F = 0 THEN 1 ELSE 0 END) AS 空_0909,
       SUM(CASE WHEN b.D_KENSIN = '2026/09/10' AND b.F = 0 THEN 1 ELSE 0 END) AS 空_0910
FROM ($base) b
WHERE b.NM NOT LIKE N'【%' AND b.CD NOT LIKE '600%' AND b.CD NOT LIKE '610%' AND b.CD NOT LIKE '11%'
GROUP BY b.CD, b.NM
HAVING COUNT(*) - SUM(b.F) > 0
ORDER BY 5 DESC, 1
"@ 120

Q '--- 2. 参考: セットの親・判定の枠で空があるもの (自動判定で埋まるものが多い) ---' @"
SELECT b.CD AS 項目CD, b.NM AS 項目名, COUNT(*) AS 枠あり, SUM(b.F) AS 値あり, COUNT(*) - SUM(b.F) AS 空
FROM ($base) b
WHERE b.NM LIKE N'【%' OR b.CD LIKE '600%' OR b.CD LIKE '610%' OR b.CD LIKE '11%'
GROUP BY b.CD, b.NM
HAVING COUNT(*) - SUM(b.F) > 0
ORDER BY 5 DESC, 1
"@ 80

Q '--- 3. 参考: 全部埋まっている検査項目の数と、一部でも空がある検査項目の数 ---' @"
SELECT SUM(CASE WHEN 空 = 0 THEN 1 ELSE 0 END) AS 全部埋まっている項目, SUM(CASE WHEN 空 > 0 THEN 1 ELSE 0 END) AS 空がある項目
FROM (SELECT b.CD, COUNT(*) - SUM(b.F) AS 空 FROM ($base) b
      WHERE b.NM NOT LIKE N'【%' AND b.CD NOT LIKE '600%' AND b.CD NOT LIKE '610%' AND b.CD NOT LIKE '11%'
      GROUP BY b.CD) x
"@ 5

W ''
W '=== 読み方 ==='
W '  1 に出た項目が「枠はあるのに入っていない」もの。受けていない検査 (便を出していない人の便潜血、胃部を受けていない人の胃部 など) は'
W '     空で正しい。それ以外 (全員受けているはずの検査) が出ていれば、取込漏れか、結果がまだ来ていないもの。'
W '  2 の判定の枠は健診ナビの自動判定で埋まる。医師指示事項 (11xxx) は空で正常。'
notepad $out
