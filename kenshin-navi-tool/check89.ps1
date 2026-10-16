<#
  eGFR (068676) の入り具合を調べる (check89.ps1)
  check89.bat をダブルクリックすると実行され、結果 r_check89.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  「eGFR が入っていない」との報告。SRL のファイルには全員分の eGFR があり、ツールは枠がある人には書いた
  (枠が無い人は「取込対象外(この人のコースに枠なし)」)。
    1) 石屋で eGFR の枠がある人・値がある人・枠が無い人を 受診日×コース で数える
    2) 項目マスタで eGFR と、それを含みそうなセット項目
    3) 石屋以外 (2026 年) で eGFR の枠がある人のうち値が入っている割合 (健診ナビが自動計算しているかの手がかり)
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check89.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== eGFR の確認 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★石屋 (9/3・9/9・9/10): 受診日×コースごとの eGFR の枠と値 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, COUNT(*) AS 人数,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068676') THEN 1 ELSE 0 END) AS eGFR枠あり,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068676' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') THEN 1 ELSE 0 END) AS eGFR値あり,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068674' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') THEN 1 ELSE 0 END) AS ｸﾚｱﾁﾆﾝ値あり
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
GROUP BY s.D_KENSIN, s.COURSE_CD ORDER BY 1, 2
"@ 30

Q '--- 2. 項目マスタ: eGFR と 腎機能らしいセット項目 ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068674','068675','068676') OR MEISYO1 LIKE N'%eGFR%' OR MEISYO1 LIKE N'%GFR%' OR MEISYO1 LIKE N'%腎%' OR MEISYO1 LIKE N'%ｸﾚｱﾁﾆﾝ%' OR MEISYO1 LIKE N'%クレアチニン%'
ORDER BY 1
"@ 30

Q '--- 3. 石屋以外 (2026 年): eGFR の枠がある人のうち、値が入っている人数 (自動計算されているかの手がかり) ---' @"
SELECT CONVERT(varchar(7), s.D_KENSIN, 111) AS 年月, COUNT(*) AS eGFR枠あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値あり,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) = '068674' AND LTRIM(RTRIM(ISNULL(c.KEKKA,''))) <> '') THEN 1 ELSE 0 END) AS ｸﾚｱﾁﾆﾝ値あり
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068676'
  AND NOT (s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831)
GROUP BY CONVERT(varchar(7), s.D_KENSIN, 111) ORDER BY 1
"@ 20

Q '--- 4. 石屋で eGFR の枠はあるのに値が空の人 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース
FROM T_KENSIN s
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068676' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) = '')
ORDER BY 1, 2
"@ 60

W ''
W '=== 読み方 ==='
W '  1: 「eGFR枠あり」= その人のコースに eGFR の項目がある人。「枠あり」より「値あり」が少なければ、4 の人に SRL を 3→4 で入れ直せる。'
W '     「枠あり」が人数より少ない分は、コースに eGFR が無い人。健診ナビの「検査追加/取消(個別)」で eGFR (068676) を足すか、'
W '     コースの契約項目に eGFR を足してから (健診ナビ側の設定作業) 取り込む。'
W '  3: 石屋以外で枠はあるのに値が無い人が多ければ、健診ナビは eGFR を自動計算していない (= 取込で入れる必要がある)。'
notepad $out
