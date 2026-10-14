<#
  石屋だけ: 胸部X線と胃部X線の入り具合 (check87.ps1)
  check87.bat をダブルクリックすると実行され、結果 r_check87.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  check84 は列が多くて右端 (胸部・胃部) が貼り付けで切れるので、石屋 (受付番号 4631〜4831) だけ・列を絞って出す。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check87.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 胸部X線・胃部X線の確認 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$person = @"
  SELECT s.PK_SEQ, s.D_KENSIN, s.UKE_NO_KENSA, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A') AS 胸部所見,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A') AS 胸部判定,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077010Z') AS 胸部ﾌｨﾙﾑ,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A') AS 胃部所見,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A') AS 心電図,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067052') AS 眼底
  FROM T_KENSIN s
  WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
    AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
"@

Q '--- 1. ★受診日ごと (石屋だけ): 値が入っている人数 ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN p.胸部所見 = N'○' THEN 1 ELSE 0 END) AS 胸部所見,
       SUM(CASE WHEN p.胸部判定 = N'○' THEN 1 ELSE 0 END) AS 胸部判定,
       SUM(CASE WHEN p.胸部ﾌｨﾙﾑ = N'○' THEN 1 ELSE 0 END) AS 胸部ﾌｨﾙﾑ,
       SUM(CASE WHEN p.胃部所見 = N'○' THEN 1 ELSE 0 END) AS 胃部所見あり,
       SUM(CASE WHEN p.胃部所見 = N'△' THEN 1 ELSE 0 END) AS 胃部枠あり空,
       SUM(CASE WHEN p.胃部所見 = N'×' THEN 1 ELSE 0 END) AS 胃部枠なし,
       SUM(CASE WHEN p.心電図 = N'○' THEN 1 ELSE 0 END) AS 心電図,
       SUM(CASE WHEN p.眼底 = N'○' THEN 1 ELSE 0 END) AS 眼底
FROM (
$person
) p
GROUP BY p.D_KENSIN ORDER BY 1
"@ 10

Q '--- 2. 胸部所見・胸部判定が ○ でない石屋の人 (0 行なら全員入っている) ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, p.UKE_NO_KENSA AS 受付番号, p.コース, p.胸部所見, p.胸部判定, p.胸部ﾌｨﾙﾑ
FROM (
$person
) p
WHERE p.胸部所見 <> N'○' OR p.胸部判定 <> N'○'
ORDER BY 1, 2
"@ 50

Q '--- 3. 胃部所見が「枠はあるが空」の石屋の人 (受診日・コースごとの人数) ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, p.コース, COUNT(*) AS 人数,
       MIN(p.UKE_NO_KENSA) AS 受付番号min, MAX(p.UKE_NO_KENSA) AS 受付番号max
FROM (
$person
) p
WHERE p.胃部所見 = N'△'
GROUP BY p.D_KENSIN, p.コース ORDER BY 1, 2
"@ 30

W ''
W '=== 読み方 ==='
W '  1: 胸部所見・胸部判定が 58 / 90 / 51 なら胸部X線は全員入っている。胸部ﾌｨﾙﾑ は 9/3 のファイルに無いので 9/3 は 0 で正常。'
W '  3: 胃部X線の枠があるのに空の人。その日の胃部X線の結果がどこにあるか (まだ来ていない・別ファイル) を確認する。'
notepad $out
