<#
  コースを新しく作るための下調べ (check61.ps1)
  check61.bat をダブルクリックすると実行され、結果 r_check61.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  マイクロンメモリジャパンの既存コース (Bコースなど) が健診ナビにどう入っているかを見て、
  ドックコース・A2コースを同じ形で作るための材料にする。
#>
[CmdletBinding()]
param([string]$Name = 'マイクロン')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check61.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== コースを新しく作るための下調べ ($Name) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 団体 (事業所) をさがす ---' @"
SELECT TOP 20 LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD1, LTRIM(RTRIM(DANTAI_CD2)) AS 団体CD2,
       LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 団体名, LTRIM(RTRIM(ISNULL(KANA,''))) AS カナ
FROM T_DANTAI1
WHERE MEISYO1 LIKE '%$Name%' OR KANA LIKE '%$Name%'
ORDER BY DANTAI_CD1
"@ 30

Q '--- 2. その団体のコース一覧 (T_COURSE1) ---' @"
SELECT LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(c.COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO1,''))) AS コース名, LTRIM(RTRIM(ISNULL(c.MEISYO2,''))) AS 略称,
       c.KIJUN_CD AS 基準値CD
FROM T_COURSE1 c
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE d.MEISYO1 LIKE '%$Name%'
ORDER BY c.COURSE_CD
"@ 60

Q '--- 3. T_COURSE1 の列一覧 ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ, c.is_nullable AS NULL可
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_COURSE1') ORDER BY c.column_id
"@ 120

Q '--- 4. T_COURSE2 の列一覧 (コースに入る検査項目) ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ, c.is_nullable AS NULL可
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_COURSE2') ORDER BY c.column_id
"@ 120

Q '--- 5. T_COURSE3 / T_COURSE4 の列一覧 ---' @"
SELECT o.name AS テーブル, c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
JOIN sys.objects o ON o.object_id = c.object_id
WHERE o.name IN ('T_COURSE3','T_COURSE4') ORDER BY o.name, c.column_id
"@ 150

Q '--- 6. ★その団体の各コースに入っている検査項目 (T_COURSE2) ---' @"
SELECT LTRIM(RTRIM(c2.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c1.MEISYO1,''))) AS コース名,
       LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目
FROM T_COURSE2 c2
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c2.DANTAI_CD1
LEFT JOIN T_COURSE1 c1 ON c1.DANTAI_CD1 = c2.DANTAI_CD1 AND c1.COURSE_CD = c2.COURSE_CD
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE d.MEISYO1 LIKE '%$Name%'
ORDER BY c2.COURSE_CD, c2.KOMOKU_CD
"@ 800

Q '--- 7. その団体のコースの料金 (T_COURSE3) ---' @"
SELECT TOP 200 LTRIM(RTRIM(c3.COURSE_CD)) AS コースCD, c3.*
FROM T_COURSE3 c3
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c3.DANTAI_CD1
WHERE d.MEISYO1 LIKE '%$Name%'
ORDER BY c3.COURSE_CD
"@ 200

Q '--- 8. 参考: 他の団体のドックコース (名前に「ドック」が付くコース) ---' @"
SELECT TOP 30 LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO1,''))) AS コース名,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE c.MEISYO1 LIKE '%ドック%'
ORDER BY 項目数 DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  1 で団体コードを確認。2 で今あるコース (Bコースが入っているはず) を見る。'
W '  6 が本命。Bコースにどの検査項目 (項目CD) が入っているかが出るので、'
W '  そこにドック用の項目を足したものが、新しく作るドックコースの中身になる。'
W '  8 は、他の団体に既にドックコースがあればそれを雛形にできるかを見るため。'
W '  ※ 読むだけです。コースは作っていません。'
notepad $out
