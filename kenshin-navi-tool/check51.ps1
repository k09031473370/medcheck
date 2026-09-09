<#
  血糖・中性脂肪が帳票に出ない理由を調べる (check51.ps1)
  check51.bat をダブルクリックすると実行され、結果 r_check51.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  結果Excel照合で、88人全員の帳票に 血糖(空腹時/随時) と 中性脂肪(空腹時/随時) が出ていなかった。
  取込で値を入れたのは 血糖=068682 / 中性脂肪=R06-0002。
  帳票303は **[空腹時血糖] **[随時血糖] **[中性脂肪] **[随時中性脂肪] という略称の項目を見ている。
  どの項目コードがその略称を持っているか、その枠が8/21の人にあるかを見る。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check51.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 血糖・中性脂肪が帳票に出ない理由 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KOMOKU の列 (略称の列がどれかを見る) ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型
FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_KOMOKU' ORDER BY c.ORDINAL_POSITION
"@ 60

Q '--- 2. ★取込で値を入れた項目と、LAPで確認済みの項目 (全列) ---' @"
SELECT * FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068682','R06-0002','068651','068689','068705')
"@ 10

Q '--- 3. ★名前に 血糖 / 中性脂肪 / TG を含む項目 (全列) ---' @"
SELECT * FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%血糖%' OR MEISYO1 LIKE N'%中性脂肪%' OR MEISYO1 LIKE N'%TG%' OR MEISYO1 LIKE N'%ﾄﾘｸﾞﾘ%'
ORDER BY KOMOKU_CD
"@ 40

Q "--- 4. ★$Ymd の人に、それらの枠があるか・値が入っているか ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 枠の数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#') THEN 1 ELSE 0 END) AS 値あり,
       MIN(k.KEKKA) AS 例
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND (m.MEISYO1 LIKE N'%血糖%' OR m.MEISYO1 LIKE N'%中性脂肪%' OR m.MEISYO1 LIKE N'%TG%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) IN ('068682','R06-0002'))
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1
"@ 30

Q '--- 5. 帳票の設定側: 略称 空腹時血糖 / 随時血糖 / 中性脂肪 / 随時中性脂肪 / LAP を持つ行 (どの表にあっても) ---' @"
SELECT t.TABLE_NAME AS 表, c.COLUMN_NAME AS 列
FROM INFORMATION_SCHEMA.COLUMNS c JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME
WHERE t.TABLE_TYPE = 'BASE TABLE' AND c.DATA_TYPE IN ('char','varchar','nchar','nvarchar')
  AND (c.COLUMN_NAME LIKE '%RYAKU%' OR c.COLUMN_NAME LIKE '%RYK%')
ORDER BY 1, 2
"@ 40

Q "--- 6. 自動判定が何を見たか: $Ymd の【判定】糖代謝(600340)・脂質代謝(600320) の分布 ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定, COUNT(*) AS 人数
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0 AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('600340','600320')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), k.HANTEI_KIGO ORDER BY 1, 2
"@ 20

W ''
W '=== 読み方 ==='
W '  2 と 3 を見くらべて、略称 (RYAKU… の列) が 空腹時血糖/随時血糖/中性脂肪/随時中性脂肪 になっている'
W '  項目コードを探します。それが 068682 / R06-0002 と違えば、帳票はそちらを見ていて、'
W '  値の入っている項目は見ていない、ということです。'
W '  4 で、帳票が見ている項目の枠が 8/21 の人にあるか (枠の数) と、値が空か (値あり=0) が分かります。'
W '  直し方は2つ:'
W '    a) 帳票303の略称を、値の入っている項目の略称に書き換える (LAPと同じやり方)'
W '    b) 値を、帳票が見ている項目にも入れる (枠がある場合のみ)'
notepad $out
