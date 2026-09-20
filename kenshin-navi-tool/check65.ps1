<#
  オプション申し込みの入れ方をさがす (check65.ps1)
  check65.bat をダブルクリックすると実行され、結果 r_check65.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  指示書のオプション
    乳がん検診マンモグラフィ        27名 4000円
    子宮細胞診(自己採取)※ドック      4名 1350円
    子宮細胞診(自己採取)※A2・B      26名  840円
    情報機器作業健診               71名 2000円
    鉛                          6名 6000円

  check64 で分かったこと
    ・M_OPTION / M_OPTION_NEW (21列, COURSE_CD と DANTAI_CD1 あり) がある
    ・雛形カタログ 9999999997 には SHSIKYU 子宮がん単独(院内) 6項目 しかない
    ・T_KOMOKU に OPTION_GRP という列がある
  → オプションはコースではなく M_OPTION で管理されている可能性が高い。その形を見る。
#>
[CmdletBinding()]
param([string]$Src = '9999999997')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check65.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== オプション申し込みの入れ方 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. M_OPTION の列一覧 ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ, c.is_nullable AS NULL可
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('M_OPTION') ORDER BY c.column_id
"@ 60

Q '--- 2. M_OPTION_NEW の列一覧 (M_OPTION との違い) ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('M_OPTION_NEW') ORDER BY c.column_id
"@ 60

Q '--- 3. M_OPTION / M_OPTION_NEW の件数と、どの団体が持っているか ---' @"
SELECT 'M_OPTION' AS テーブル, COUNT(*) AS 件数, COUNT(DISTINCT DANTAI_CD1) AS 団体数 FROM M_OPTION
UNION ALL
SELECT 'M_OPTION_NEW', COUNT(*), COUNT(DISTINCT DANTAI_CD1) FROM M_OPTION_NEW
"@ 20

Q '--- 4. ★M_OPTION の中身 (オプションを多く持っている団体 上位) ---' @"
SELECT TOP 20 LTRIM(RTRIM(o.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       COUNT(*) AS オプション数
FROM M_OPTION o LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = o.DANTAI_CD1
GROUP BY o.DANTAI_CD1, d.MEISYO1 ORDER BY COUNT(*) DESC
"@ 30

Q '--- 5. ★M_OPTION の実物 (先頭30行。全列) ---' @"
SELECT TOP 30 * FROM M_OPTION ORDER BY DANTAI_CD1, COURSE_CD
"@ 40

Q "--- 6. 雛形カタログ $Src の M_OPTION ---" @"
SELECT * FROM M_OPTION WHERE DANTAI_CD1 = '$Src'
"@ 60

Q '--- 7. ★マンモ・子宮・情報機器・鉛 のコースを全団体からさがす ---' @"
SELECT TOP 40 LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       c.KENSIN_SYUBETU AS 健診種別,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE c.MEISYO LIKE N'%ﾏﾝﾓ%' OR c.MEISYO LIKE N'%マンモ%' OR c.MEISYO LIKE N'%乳%'
   OR c.MEISYO LIKE N'%子宮%' OR c.MEISYO LIKE N'%情報機器%' OR c.MEISYO LIKE N'%VDT%'
   OR c.MEISYO LIKE N'%鉛%' OR c.MEISYO LIKE N'%特殊%' OR c.MEISYO LIKE N'%有機%'
ORDER BY c.MEISYO, c.DANTAI_CD1
"@ 50

Q '--- 8. ★マンモ・子宮・情報機器・鉛 の検査項目を項目マスタからさがす ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(TAN_I,''))) AS 単位, LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分,
       OPTION_GRP AS オプション群
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%ﾏﾝﾓ%' OR MEISYO1 LIKE N'%マンモ%' OR MEISYO1 LIKE N'%乳房%'
   OR MEISYO1 LIKE N'%子宮%' OR MEISYO1 LIKE N'%細胞診%' OR MEISYO1 LIKE N'%情報機器%'
   OR MEISYO1 LIKE N'%VDT%' OR MEISYO1 LIKE N'%鉛%' OR MEISYO1 LIKE N'%ﾃﾞﾙﾀ%'
   OR MEISYO1 LIKE N'%ｱﾐﾉﾚﾌﾞﾘﾝ%' OR MEISYO1 LIKE N'%近見%' OR MEISYO1 LIKE N'%調節%'
ORDER BY KOMOKU_CD
"@ 200

Q '--- 9. OPTION_GRP が入っている項目 (オプションの束ね方) ---' @"
SELECT OPTION_GRP AS オプション群, COUNT(*) AS 項目数,
       MIN(LTRIM(RTRIM(ISNULL(MEISYO1,'')))) AS 例1, MAX(LTRIM(RTRIM(ISNULL(MEISYO1,'')))) AS 例2
FROM T_KOMOKU WHERE ISNULL(OPTION_GRP,0) <> 0
GROUP BY OPTION_GRP ORDER BY OPTION_GRP
"@ 60

Q "--- 10. 子宮がん単独 $Src / SHSIKYU の6項目 ---" @"
SELECT LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目
FROM T_COURSE2 c2 LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE c2.DANTAI_CD1 = '$Src' AND LTRIM(RTRIM(c2.COURSE_CD)) = 'SHSIKYU'
ORDER BY c2.KOMOKU_CD
"@ 30

Q '--- 11. 実際にオプションを受診した例 (T_RYOUKIN に入っている料金明細) ---' @"
SELECT TOP 30 LTRIM(RTRIM(r.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目,
       COUNT(*) AS 件数, MIN(r.GOUKEI) AS 最小金額, MAX(r.GOUKEI) AS 最大金額
FROM T_RYOUKIN r LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
GROUP BY r.KOMOKU_CD, k.MEISYO1
ORDER BY COUNT(*) DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  1〜6 で M_OPTION がオプション申し込みの置き場かどうかを見る。'
W '  7 でマンモ・情報機器・鉛がコースとして存在するかを全団体からさがす。'
W '  8・9 で、なければ項目マスタから項目CDを拾う。'
W '  11 は、実際に受付でどの項目にいくら入力しているかの実績。料金の入れ方の参考。'
W '  ※ 読むだけです。何も作っていません。'
notepad $out
