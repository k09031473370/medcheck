<#
  福生 10/2 に入った86人の氏名を出す (check72.ps1)
  check72.bat をダブルクリックすると実行され、結果 r_check72.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check71 で分かったこと
    ・福生は 10/2 の 86人だけ入っていた (振り分け表は103人)。10/6・10/7 は0人
    ・コースは FA 福生A / FB 福生B / FC 福生C が正しく付いている
    ・二重登録・取消は0件
    ・マイクロン(0000000477)は YTDA2/YTDB/YTDD1 が既にあり、295人入っている
  check71 のエラー
    ・T_KOJIN1 に D_BIRTH という列が無かった → 列名を調べ直す
    ・集計の中に副問い合わせを入れてしまった → 書き直す

  ここでやること
    入った86人の氏名を全部出して、振り分け表の103人と突き合わせ、
    誰が入っていないかを特定する。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/10/02')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check72.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 福生 $Ymd に入った人の一覧 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KOJIN1 の列一覧 (生年月日の列名を確かめる) ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KOJIN1') ORDER BY c.column_id
"@ 80

Q "--- 2. ★$Ymd 福生に入った人 全員 (これを振り分け表と突き合わせる) ---" @"
SELECT LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
       s.KOJIN_ID, s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '$Ymd' AND d.MEISYO1 LIKE N'%福生%'
ORDER BY d.MEISYO1, g.KANJI_SIMEI
"@ 150

Q "--- 3. カナが入っている人の数 ---" @"
SELECT COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) = '' THEN 1 ELSE 0 END) AS カナなし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.SEIBETU,''))) = '' THEN 1 ELSE 0 END) AS 性別なし
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '$Ymd' AND d.MEISYO1 LIKE N'%福生%'
"@ 20

Q '--- 4. 福生のコース一覧 (FA/FB/FC。団体ごとの項目数のばらつきを見る) ---' @"
SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE d.MEISYO1 LIKE N'%福生%' AND LTRIM(RTRIM(c.COURSE_CD)) IN ('FA','FB','FC')
ORDER BY c.COURSE_CD, 項目数, c.DANTAI_CD1
"@ 200

Q '--- 5. 福生の各コースが何種類の構成になっているか (項目数で見る) ---' @"
SELECT コースCD, 項目数, COUNT(*) AS 団体数 FROM (
  SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, c.DANTAI_CD1,
         (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
  FROM T_COURSE1 c LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
  WHERE d.MEISYO1 LIKE N'%福生%' AND LTRIM(RTRIM(c.COURSE_CD)) IN ('FA','FB','FC')
) t GROUP BY コースCD, 項目数 ORDER BY コースCD, 項目数
"@ 40

W ''
W '=== 読み方 ==='
W '  2 の一覧をそのまま貼ってもらえれば、振り分け表の103人と突き合わせて'
W '  入っていない17人の氏名を出します。'
W '  4・5 は、同じ FA 福生A でも団体によって項目数が違っていたので、'
W '  どれが正しい構成かを確かめるため。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
