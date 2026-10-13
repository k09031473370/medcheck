<#
  9/3・9/9・9/10 の全員に、血液・受診票の値が入ったかを確認する (check84.ps1)
  check84.bat をダブルクリックすると実行され、結果 r_check84.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  2026/10/13 に SRL(血液)・910(受診票・問診票) の取込をやり直したあと、
    1) 受診日ごとに「値が入っている人数」を主な項目別に数える
    2) 主な項目のどれかが入っていない人を 1 人ずつ出す
  ためのもの。欠席の人 (予約だけ残っている人) はここに出て当然なので、名前で見分ける。

  記号: ○ = 値が入っている / △ = 枠はあるが値が空 / × = 枠が無い
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check84.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 9/3・9/9・9/10 全員の取込確認 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 1 人ごとに、項目の状態 (○△×) を出す共通部分
$person = @"
  SELECT s.PK_SEQ, s.D_KENSIN, s.UKE_NO_KENSA,
         LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 漢字,
         LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS 枠数,
         (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS 値あり,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068682') AS 血糖,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = 'R06-0002') AS 中性脂肪,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068604') AS ｱﾙﾌﾞﾐﾝ,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068782') AS ﾋﾞﾘﾙﾋﾞﾝ,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069245') AS 便潜血,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '001211') AS 身長,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '001253') AS 血圧,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067012') AS 視力,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069207') AS 尿,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '084001') AS 問診,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A') AS 心電図,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067052') AS 眼底,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A') AS 胃部,
         (SELECT CASE WHEN COUNT(*) = 0 THEN N'×' WHEN MAX(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) = 1 THEN N'○' ELSE N'△' END FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A') AS 胸部
  FROM T_KENSIN s
  LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
  LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
  WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
"@

Q '--- 1. ★受診日ごとの人数と、値が入っている人数 (項目別) ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, COUNT(*) AS 予約人数,
       SUM(CASE WHEN p.値あり > 0 THEN 1 ELSE 0 END) AS 何か入っている人,
       SUM(CASE WHEN p.血糖 = N'○' THEN 1 ELSE 0 END) AS 血糖,
       SUM(CASE WHEN p.中性脂肪 = N'○' THEN 1 ELSE 0 END) AS 中性脂肪,
       SUM(CASE WHEN p.ｱﾙﾌﾞﾐﾝ = N'○' THEN 1 ELSE 0 END) AS ｱﾙﾌﾞﾐﾝ,
       SUM(CASE WHEN p.ﾋﾞﾘﾙﾋﾞﾝ = N'○' THEN 1 ELSE 0 END) AS ﾋﾞﾘﾙﾋﾞﾝ,
       SUM(CASE WHEN p.便潜血 = N'○' THEN 1 ELSE 0 END) AS 便潜血,
       SUM(CASE WHEN p.身長 = N'○' THEN 1 ELSE 0 END) AS 身長,
       SUM(CASE WHEN p.血圧 = N'○' THEN 1 ELSE 0 END) AS 血圧,
       SUM(CASE WHEN p.視力 = N'○' THEN 1 ELSE 0 END) AS 視力,
       SUM(CASE WHEN p.尿 = N'○' THEN 1 ELSE 0 END) AS 尿,
       SUM(CASE WHEN p.問診 = N'○' THEN 1 ELSE 0 END) AS 問診,
       SUM(CASE WHEN p.心電図 = N'○' THEN 1 ELSE 0 END) AS 心電図,
       SUM(CASE WHEN p.眼底 = N'○' THEN 1 ELSE 0 END) AS 眼底,
       SUM(CASE WHEN p.胃部 = N'○' THEN 1 ELSE 0 END) AS 胃部,
       SUM(CASE WHEN p.胸部 = N'○' THEN 1 ELSE 0 END) AS 胸部
FROM (
$person
) p
GROUP BY p.D_KENSIN ORDER BY 1
"@ 10

Q '--- 2. ★血糖・ｱﾙﾌﾞﾐﾝ・ﾋﾞﾘﾙﾋﾞﾝ・身長・血圧・問診 のどれかが○でない人 (欠席の人はここに出て当然) ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, p.UKE_NO_KENSA AS 受付番号, p.漢字, p.団体, p.コース,
       p.枠数, p.値あり, p.血糖, p.中性脂肪, p.ｱﾙﾌﾞﾐﾝ, p.ﾋﾞﾘﾙﾋﾞﾝ, p.便潜血, p.身長, p.血圧, p.視力, p.尿, p.問診, p.心電図, p.眼底, p.胃部, p.胸部
FROM (
$person
) p
WHERE p.血糖 <> N'○' OR p.ｱﾙﾌﾞﾐﾝ <> N'○' OR p.ﾋﾞﾘﾙﾋﾞﾝ <> N'○' OR p.身長 <> N'○' OR p.血圧 <> N'○' OR p.問診 <> N'○'
ORDER BY p.D_KENSIN, p.UKE_NO_KENSA
"@ 250

Q '--- 3. 参考: 中性脂肪・心電図・眼底・胃部・胸部 のどれかが × (枠なし) か △ (枠はあるが空) の人 ---' @"
SELECT CONVERT(varchar(10), p.D_KENSIN, 111) AS 受診日, p.UKE_NO_KENSA AS 受付番号, p.漢字, p.団体, p.コース,
       p.枠数, p.値あり, p.中性脂肪, p.心電図, p.眼底, p.胃部, p.胸部
FROM (
$person
) p
WHERE p.値あり > 0
  AND (p.中性脂肪 <> N'○' OR p.心電図 <> N'○' OR p.眼底 <> N'○' OR p.胃部 <> N'○' OR p.胸部 <> N'○')
ORDER BY p.D_KENSIN, p.UKE_NO_KENSA
"@ 250

W ''
W '=== 読み方 ==='
W '  記号: ○ = 値が入っている / △ = 枠はあるが値が空 / × = 枠が無い (そのコースにその検査が無い)'
W '  1: 受診日ごとに、各項目に値が入っている人数。予約人数との差は、欠席・取消し忘れ・未取込のどれか。'
W '  2: 血液か受診票が入っていない人。欠席 (秋山浩章・石井邦夫・片山英太) と国立高専の人は入っていなくて正しい。'
W '     それ以外の人がここに出たら、その人は取込ができていないので受付番号を教えてください。'
W '  3: 任意検査 (心電図・眼底・胃部・胸部) が入っていない人。胸部X線は業者からまだ来ていないので △ で正常。'
W '     眼底・胃部が × の人は、そのコースに検査が無い人なので、受けていなければそのままでよい。'
notepad $out
