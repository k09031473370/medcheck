<#
  コース複製の最終確認 (check64.ps1)
  check64.bat をダブルクリックすると実行され、結果 r_check64.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check63 で分かったこと
    ・東振協の3コースは雛形カタログ 9999999997 に既にある
        SHTD1 東振協D1(院内)  70項目  ← 指示書の「ドック健診」(備考に D1コース と明記)
        SHTB  東振協B(院内)   52項目  ← 生活習慣病健診Bコース
        SHTA2 東振協A2(院内)  37項目  ← 簡易生活習慣病健診A2コース
      いずれも 帳票201 / 健診種別60 / F_TOSINKYO=1 / KIJUN_CD=1
    ・矢吹海運(0000000048)・東京国際埠頭(0000000049) が3本セットで持っている
      → マイクロン(0000000475) も同じ形にすればよい

  ここで見たいこと
    実際に3本持っている団体が T_COURSE1/3/4 をどう入れているか (これが複製の見本)
    COURSE_CD を持つテーブルが他にないか (複製し忘れがないか)
#>
[CmdletBinding()]
param(
    [string]$Src    = '9999999997',   # 雛形カタログ
    [string]$Sample = '0000000048',   # 3本セットで持っている団体 (矢吹海運)
    [string]$New    = '0000000475'    # マイクロンメモリジャパン
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check64.txt'
$CS   = "'SHTA2','SHTB','SHTD1'"
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== コース複製の最終確認 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. COURSE_CD を持つテーブル (複製し忘れがないか) ---' @"
SELECT o.name AS テーブル, MAX(CASE WHEN c.name = 'DANTAI_CD1' THEN 1 ELSE 0 END) AS 団体CDあり,
       COUNT(*) AS 列数
FROM sys.columns c JOIN sys.objects o ON o.object_id = c.object_id
WHERE o.type = 'U' AND o.object_id IN (SELECT object_id FROM sys.columns WHERE name = 'COURSE_CD')
GROUP BY o.name ORDER BY o.name
"@ 60

Q "--- 2. ★見本: $Sample (矢吹海運) の T_COURSE1 全列 ---" @"
SELECT LTRIM(RTRIM(COURSE_CD)) AS コースCD, LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD,
       LTRIM(RTRIM(ISNULL(KIJUN_CD,''))) AS 基準値CD, LTRIM(RTRIM(ISNULL(MEISYO,''))) AS コース名,
       LTRIM(RTRIM(ISNULL(MEISYO_RYAKUSYO,''))) AS 略称, LTRIM(RTRIM(ISNULL(TYOHYO,''))) AS 帳票,
       LTRIM(RTRIM(ISNULL(IRAI_ID,''))) AS 依頼ID, KENSIN_SYUBETU AS 健診種別,
       LTRIM(RTRIM(ISNULL(JIKO_FUTAN_BUN,''))) AS 自己負担文,
       LTRIM(RTRIM(ISNULL(F_SOTOZEI,''))) AS 外税, LTRIM(RTRIM(ISNULL(F_SYOONIN,''))) AS 承認,
       LTRIM(RTRIM(ISNULL(F_FAX,''))) AS FAX, LTRIM(RTRIM(ISNULL(F_RIYOUKEN,''))) AS 利用券,
       LTRIM(RTRIM(ISNULL(F_MUKOU,''))) AS 無効, LTRIM(RTRIM(ISNULL(KOKYAKU_CD,''))) AS 顧客CD,
       LTRIM(RTRIM(ISNULL(F_KYOKAI,''))) AS 協会, LTRIM(RTRIM(ISNULL(F_TENKIYOUSI,''))) AS 転記用紙,
       LTRIM(RTRIM(ISNULL(F_TOKUTEI,''))) AS 特定, LTRIM(RTRIM(ISNULL(F_TOSINKYO,''))) AS 東振協,
       F_TOYAKU_SYUBETU AS 投薬種別, LTRIM(RTRIM(ISNULL(LABEL_MAISUU,''))) AS ラベル枚数
FROM T_COURSE1 WHERE DANTAI_CD1 IN ('$Src','$Sample') AND LTRIM(RTRIM(COURSE_CD)) IN ($CS)
ORDER BY DANTAI_CD1, COURSE_CD
"@ 30

Q "--- 3. ★料金 T_COURSE3 ($Src と $Sample。どの列に入れるか) ---" @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(COURSE_CD)) AS コースCD,
       KENPO_RYOUKIN AS 健保料金, KENPO_JOUGEN AS 健保上限,
       DANTAI_RYOUKIN AS 団体料金, DANTAI_JOUGEN AS 団体上限,
       KOJIN_RYOUKIN AS 個人料金, KOJIN_JOUGEN AS 個人上限,
       DAIKOU_RYOUKIN AS 代行料金, DAIKOU_JOUGEN AS 代行上限,
       SONOTA_RYOUKIN AS その他料金, SONOTA_JOUGEN AS その他上限
FROM T_COURSE3 WHERE DANTAI_CD1 IN ('$Src','$Sample') AND LTRIM(RTRIM(COURSE_CD)) IN ($CS)
ORDER BY DANTAI_CD1, COURSE_CD
"@ 30

Q "--- 4. ★対象条件 T_COURSE4 ($Src と $Sample) ---" @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(ISNULL(NENREI_SYOSAI,''))) AS 年齢詳細, LTRIM(RTRIM(ISNULL(GUUKISUU,''))) AS 偶奇,
       LTRIM(RTRIM(ISNULL(K_NENSAN,''))) AS 年齢計算, LTRIM(RTRIM(ISNULL(SEIBETU,''))) AS 性別,
       LTRIM(RTRIM(ISNULL(ZOKUGARA,''))) AS 続柄, LTRIM(RTRIM(ISNULL(BUSYO_CD,''))) AS 部署,
       LTRIM(RTRIM(ISNULL(KOYOKEITAI,''))) AS 雇用形態,
       LTRIM(RTRIM(ISNULL(D_YUUKOO_F,''))) AS 有効開始, LTRIM(RTRIM(ISNULL(D_YUUKOO_T,''))) AS 有効終了,
       LTRIM(RTRIM(ISNULL(SEIGENNITI,''))) AS 制限日
FROM T_COURSE4 WHERE DANTAI_CD1 IN ('$Src','$Sample') AND LTRIM(RTRIM(COURSE_CD)) IN ($CS)
ORDER BY DANTAI_CD1, COURSE_CD
"@ 30

Q "--- 5. 雛形と見本で項目 (T_COURSE2) が同じか ---" @"
SELECT LTRIM(RTRIM(COURSE_CD)) AS コースCD,
       SUM(CASE WHEN DANTAI_CD1 = '$Src' THEN 1 ELSE 0 END) AS 雛形の項目数,
       SUM(CASE WHEN DANTAI_CD1 = '$Sample' THEN 1 ELSE 0 END) AS 見本の項目数,
       COUNT(DISTINCT KOMOKU_CD) AS 項目CDの種類
FROM T_COURSE2 WHERE DANTAI_CD1 IN ('$Src','$Sample') AND LTRIM(RTRIM(COURSE_CD)) IN ($CS)
GROUP BY COURSE_CD ORDER BY COURSE_CD
"@ 30

Q "--- 6. ★東振協D1 ($Src / SHTD1) の70項目 (指示書の●と突き合わせる) ---" @"
SELECT c2.NARABI AS 並び, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目, LTRIM(RTRIM(ISNULL(k.TAN_I,''))) AS 単位
FROM T_COURSE2 c2
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE c2.DANTAI_CD1 = '$Src' AND LTRIM(RTRIM(c2.COURSE_CD)) = 'SHTD1'
ORDER BY c2.NARABI, c2.KOMOKU_CD
"@ 120

Q "--- 7. 東振協B ($Src / SHTB) の52項目 ---" @"
SELECT c2.NARABI AS 並び, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目, LTRIM(RTRIM(ISNULL(k.TAN_I,''))) AS 単位
FROM T_COURSE2 c2
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE c2.DANTAI_CD1 = '$Src' AND LTRIM(RTRIM(c2.COURSE_CD)) = 'SHTB'
ORDER BY c2.NARABI, c2.KOMOKU_CD
"@ 80

Q "--- 8. 東振協A2 ($Src / SHTA2) の37項目 ---" @"
SELECT c2.NARABI AS 並び, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目, LTRIM(RTRIM(ISNULL(k.TAN_I,''))) AS 単位
FROM T_COURSE2 c2
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE c2.DANTAI_CD1 = '$Src' AND LTRIM(RTRIM(c2.COURSE_CD)) = 'SHTA2'
ORDER BY c2.NARABI, c2.KOMOKU_CD
"@ 60

Q "--- 9. $New (マイクロン) に今あるコース (空のはず) ---" @"
SELECT LTRIM(RTRIM(COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(MEISYO,''))) AS コース名
FROM T_COURSE1 WHERE DANTAI_CD1 = '$New'
"@ 30

Q '--- 10. オプションのコース (マンモ・子宮・情報機器・鉛) がマスタにあるか ---' @"
SELECT TOP 40 LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(c.COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c
WHERE c.DANTAI_CD1 = '$Src'
  AND (c.MEISYO LIKE N'%ﾏﾝﾓ%' OR c.MEISYO LIKE N'%マンモ%' OR c.MEISYO LIKE N'%乳%'
       OR c.MEISYO LIKE N'%子宮%' OR c.MEISYO LIKE N'%情報機器%' OR c.MEISYO LIKE N'%VDT%'
       OR c.MEISYO LIKE N'%鉛%' OR c.MEISYO LIKE N'%特殊%')
ORDER BY c.COURSE_CD
"@ 50

W ''
W '=== 読み方 ==='
W '  1 で T_COURSE1〜4 以外に複製すべきテーブルがないかを確認。'
W '  2〜4 が複製の見本。雛形と見本で違う列があれば、そこが団体ごとに変える列。'
W '  6〜8 の項目を指示書の●/△と突き合わせて、過不足がないかを見る。'
W '  ※ 読むだけです。コースは作っていません。'
notepad $out
