<#
  コースを新しく作るための下調べ その3 (check63.ps1)
  check63.bat をダブルクリックすると実行され、結果 r_check63.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check62 で分かったこと
    ・T_DANTAI1 は DANTAI_CD1, MEISYO1, KENPO_CD, MEISYO1E, DANTAIGROUP_CD の5列
    ・団体 9999999997 が雛形カタログらしい (44コース。院内ドック一式がここにある)
    ・ドックの雛形候補: SHSD 日帰りドック(院内) 69項目 / KIJUN_CD=1 / 帳票101
    ・城西学園は 0000000370 の YSA1 定期健診城西 48項目 (団体ごとに専用コースを作る運用)
    ・T_COURSE2 の SET_CD はほぼ空 (43479件が空) なので、項目を1行ずつ並べる形

  ここで見たいこと
    マイクロンが団体マスタにいるか / 雛形コースの中身 (項目CDと名前) / 新しい団体CDの採り方
#>
[CmdletBinding()]
param(
    [string]$Name   = 'マイクロン',
    [string]$Dock   = 'SHSD',        # ドックの雛形コース
    [string]$DockD  = '9999999997',  # その団体
    [string]$Teiki  = 'YSA1',        # 東振協系の実績コース (城西)
    [string]$TeikiD = '0000000370'
)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check63.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== コースを新しく作るための下調べ3 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★「$Name」が団体マスタにいるか ---" @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 団体名,
       KENPO_CD AS 健保CD, LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,''))) AS 団体グループ
FROM T_DANTAI1
WHERE MEISYO1 LIKE N'%$Name%' OR MEISYO1E LIKE N'%$Name%' OR MEISYO1 LIKE N'%ﾏｲｸﾛﾝ%'
"@ 30

Q '--- 2. 団体CDの採り方 (いちばん大きい番号のあたり) ---' @"
SELECT TOP 15 LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 団体名
FROM T_DANTAI1
WHERE DANTAI_CD1 NOT LIKE '9999%'
ORDER BY DANTAI_CD1 DESC
"@ 30

Q "--- 3. 雛形カタログ団体 $DockD のコース一覧 (44本) ---" @"
SELECT LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       LTRIM(RTRIM(ISNULL(c.KIJUN_CD,''))) AS 基準値CD, LTRIM(RTRIM(ISNULL(c.TYOHYO,''))) AS 帳票,
       c.KENSIN_SYUBETU AS 健診種別, LTRIM(RTRIM(ISNULL(c.F_TOSINKYO,''))) AS 東振協F,
       LTRIM(RTRIM(ISNULL(c.F_TOKUTEI,''))) AS 特定F, LTRIM(RTRIM(ISNULL(c.F_MUKOU,''))) AS 無効F,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c
WHERE c.DANTAI_CD1 = '$DockD'
ORDER BY c.COURSE_CD
"@ 60

Q '--- 4. 東振協フラグの立っているコース (F_TOSINKYO=1) ---' @"
SELECT TOP 40 LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体名,
       LTRIM(RTRIM(c.COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE LTRIM(RTRIM(ISNULL(c.F_TOSINKYO,''))) = '1'
ORDER BY c.DANTAI_CD1, c.COURSE_CD
"@ 50

Q "--- 5. ★ドックの雛形 $DockD / $Dock の検査項目 (69個) ---" @"
SELECT c2.NARABI AS 並び, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(k.TAN_I,''))) AS 単位,
       LTRIM(RTRIM(ISNULL(k.K_KOMOKU,''))) AS 項目区分,
       LTRIM(RTRIM(ISNULL(c2.SET_CD,''))) AS セットCD
FROM T_COURSE2 c2
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE c2.DANTAI_CD1 = '$DockD' AND LTRIM(RTRIM(c2.COURSE_CD)) = '$Dock'
ORDER BY c2.NARABI, c2.KOMOKU_CD
"@ 120

Q "--- 6. ★城西の $TeikiD / $Teiki の検査項目 (48個。東振協A2/Bの参考) ---" @"
SELECT c2.NARABI AS 並び, LTRIM(RTRIM(c2.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(k.MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(k.TAN_I,''))) AS 単位
FROM T_COURSE2 c2
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(c2.KOMOKU_CD))
WHERE c2.DANTAI_CD1 = '$TeikiD' AND LTRIM(RTRIM(c2.COURSE_CD)) = '$Teiki'
ORDER BY c2.NARABI, c2.KOMOKU_CD
"@ 80

Q "--- 7. $DockD / $Dock の料金と対象条件 (T_COURSE3 / T_COURSE4) ---" @"
SELECT '料金' AS 種別, LTRIM(RTRIM(COURSE_CD)) AS コースCD,
       KENPO_RYOUKIN AS 健保料金, KENPO_JOUGEN AS 健保上限,
       DANTAI_RYOUKIN AS 団体料金, DANTAI_JOUGEN AS 団体上限,
       KOJIN_RYOUKIN AS 個人料金, DAIKOU_RYOUKIN AS 代行料金, SONOTA_RYOUKIN AS その他料金
FROM T_COURSE3 WHERE DANTAI_CD1 = '$DockD' AND LTRIM(RTRIM(COURSE_CD)) IN ('$Dock','A','B')
"@ 30

Q "--- 8. $DockD の T_COURSE4 (対象条件) ---" @"
SELECT LTRIM(RTRIM(COURSE_CD)) AS コースCD, LTRIM(RTRIM(ISNULL(NENREI_SYOSAI,''))) AS 年齢詳細,
       LTRIM(RTRIM(ISNULL(GUUKISUU,''))) AS 偶奇, LTRIM(RTRIM(ISNULL(K_NENSAN,''))) AS 年齢計算,
       LTRIM(RTRIM(ISNULL(SEIBETU,''))) AS 性別, LTRIM(RTRIM(ISNULL(ZOKUGARA,''))) AS 続柄,
       LTRIM(RTRIM(ISNULL(D_YUUKOO_F,''))) AS 有効開始, LTRIM(RTRIM(ISNULL(D_YUUKOO_T,''))) AS 有効終了
FROM T_COURSE4 WHERE DANTAI_CD1 = '$DockD' AND LTRIM(RTRIM(COURSE_CD)) IN ('$Dock','A','B')
"@ 30

Q '--- 9. 指示書のドック独自項目が項目マスタにあるか (名前でさがす) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(TAN_I,''))) AS 単位, LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%総蛋白%' OR MEISYO1 LIKE N'%ｱﾙﾌﾞﾐﾝ%' OR MEISYO1 LIKE N'%アルブミン%'
   OR MEISYO1 LIKE N'%尿素窒素%' OR MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%' OR MEISYO1 LIKE N'%LDH%'
   OR MEISYO1 LIKE N'%ｺﾘﾝ%' OR MEISYO1 LIKE N'%ｱﾐﾗｰｾﾞ%' OR MEISYO1 LIKE N'%HBs%'
   OR MEISYO1 LIKE N'%CRP%' OR MEISYO1 LIKE N'%血清鉄%' OR MEISYO1 LIKE N'%ABO%'
   OR MEISYO1 LIKE N'%眼底%' OR MEISYO1 LIKE N'%眼圧%' OR MEISYO1 LIKE N'%肺機能%'
   OR MEISYO1 LIKE N'%ｽﾊﾟｲﾛ%' OR MEISYO1 LIKE N'%超音波%' OR MEISYO1 LIKE N'%沈渣%'
   OR MEISYO1 LIKE N'%比重%' OR MEISYO1 LIKE N'%便潜血%' OR MEISYO1 LIKE N'%A/G%' OR MEISYO1 LIKE N'%A／G%'
ORDER BY KOMOKU_CD
"@ 200

W ''
W '=== 読み方 ==='
W '  1 でマイクロンが登録済みか。0件なら団体から作る (2 で次の番号を見る)。'
W '  3 で雛形カタログにどんなコースがあるか。ドックに近いものを選ぶ。'
W '  5 が本命。日帰りドックの69項目と、指示書の●を突き合わせれば、'
W '  そのままマイクロン用ドックコースの中身になる。'
W '  6 は東振協系の実績コース。A2/B はこれを雛形にできる。'
W '  ※ 読むだけです。コースは作っていません。'
notepad $out
