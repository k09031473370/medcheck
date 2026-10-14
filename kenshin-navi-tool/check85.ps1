<#
  胸部X線 (読影業者ファイル) を取り込む前の確認 (check85.ps1)
  check85.bat をダブルクリックすると実行され、結果 r_check85.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  読影業者のファイル (9/3 と 9/9・9/10) には
    部位コード  CG / AH / CO / AF / BI / CI
    所見コード  A16V / A16 / M01 / A02 / E04 / E07 / A17
    判定        心配なし / 要経過観察 / 要精検 / 異常なし
    じん肺      粒状影A・B / タイプ / 大陰影 / 不整形陰影 / 付加事項 (pl)
    肺機能      肺活量予測値 / 努力肺活量 / 1秒量 (9/3 のみ)
  が入っている。健診ナビ側に
    1) その部位・所見コードがマスタ (ZK020 / ZK021) にあるか、判定記号は何か
    2) 胸部X線・じん肺・肺機能の項目が項目マスタにあるか
    3) 石屋の受診者にその枠があるか
  を見て、対応表 (mapping_xray_ishiya93 / 910) で取り込めるものを決める。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check85.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 胸部X線 取込前の確認 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★部位マスタ ZK020: ファイルに出てくる部位コード ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS 所見区分, LTRIM(RTRIM(KEKKA_CD)) AS コード, LTRIM(RTRIM(ISNULL(SYOKEN,''))) AS 名称,
       LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定記号
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK020' AND LTRIM(RTRIM(KEKKA_CD)) IN ('CG','AH','CO','AF','BI','CI')
ORDER BY 2
"@ 20

Q '--- 2. ★所見マスタ ZK021: ファイルに出てくる所見コードと、健診ナビが付ける判定記号 ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS 所見区分, LTRIM(RTRIM(KEKKA_CD)) AS コード, LTRIM(RTRIM(ISNULL(SYOKEN,''))) AS 名称,
       LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定記号
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK021' AND LTRIM(RTRIM(KEKKA_CD)) IN ('A16V','A16','M01','A02','E04','E07','A17')
ORDER BY 2
"@ 20

Q '--- 3. 項目マスタ: 胸部X線 (0770…)・じん肺・肺機能らしい項目 ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名, LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見区分
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) LIKE '0770%'
   OR MEISYO1 LIKE N'%胸部%' OR MEISYO1 LIKE N'%じん肺%' OR MEISYO1 LIKE N'%塵肺%' OR MEISYO1 LIKE N'%粒状%'
   OR MEISYO1 LIKE N'%大陰影%' OR MEISYO1 LIKE N'%不整形%' OR MEISYO1 LIKE N'%肺活量%' OR MEISYO1 LIKE N'%肺機能%'
   OR MEISYO1 LIKE N'%1秒%' OR MEISYO1 LIKE N'%１秒%' OR MEISYO1 LIKE N'%ｽﾊﾟｲﾛ%' OR MEISYO1 LIKE N'%呼吸%'
ORDER BY 1
"@ 80

Q '--- 4. ★石屋 (9/3・9/9・9/10) の人が持っている、その項目の枠の数と、もう値が入っている数 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名,
       COUNT(DISTINCT s.PK_SEQ) AS 枠を持つ人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値が入っている数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '0770%'
       OR m.MEISYO1 LIKE N'%胸部%' OR m.MEISYO1 LIKE N'%じん肺%' OR m.MEISYO1 LIKE N'%塵肺%' OR m.MEISYO1 LIKE N'%粒状%'
       OR m.MEISYO1 LIKE N'%大陰影%' OR m.MEISYO1 LIKE N'%不整形%' OR m.MEISYO1 LIKE N'%肺活量%' OR m.MEISYO1 LIKE N'%肺機能%'
       OR m.MEISYO1 LIKE N'%1秒%' OR m.MEISYO1 LIKE N'%１秒%' OR m.MEISYO1 LIKE N'%呼吸%')
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY 1
"@ 80

Q '--- 5. 参考: 石屋で 胸部X線所見1 (077010A) の枠が無い人 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 漢字,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND NOT EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A')
ORDER BY 1, 2
"@ 50

W ''
W '=== 読み方 ==='
W '  1・2: ファイルの部位・所見コードが全部出ていれば、そのまま取り込める。出ないコードがあれば教えてください。'
W '      2 の「判定記号」が、健診ナビがその所見に付ける判定。業者の判定 (心配なし=A相当, 要経過観察=C, 要精検=D) と違う所見があれば'
W '      先生に相談 (健診ナビの判定基準はツールから変えない)。'
W '  3・4: 胸部X線所見1 (077010A) に石屋全員の枠があれば取込できる。じん肺・肺機能の項目が 3 に無い、または 4 で枠を持つ人が 0 なら、'
W '      その分は健診ナビに入れる場所が無い (結果票にも載らない) ので、業者のファイルを紙で保管する等の運用になる。'
W '  5: ここに出た人は胸部X線の枠が無いので、健診ナビの「検査追加/取消(個別)」で胸部X線を足してから取り込む。'
notepad $out
