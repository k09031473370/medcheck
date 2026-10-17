<#
  石屋: HbA1c・血小板・eGFR の判定を入れたことで 総合判定 が変わるはずの人を拾う (check97.ps1)
  check97.bat をダブルクリックすると実行され、結果 r_check97.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  考え方: 総合判定(編集) 11000 には「各判定のうち最も重いもの」が入る (check55 で確認済み)。
          3 項目の判定のうち最も重いものが、今の総合判定より重ければ、その人は総合判定が変わるはず。
          判定の重さ: A=1 B=2 C=3 D=4 E=5 F=6 G=7 J=7 (空は 0)
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check97.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: HbA1c・血小板・eGFR の判定で 総合判定 が変わる人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$RANK = "CASE {0} WHEN 'A' THEN 1 WHEN 'B' THEN 2 WHEN 'C' THEN 3 WHEN 'D' THEN 4 WHEN 'E' THEN 5 WHEN 'F' THEN 6 WHEN 'G' THEN 7 WHEN 'J' THEN 7 ELSE 0 END"
$R_A1C = $RANK -f 'b.A1C'
$R_PLT = $RANK -f 'b.PLT'
$R_EGF = $RANK -f 'b.EGF'
$R_SG  = $RANK -f 'b.SOGO'
# GROUP BY でも使えるように、サブクエリではなく CASE で 3 つの最大を取る
$R_MAX = "(CASE WHEN $R_A1C >= $R_PLT AND $R_A1C >= $R_EGF THEN $R_A1C WHEN $R_PLT >= $R_EGF THEN $R_PLT ELSE $R_EGF END)"
$LETTER = "CASE {0} WHEN 7 THEN 'F以上' WHEN 6 THEN 'F' WHEN 5 THEN 'E' WHEN 4 THEN 'D' WHEN 3 THEN 'C' WHEN 2 THEN 'B' WHEN 1 THEN 'A' ELSE '' END"
$MAX3 = $LETTER -f $R_MAX
$BASE = @"
(SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS COURSE,
        g.KANJI_SIMEI,
        LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) AS A1C, LTRIM(RTRIM(ISNULL(a.KEKKA,''))) AS A1C_V,
        LTRIM(RTRIM(ISNULL(p.HANTEI_KIGO,''))) AS PLT, LTRIM(RTRIM(ISNULL(p.KEKKA,''))) AS PLT_V,
        LTRIM(RTRIM(ISNULL(e.HANTEI_KIGO,''))) AS EGF, LTRIM(RTRIM(ISNULL(e.KEKKA,''))) AS EGF_V,
        LTRIM(RTRIM(ISNULL(sg.HANTEI_KIGO,''))) AS SOGO,
        CASE WHEN EXISTS (SELECT 1 FROM T_KENSA x JOIN T_KOMOKU xm ON LTRIM(RTRIM(xm.KOMOKU_CD)) = LTRIM(RTRIM(x.KOMOKU_CD))
                          WHERE x.PK_SEQ = s.PK_SEQ AND xm.MEISYO1 LIKE N'%BMI%' AND LTRIM(RTRIM(ISNULL(x.KEKKA,''))) <> '') THEN 1 ELSE 0 END AS JUDGED
 FROM T_KENSIN s
 LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
 LEFT JOIN T_KENSA a  ON a.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD))  = '068689'
 LEFT JOIN T_KENSA p  ON p.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(p.KOMOKU_CD))  = '069010'
 LEFT JOIN T_KENSA e  ON e.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(e.KOMOKU_CD))  = '068676'
 LEFT JOIN T_KENSA sg ON sg.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(sg.KOMOKU_CD)) = '11000'
 WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831) b
"@

Q '--- 1. 今の総合判定 (11000) の分布 (受診日ごと)。空 = 自動判定がまだ、または総合判定の枠に判定が無い ---' @"
SELECT b.YMD AS 受診日, CASE WHEN b.SOGO = '' THEN '(空)' ELSE b.SOGO END AS 総合判定, COUNT(*) AS 人数,
       SUM(b.JUDGED) AS うちBMIあり
FROM $BASE GROUP BY b.YMD, b.SOGO ORDER BY 1, 2
"@ 40

Q '--- 2. ★★総合判定が今より重くなるはずの人 (3 項目の最も重い判定 > 今の総合判定。今の総合判定が空の人は除く) ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.COURSE AS コース, b.KANJI_SIMEI AS 氏名,
       b.A1C_V AS HbA1c, b.A1C AS 判定, b.PLT_V AS 血小板, b.PLT AS 判定, b.EGF_V AS eGFR, b.EGF AS 判定,
       b.SOGO AS 今の総合判定, $MAX3 AS 後の総合判定
FROM $BASE
WHERE b.SOGO <> '' AND $R_MAX > $R_SG
ORDER BY b.YMD, b.UKE
"@ 250

Q '--- 3. 変わる人数のまとめ (今の総合判定 → 後の総合判定) ---' @"
SELECT b.SOGO AS 今の総合判定, CASE WHEN $R_MAX > $R_SG THEN $MAX3 ELSE b.SOGO END AS 後の総合判定, COUNT(*) AS 人数
FROM $BASE WHERE b.SOGO <> ''
GROUP BY b.SOGO, CASE WHEN $R_MAX > $R_SG THEN $MAX3 ELSE b.SOGO END
ORDER BY 1, 2
"@ 40

Q '--- 4. 今の総合判定が空の人 (自動判定がまだ、など)。3 項目の最も重い判定だけ出す。自動判定のあとにもう一度この check を押す ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.COURSE AS コース, b.KANJI_SIMEI AS 氏名,
       b.A1C AS HbA1c判定, b.PLT AS 血小板判定, b.EGF AS eGFR判定, $MAX3 AS 三項目の最重, b.JUDGED AS BMIあり
FROM $BASE
WHERE b.SOGO = ''
ORDER BY b.YMD, b.UKE
"@ 250

Q '--- 5. 参考: 3 項目が入っている「くくり判定」(600 番台) の枠と判定の分布 (糖代謝・血液・腎のくくりがどれか確認用) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS くくり判定, ISNULL(k.HANTEI_KIGO,'') AS 判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '600%'
  AND (m.MEISYO1 LIKE N'%糖%' OR m.MEISYO1 LIKE N'%血液%' OR m.MEISYO1 LIKE N'%腎%' OR m.MEISYO1 LIKE N'%血球%' OR m.MEISYO1 LIKE N'%貧血%')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), ISNULL(k.HANTEI_KIGO,'')
ORDER BY 1, 3
"@ 60

W ''
W '=== 読み方 ==='
W '  2: ここに出た人は、3 項目のどれかが今の総合判定より重い。健診ナビの自動判定が判定欄を拾い直して総合判定を上げるか、'
W '     拾わなければ結果入力画面で総合判定を「後の総合判定」に手で直す。'
W '  3: 何人がどう変わるかの数。'
W '  4: 総合判定がまだ付いていない人 (9/10 の自動判定前の人など)。自動判定のあとに check97 を押し直すと 2 に出る。'
W '  5: 糖代謝・血液・腎のくくり判定も同じ理屈で変わる可能性がある。くくりの項目CDが分かれば次の check で同じ計算をする。'
notepad $out
