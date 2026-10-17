<#
  石屋: HbA1c・血小板・eGFR の判定で「くくり判定」(糖代謝・血液一般・腎機能) と総合判定が変わるはずの人を拾う (check98.ps1)
  check98.bat をダブルクリックすると実行され、結果 r_check98.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  対象: 総合判定 (11000) が付いている人だけ (自動判定前の人は対象外。ユーザー指示 2026/10/17)。
  くくり: HbA1c 068689 → 【判定】糖代謝 600340 / 血小板 069010 → 【判定】血液一般 600310 / eGFR 068676 → 【判定】腎機能 600220
  判定の重さ: A=1 B=2 C=3 D=4 E=5 F=6 G=7 J=7 (空は 0)
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check98.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: HbA1c・血小板・eGFR の判定で くくり判定・総合判定 が変わる人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$RANK = "CASE {0} WHEN 'A' THEN 1 WHEN 'B' THEN 2 WHEN 'C' THEN 3 WHEN 'D' THEN 4 WHEN 'E' THEN 5 WHEN 'F' THEN 6 WHEN 'G' THEN 7 WHEN 'J' THEN 7 ELSE 0 END"
$R_A1C = $RANK -f 'b.A1C'; $R_PLT = $RANK -f 'b.PLT'; $R_EGF = $RANK -f 'b.EGF'
$R_TOU = $RANK -f 'b.TOU'; $R_KET = $RANK -f 'b.KET'; $R_JIN = $RANK -f 'b.JIN'; $R_SG = $RANK -f 'b.SOGO'
$R_MAX = "(CASE WHEN $R_A1C >= $R_PLT AND $R_A1C >= $R_EGF THEN $R_A1C WHEN $R_PLT >= $R_EGF THEN $R_PLT ELSE $R_EGF END)"
$LETTER = "CASE {0} WHEN 7 THEN 'G' WHEN 6 THEN 'F' WHEN 5 THEN 'E' WHEN 4 THEN 'D' WHEN 3 THEN 'C' WHEN 2 THEN 'B' WHEN 1 THEN 'A' ELSE '' END"
$MAX3 = $LETTER -f $R_MAX
$BASE = @"
(SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS COURSE,
        g.KANJI_SIMEI,
        LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) AS A1C, LTRIM(RTRIM(ISNULL(a.KEKKA,''))) AS A1C_V,
        LTRIM(RTRIM(ISNULL(p.HANTEI_KIGO,''))) AS PLT, LTRIM(RTRIM(ISNULL(p.KEKKA,''))) AS PLT_V,
        LTRIM(RTRIM(ISNULL(e.HANTEI_KIGO,''))) AS EGF, LTRIM(RTRIM(ISNULL(e.KEKKA,''))) AS EGF_V,
        LTRIM(RTRIM(ISNULL(t.HANTEI_KIGO,''))) AS TOU, LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS KET, LTRIM(RTRIM(ISNULL(j.HANTEI_KIGO,''))) AS JIN,
        CASE WHEN t.PK_SEQ IS NULL THEN 0 ELSE 1 END AS HAS_TOU, CASE WHEN k.PK_SEQ IS NULL THEN 0 ELSE 1 END AS HAS_KET, CASE WHEN j.PK_SEQ IS NULL THEN 0 ELSE 1 END AS HAS_JIN,
        LTRIM(RTRIM(ISNULL(sg.HANTEI_KIGO,''))) AS SOGO
 FROM T_KENSIN s
 LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
 LEFT JOIN T_KENSA a  ON a.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD))  = '068689'
 LEFT JOIN T_KENSA p  ON p.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(p.KOMOKU_CD))  = '069010'
 LEFT JOIN T_KENSA e  ON e.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(e.KOMOKU_CD))  = '068676'
 LEFT JOIN T_KENSA t  ON t.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(t.KOMOKU_CD))  = '600340'
 LEFT JOIN T_KENSA k  ON k.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD))  = '600310'
 LEFT JOIN T_KENSA j  ON j.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD))  = '600220'
 LEFT JOIN T_KENSA sg ON sg.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(sg.KOMOKU_CD)) = '11000'
 WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
   AND LTRIM(RTRIM(ISNULL(sg.HANTEI_KIGO,''))) <> '') b
"@

Q '--- 1. 対象人数 (総合判定が付いている人) と、くくり判定の枠がある人数 ---' @"
SELECT COUNT(*) AS 対象人数, SUM(b.HAS_TOU) AS 糖代謝の枠あり, SUM(b.HAS_KET) AS 血液一般の枠あり, SUM(b.HAS_JIN) AS 腎機能の枠あり FROM $BASE
"@ 5

Q '--- 2. ★★糖代謝 (600340) が HbA1c の判定より軽い人 → 糖代謝を HbA1c の判定に上げる ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.KANJI_SIMEI AS 氏名, b.A1C_V AS HbA1c, b.A1C AS HbA1c判定,
       b.TOU AS 今の糖代謝, b.A1C AS 後の糖代謝, b.SOGO AS 今の総合判定
FROM $BASE WHERE b.HAS_TOU = 1 AND $R_A1C > $R_TOU
ORDER BY b.YMD, b.UKE
"@ 250

Q '--- 3. ★★血液一般 (600310) が 血小板 の判定より軽い人 → 血液一般を 血小板 の判定に上げる ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.KANJI_SIMEI AS 氏名, b.PLT_V AS 血小板, b.PLT AS 血小板判定,
       b.KET AS 今の血液一般, b.PLT AS 後の血液一般, b.SOGO AS 今の総合判定
FROM $BASE WHERE b.HAS_KET = 1 AND $R_PLT > $R_KET
ORDER BY b.YMD, b.UKE
"@ 250

Q '--- 4. ★★腎機能 (600220) が eGFR の判定より軽い人 → 腎機能を eGFR の判定に上げる ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.KANJI_SIMEI AS 氏名, b.EGF_V AS eGFR, b.EGF AS eGFR判定,
       b.JIN AS 今の腎機能, b.EGF AS 後の腎機能, b.SOGO AS 今の総合判定
FROM $BASE WHERE b.HAS_JIN = 1 AND $R_EGF > $R_JIN
ORDER BY b.YMD, b.UKE
"@ 250

Q '--- 5. ★総合判定 (11000) が 3 項目の最も重い判定より軽い人 → 総合判定を上げる ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.KANJI_SIMEI AS 氏名,
       b.A1C AS HbA1c判定, b.PLT AS 血小板判定, b.EGF AS eGFR判定, b.SOGO AS 今の総合判定, $MAX3 AS 後の総合判定
FROM $BASE WHERE $R_MAX > $R_SG
ORDER BY b.YMD, b.UKE
"@ 250

Q '--- 6. まとめ: 直すところの件数 ---' @"
SELECT N'糖代謝' AS 区分, COUNT(*) AS 人数 FROM $BASE WHERE b.HAS_TOU = 1 AND $R_A1C > $R_TOU
UNION ALL SELECT N'血液一般', COUNT(*) FROM $BASE WHERE b.HAS_KET = 1 AND $R_PLT > $R_KET
UNION ALL SELECT N'腎機能', COUNT(*) FROM $BASE WHERE b.HAS_JIN = 1 AND $R_EGF > $R_JIN
UNION ALL SELECT N'総合判定', COUNT(*) FROM $BASE WHERE $R_MAX > $R_SG
UNION ALL SELECT N'どれか 1 つでも直す人', COUNT(*) FROM $BASE
  WHERE (b.HAS_TOU = 1 AND $R_A1C > $R_TOU) OR (b.HAS_KET = 1 AND $R_PLT > $R_KET) OR (b.HAS_JIN = 1 AND $R_EGF > $R_JIN) OR $R_MAX > $R_SG
"@ 10

Q '--- 7. ★★1 人 1 行: 直すところがある人の一覧 (画面で直すときの作業表) ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.KANJI_SIMEI AS 氏名,
       CASE WHEN b.HAS_TOU = 1 AND $R_A1C > $R_TOU THEN b.TOU + N'→' + b.A1C ELSE '' END AS 糖代謝,
       CASE WHEN b.HAS_KET = 1 AND $R_PLT > $R_KET THEN b.KET + N'→' + b.PLT ELSE '' END AS 血液一般,
       CASE WHEN b.HAS_JIN = 1 AND $R_EGF > $R_JIN THEN b.JIN + N'→' + b.EGF ELSE '' END AS 腎機能,
       CASE WHEN $R_MAX > $R_SG THEN b.SOGO + N'→' + $MAX3 ELSE '' END AS 総合判定
FROM $BASE
WHERE (b.HAS_TOU = 1 AND $R_A1C > $R_TOU) OR (b.HAS_KET = 1 AND $R_PLT > $R_KET) OR (b.HAS_JIN = 1 AND $R_EGF > $R_JIN) OR $R_MAX > $R_SG
ORDER BY b.YMD, b.UKE
"@ 250

W ''
W '=== 読み方 ==='
W '  7 が作業表。「A→D」はその欄を A から D に直す、の意味。空欄は直さなくてよい。'
W '  自動判定をかけ直したあとにもう一度押して、7 が空になっていれば健診ナビが拾っている。残っていれば結果入力画面で直す。'
W '  総合判定が空の人 (自動判定前) は対象外 (ユーザー指示 2026/10/17)。'
notepad $out
