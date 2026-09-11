<#
  中性脂肪 (R06-0002) に判定を付けたら「脂質代謝」「総合判定」が A から落ちる人がいるか (check60.ps1)
  check60.bat をダブルクリックすると実行され、結果 r_check60.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。(自動判定も走らせません)

  中性脂肪は今回 R06-0002 に基準値が無く、判定が付いていない。
  学会基準 (F 0〜29 / A 30〜149 / B 150〜299 / D 300〜499 / F 500以上) を当てはめたと仮定して、
  「脂質代謝」(600320) と 総合判定(編集) (11000) が今より悪くなる人を計算で出す。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check60.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 中性脂肪を判定したら A から落ちる人 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 1人1行: 中性脂肪の値・仮の判定、今の脂質代謝、今の総合判定(編集)
# 判定の重さ: A=1 B=2 C=3 D=4 E=5 F=6 G=7 J=7 (空は 0)
$BASE = @"
(SELECT s.PK_SEQ, g.KANJI_SIMEI, g.KANA_SIMEI,
        LTRIM(RTRIM(ISNULL(tg.KEKKA,''))) AS TG_VAL,
        CASE WHEN ISNUMERIC(LTRIM(RTRIM(ISNULL(tg.KEKKA,'')))) = 1
             THEN CONVERT(float, LTRIM(RTRIM(tg.KEKKA))) END AS TG_NUM,
        LTRIM(RTRIM(ISNULL(sh.HANTEI_KIGO,''))) AS SHISHITSU,
        LTRIM(RTRIM(ISNULL(sg.HANTEI_KIGO,''))) AS SOGO
 FROM T_KENSIN s
 LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
 LEFT JOIN T_KENSA tg ON tg.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(tg.KOMOKU_CD)) = 'R06-0002'
 LEFT JOIN T_KENSA sh ON sh.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(sh.KOMOKU_CD)) = '600320'
 LEFT JOIN T_KENSA sg ON sg.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(sg.KOMOKU_CD)) = '11000'
 WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0) b
"@

# 仮の中性脂肪判定 (学会基準) と、その重さ
$TGH = @"
CASE WHEN b.TG_NUM IS NULL THEN ''
     WHEN b.TG_NUM <  30 THEN 'F'
     WHEN b.TG_NUM < 150 THEN 'A'
     WHEN b.TG_NUM < 300 THEN 'B'
     WHEN b.TG_NUM < 500 THEN 'D'
     ELSE 'F' END
"@
$RANK = "CASE {0} WHEN 'A' THEN 1 WHEN 'B' THEN 2 WHEN 'C' THEN 3 WHEN 'D' THEN 4 WHEN 'E' THEN 5 WHEN 'F' THEN 6 WHEN 'G' THEN 7 WHEN 'J' THEN 7 ELSE 0 END"
$R_TG = $RANK -f "($TGH)"
$R_SH = $RANK -f 'b.SHISHITSU'
$R_SG = $RANK -f 'b.SOGO'

# 中性脂肪を入れた後の 脂質代謝 / 総合判定 (今の値と中性脂肪判定の重い方)
$NEW_SH = "CASE WHEN $R_TG > $R_SH THEN ($TGH) ELSE b.SHISHITSU END"
$NEW_SG = "CASE WHEN $R_TG > $R_SG THEN ($TGH) ELSE b.SOGO END"

Q '--- 1. 今の分布 (脂質代謝 600320 / 総合判定 11000) ---' @"
SELECT '脂質代謝' AS 区分, b.SHISHITSU AS 判定, COUNT(*) AS 人数 FROM $BASE GROUP BY b.SHISHITSU
UNION ALL
SELECT '総合判定', b.SOGO, COUNT(*) FROM $BASE GROUP BY b.SOGO
ORDER BY 1, 2
"@ 40

Q '--- 2. ★中性脂肪を判定したときの 脂質代謝 の変わり方 (今→後) ---' @"
SELECT b.SHISHITSU AS 今の脂質代謝, $NEW_SH AS 後の脂質代謝, COUNT(*) AS 人数
FROM $BASE
GROUP BY b.SHISHITSU, $NEW_SH
ORDER BY 1, 2
"@ 40

Q '--- 3. ★中性脂肪を判定したときの 総合判定 の変わり方 (今→後) ---' @"
SELECT b.SOGO AS 今の総合判定, $NEW_SG AS 後の総合判定, COUNT(*) AS 人数
FROM $BASE
GROUP BY b.SOGO, $NEW_SG
ORDER BY 1, 2
"@ 40

Q '--- 4. ★★脂質代謝 または 総合判定 が今より悪くなる人 (明細) ---' @"
SELECT b.KANJI_SIMEI AS 氏名, b.TG_VAL AS 中性脂肪, $TGH AS 中性脂肪判定,
       b.SHISHITSU AS 今の脂質代謝, $NEW_SH AS 後の脂質代謝,
       b.SOGO AS 今の総合判定, $NEW_SG AS 後の総合判定
FROM $BASE
WHERE $R_TG > $R_SH OR $R_TG > $R_SG
ORDER BY b.TG_NUM DESC
"@ 100

Q '--- 5. 参考: 中性脂肪 150以上 の人 全員 (変わらない人も含む) ---' @"
SELECT b.KANJI_SIMEI AS 氏名, b.TG_VAL AS 中性脂肪, $TGH AS 中性脂肪判定,
       b.SHISHITSU AS 今の脂質代謝, b.SOGO AS 今の総合判定,
       CASE WHEN $R_TG > $R_SH THEN '脂質代謝が落ちる' ELSE '' END AS 脂質,
       CASE WHEN $R_TG > $R_SG THEN '総合判定が落ちる' ELSE '' END AS 総合
FROM $BASE
WHERE b.TG_NUM >= 150 OR b.TG_NUM < 30
ORDER BY b.TG_NUM DESC
"@ 100

Q '--- 6. 参考: その人たちの脂質の値 (TC/HDL/LDL/nonHDL/中性脂肪) ---' @"
SELECT b.KANJI_SIMEI AS 氏名, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 検査項目,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']' AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定
FROM $BASE
JOIN T_KENSA k ON k.PK_SEQ = b.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE (b.TG_NUM >= 150 OR b.TG_NUM < 30)
  AND (m.MEISYO1 LIKE '%コレステロール%' OR m.MEISYO1 LIKE '%中性脂肪%'
       OR m.MEISYO1 LIKE '%HDL%' OR m.MEISYO1 LIKE '%LDL%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) = '600320')
ORDER BY b.TG_NUM DESC, b.KANA_SIMEI, k.KOMOKU_CD
"@ 200

W ''
W '=== 読み方 ==='
W '  2 が「脂質代謝」、3 が「総合判定」の答えです。今→後 が同じ行は変わらない人。'
W '  「今の脂質代謝 A → 後の脂質代謝 B」の行があれば、その人数が「中性脂肪のせいで A から落ちる人」です。'
W '  4 に、落ちる人の名前と値が出ます。0件なら誰も落ちません。'
W '  5・6 は確認用です。中性脂肪 150以上の人の今の判定と、脂質の他の値を並べています。'
W '  ※ 計算で出しているだけで、DBには何も書いていません。自動判定も走らせていません。'
notepad $out
