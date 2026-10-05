<#
  眼底・胃部を健診ナビにどう入れられるかを調べる (check78.ps1)
  check78.bat をダブルクリックすると実行され、結果 r_check78.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  なぜ調べるか
    9/3・9/9・9/10 の芝浦巡回のファイルに 眼底 (Scheie・KW・所見 A20 等) と
    胃部X線 (部位コード・所見コード) が入っている。
    健診ナビ側に
      (a) 選択肢 (所見マスタ T_SYOKEN2) で持つ項目があるのか
      (b) 文字をそのまま入れる欄があるのか
    が分からないと、対応表に書けない。
    (a) なら所見の名前で突き合わせて入れられる (相手のコード表は要らない)。
    (b) なら文字をそのまま入れられる。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/09/03', [int]$UkeFrom = 4631, [int]$UkeTo = 4688)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check78.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 眼底・胃部の入れ方 ($Ymd 受付番号 $UkeFrom-$UkeTo) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. その日の受診者のコース (芝浦巡回のはず) ---' @"
SELECT LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名, COUNT(*) AS 人数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0 THEN 1 ELSE 0 END) AS 受付番号なし
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '$Ymd'
GROUP BY d.MEISYO1, s.COURSE_CD, c.MEISYO ORDER BY COUNT(*) DESC
"@ 20

Q '--- 2. ★その日の人の検査行のうち、眼底・胃・所見・コメントらしい項目 (どんな枠があるか) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名,
       LTRIM(RTRIM(ISNULL(m.K_KOMOKU,''))) AS 項目区分, LTRIM(RTRIM(ISNULL(m.SYOKEN_CD,''))) AS 所見CD,
       LTRIM(RTRIM(ISNULL(m.K_SET,''))) AS セット, COUNT(DISTINCT k.PK_SEQ) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 結果あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '$Ymd'
  AND (m.MEISYO1 LIKE N'%眼底%' OR m.MEISYO1 LIKE N'%胃%' OR m.MEISYO1 LIKE N'%所見%' OR m.MEISYO1 LIKE N'%コメント%'
       OR m.MEISYO1 LIKE N'%ｺﾒﾝﾄ%' OR m.MEISYO1 LIKE N'%備考%' OR m.MEISYO1 LIKE N'%Scheie%' OR m.MEISYO1 LIKE N'%KW%')
GROUP BY k.KOMOKU_CD, m.MEISYO1, m.K_KOMOKU, m.SYOKEN_CD, m.K_SET ORDER BY 1
"@ 80

Q '--- 3. 眼底の項目マスタ全部 (コースに関係なく) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名,
       LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分, LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見CD,
       LTRIM(RTRIM(ISNULL(K_SET,''))) AS セット
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%眼底%' OR MEISYO1 LIKE N'%Scheie%' OR MEISYO1 LIKE N'%ｼｪｰ%' OR MEISYO1 LIKE N'%Keith%' OR MEISYO1 LIKE N'%KW%'
   OR MEISYO1 LIKE N'%Scott%' OR MEISYO1 LIKE N'%網膜%'
ORDER BY KOMOKU_CD
"@ 60

Q '--- 4. ★眼底の所見マスタ (3 の所見CDの選択肢。A20 網膜前膜 のような名前があるか) ---' @"
SELECT LTRIM(RTRIM(y.SYOKEN_CD)) AS 所見CD, LTRIM(RTRIM(y.KEKKA_CD)) AS 結果CD,
       LTRIM(RTRIM(ISNULL(y.SYOKEN,''))) AS 所見, LTRIM(RTRIM(ISNULL(y.HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 y
WHERE LTRIM(RTRIM(y.SYOKEN_CD)) IN (SELECT DISTINCT LTRIM(RTRIM(SYOKEN_CD)) FROM T_KOMOKU
                                    WHERE MEISYO1 LIKE N'%眼底%' OR MEISYO1 LIKE N'%網膜%' OR MEISYO1 LIKE N'%Scheie%' OR MEISYO1 LIKE N'%KW%')
ORDER BY 1, 2
"@ 200

Q '--- 5. 胃部X線の項目マスタ ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 項目名,
       LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分, LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見CD
FROM T_KOMOKU
WHERE (MEISYO1 LIKE N'%胃%' AND (MEISYO1 LIKE N'%X%' OR MEISYO1 LIKE N'%Ｘ%' OR MEISYO1 LIKE N'%線%' OR MEISYO1 LIKE N'%所見%' OR MEISYO1 LIKE N'%部位%' OR MEISYO1 LIKE N'%判定%'))
ORDER BY KOMOKU_CD
"@ 60

Q '--- 6. ★胃部の所見マスタ (部位: 胃全般 / 所見: 顆粒状変化 のような名前があるか) ---' @"
SELECT LTRIM(RTRIM(y.SYOKEN_CD)) AS 所見CD, LTRIM(RTRIM(y.KEKKA_CD)) AS 結果CD,
       LTRIM(RTRIM(ISNULL(y.SYOKEN,''))) AS 所見, LTRIM(RTRIM(ISNULL(y.HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 y
WHERE LTRIM(RTRIM(y.SYOKEN_CD)) IN (SELECT DISTINCT LTRIM(RTRIM(SYOKEN_CD)) FROM T_KOMOKU WHERE MEISYO1 LIKE N'%胃%')
  AND (y.SYOKEN LIKE N'%全般%' OR y.SYOKEN LIKE N'%顆粒%' OR y.SYOKEN LIKE N'%体部%' OR y.SYOKEN LIKE N'%前庭%' OR y.SYOKEN LIKE N'%異常%' OR y.SYOKEN LIKE N'%なし%')
ORDER BY 1, 2
"@ 120

Q '--- 7. ★文字をそのまま入れる欄はあるか (項目区分ごとの件数と、文字が入っている実例) ---' @"
SELECT LTRIM(RTRIM(ISNULL(m.K_KOMOKU,''))) AS 項目区分, COUNT(*) AS 項目数,
       MIN(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 例1, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS 例2
FROM T_KOMOKU m GROUP BY m.K_KOMOKU ORDER BY 1
"@ 20

Q '--- 8. 過去に KEKKA に日本語の文章が入っている項目 (文字入力欄の実例をさがす) ---' @"
SELECT TOP 30 LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名,
       LTRIM(RTRIM(ISNULL(m.K_KOMOKU,''))) AS 項目区分, COUNT(*) AS 行数, MAX(LEN(k.KEKKA)) AS 最長
FROM T_KENSA k
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LEN(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) >= 6 AND k.KEKKA LIKE N'%[ぁ-んァ-ン一-龥]%'
GROUP BY k.KOMOKU_CD, m.MEISYO1, m.K_KOMOKU ORDER BY COUNT(*) DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  2 で、その日の人に眼底・胃部の枠があるかが分かる。枠が無ければ結果取込では入れられない。'
W '  3・4 眼底: 4 に「網膜前膜」「視神経乳頭陥凹」のような所見が並んでいれば、名前で突き合わせて入れられる。'
W '     Scheie / KW / Scott が選択肢 (0,Ⅰ,Ⅱ…) で持たれていれば、それも入れられる。'
W '  5・6 胃部: 6 に部位 (胃全般・体部…) と所見 (顆粒状変化…) があれば、2段階 (SHOKEN2) で入れられる。'
W '  7・8 文字をそのまま入れる欄があるか。8 に日本語の文章が入る項目があれば、それが文字欄。'
W '     眼底・胃部にその種類の項目があれば「文字だけ入れる」ができる。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
