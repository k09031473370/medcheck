<#
  胃部X線の所見マスタと、9/9・9/10 の人の状態を調べる (check79.ps1)
  check79.bat をダブルクリックすると実行され、結果 r_check79.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check78 の続き。
    ・胃部X線1〜3 (077300A/B/C) の所見CDは「ZK030/ZK031」と2つ並びで入っていたので、
      check78 の 6 では引けなかった。ZK030 (部位) と ZK031 (所見) を別々に引く。
    ・9/9・9/10 の人も受付番号が入っていないかを見る (9/3 は入っていなかった)。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check79.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 胃部X線の所見マスタ・9/9 9/10 の状態 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★胃部X線の部位マスタ (ZK030) ---' @"
SELECT LTRIM(RTRIM(KEKKA_CD)) AS 結果CD, LTRIM(RTRIM(ISNULL(SYOKEN,''))) AS 部位, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK030' ORDER BY 1
"@ 80

Q '--- 2. ★胃部X線の所見マスタ (ZK031) ---' @"
SELECT LTRIM(RTRIM(KEKKA_CD)) AS 結果CD, LTRIM(RTRIM(ISNULL(SYOKEN,''))) AS 所見, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK031' ORDER BY 1
"@ 200

Q '--- 3. 胃部X線1 に過去どう入っているか (部位+所見の実例) ---' @"
SELECT TOP 15 LTRIM(RTRIM(k.KEKKA)) AS KEKKA, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS KEKKA_CD,
       LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS 判定, COUNT(*) AS 件数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY k.KEKKA, k.KEKKA_CD, k.HANTEI_KIGO ORDER BY COUNT(*) DESC
"@ 20

Q '--- 4. 眼底所見右 (067052) に過去どう入っているか ---' @"
SELECT TOP 15 LTRIM(RTRIM(k.KEKKA)) AS KEKKA, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS KEKKA_CD,
       LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS 判定, COUNT(*) AS 件数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.KOMOKU_CD)) IN ('067052','067053') AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY k.KEKKA, k.KEKKA_CD, k.HANTEI_KIGO ORDER BY COUNT(*) DESC
"@ 20

Q '--- 5. 眼底KW右・ScheieH右 に過去どう入っているか (KEKKA に 0/Ⅰ が入るのか 1/2 か) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(k.KEKKA)) AS KEKKA, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS KEKKA_CD, COUNT(*) AS 件数
FROM T_KENSA k WHERE LTRIM(RTRIM(k.KOMOKU_CD)) IN ('067054','067056','067058','067062','H30-0051','H30-0053') AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY k.KOMOKU_CD, k.KEKKA, k.KEKKA_CD ORDER BY 1, 4 DESC
"@ 40

Q '--- 6. ★9/9・9/10 の受診者 (団体・コース・受付番号なし・カナなし) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, COUNT(*) AS 人数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0 THEN 1 ELSE 0 END) AS 受付番号なし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) = '' THEN 1 ELSE 0 END) AS カナなし
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/09','2026/09/10')
GROUP BY s.D_KENSIN, d.MEISYO1, s.COURSE_CD ORDER BY 1, 4 DESC
"@ 80

Q '--- 7. 9/3・9/9・9/10 の人数まとめ (ファイルは 58 / 90 / 51 人) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 予約数,
       SUM(CASE WHEN s.UKE_NO_KENSA IS NULL OR s.UKE_NO_KENSA = 0 THEN 1 ELSE 0 END) AS 受付番号なし,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) = '' THEN 1 ELSE 0 END) AS カナなし,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067052') THEN 1 ELSE 0 END) AS 眼底枠あり,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A') THEN 1 ELSE 0 END) AS 胃部枠あり
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
GROUP BY s.D_KENSIN ORDER BY 1
"@ 10

W ''
W '=== 読み方 ==='
W '  1・2 が本命。ファイルの「胃全般・胃穹窿部・胃体上部・胃体下部」と'
W '       「異常所見なし・透亮像・顆粒状変化・ひだ粗大・消化管術後・食道裂孔ヘルニア」が'
W '       同じ名前でマスタにあれば、名前で突き合わせて入れられる。'
W '  3〜5 は過去データの実物。KEKKA に何が入るかで、対応表の書き方 (コードか名前か) が決まる。'
W '  6・7 で 9/9・9/10 も受付番号が無ければ、先に SRL のファイルで受付番号を設定する。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
