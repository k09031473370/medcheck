<#
  当院の基準値設定を、学会の判定区分表と突き合わせられる形で全部出す (check59.ps1)
  check59.bat をダブルクリックすると実行され、結果 r_check59.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  使い方
    1. 学会の判定区分表(PDF)を用意する
       https://www.ningen-dock.jp/ningendock/wp-content/uploads/2025/12/f5777368d3f9f1c60802063246d4c6ea.pdf
    2. この結果を印刷して、PDFと並べて項目ごとに見比べる
    3. 当院は学会の4区分を8区分に読み替えている
         学会A → A   学会B → B   学会C → D   学会D → F
       (身体計測だけ 学会C → C のまま。これが意図的かどうかも確認点)

  8/21で A/B 以外の判定が出た項目 = 判定を悪くしている項目 を先に出す。
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check59.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 当院の基準値設定 (基準値1) と 8/21の判定 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W ''
W '学会の判定区分表(PDF)と並べて見比べてください。'
W '  https://www.ningen-dock.jp/ningendock/wp-content/uploads/2025/12/f5777368d3f9f1c60802063246d4c6ea.pdf'
W '読み替えの方針: 学会A→A / 学会B→B / 学会C→D / 学会D→F'
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 8/21で A/B 以外が出た検査項目
$BADITEMS = @"
(SELECT DISTINCT LTRIM(RTRIM(k.KOMOKU_CD)) AS CD
 FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
 WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
   AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('', 'A', 'B')
   AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '600%'
   AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
   AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
   AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%')
"@

Q '--- 1. ★★★判定を悪くしている項目の設定 (これを学会表と見比べる) ---' @"
SELECT LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 検査項目,
       CASE j.SEIBETU WHEN '0' THEN '共通' WHEN '1' THEN '男' WHEN '2' THEN '女' ELSE j.SEIBETU END AS 性別,
       j.HANTEI_KIGO AS 判定,
       ISNULL(j.SAISYO,'') + ' 〜 ' + ISNULL(j.SAIDAI,'') AS 範囲,
       ISNULL(j.HL,'') AS HL, ISNULL(j.HYOJI_YO,'') AS 表示, j.RENBAN AS 連番
FROM T_KIJUN1 j
JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE j.KIJUN_CD = '1'
  AND LTRIM(RTRIM(j.KOMOKU_CD)) IN (SELECT CD FROM $BADITEMS x)
ORDER BY m.MEISYO1, j.SEIBETU, j.RENBAN, j.HANTEI_KIGO
"@ 300

Q "--- 2. ★同じ項目の $Ymd 判定分布 (何人がどの判定か) ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 検査項目,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'A' THEN 1 ELSE 0 END) AS A,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'B' THEN 1 ELSE 0 END) AS B,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'C' THEN 1 ELSE 0 END) AS C,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'D' THEN 1 ELSE 0 END) AS D,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'F' THEN 1 ELSE 0 END) AS F,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('G','J') THEN 1 ELSE 0 END) AS GJ,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('','A','B','C','D','F','G','J') THEN 1 ELSE 0 END) AS その他
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN (SELECT CD FROM $BADITEMS x)
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 7 DESC, 6 DESC, 2
"@ 80

Q '--- 3. ★重複行のある項目 (同じ連番に2つ以上の判定が入っている) ---' @"
SELECT LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 検査項目,
       j.SEIBETU AS 性別, j.RENBAN AS 連番, COUNT(*) AS 行数,
       MIN(j.HANTEI_KIGO) AS 判定1, MAX(j.HANTEI_KIGO) AS 判定2
FROM T_KIJUN1 j
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE j.KIJUN_CD = '1'
GROUP BY LTRIM(RTRIM(j.KOMOKU_CD)), j.SEIBETU, j.RENBAN
HAVING COUNT(*) > 1 AND MIN(j.HANTEI_KIGO) <> MAX(j.HANTEI_KIGO)
ORDER BY 2, 4
"@ 150

Q "--- 4. $Ymd で値はあるのに基準値の設定が無い項目 (判定が付かない) ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 検査項目,
       COUNT(*) AS 値あり人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = '' THEN 1 ELSE 0 END) AS 判定なし
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '600%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%'
  AND NOT EXISTS (SELECT 1 FROM T_KIJUN1 j WHERE j.KIJUN_CD = '1'
                    AND LTRIM(RTRIM(j.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)))
  AND NOT EXISTS (SELECT 1 FROM T_KIJUN2 j2 WHERE j2.KIJUN_CD = '1'
                    AND LTRIM(RTRIM(j2.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)))
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 3 DESC
"@ 80

Q '--- 5. 参考: 判定記号の意味 ---' @"
SELECT * FROM T_HANTEIS ORDER BY 1
"@ 30

W ''
W '=== 見比べ方 ==='
W '  1 を印刷して、学会のPDFと項目ごとに並べてください。'
W '  当院は 学会C→D / 学会D→F に読み替えているので、'
W '    学会で C の範囲 → 当院で D になっていれば正しい'
W '    学会で D の範囲 → 当院で F になっていれば正しい'
W '  ずれている項目があれば、そこが「判定が悪く出すぎている(または軽すぎる)」原因です。'
W ''
W '  2 で、その項目が実際に何人を D や F にしているかが分かります。'
W '     人数の多い項目から見ると効率的です。'
W ''
W '  3 は同じ連番に判定が2つ入っている箇所です。どちらが効くかが不定なので、'
W '     江東微研に確認した方がよい箇所です。'
W ''
W '  4 は値があるのに判定が付いていない項目です。設定漏れの可能性があります。'
notepad $out
