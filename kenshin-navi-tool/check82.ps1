<#
  4742 西野 大基 / 4795 相澤 貢 のコースを調べる (check82.ps1)
  check82.bat をダブルクリックすると実行され、結果 r_check82.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  SRL・910 の取込で、4742 は中性脂肪・眼底の枠なし、4795 は血液・尿・問診・心電図・眼底の枠なし
  (枠が 15 行しかない) になった。健診ナビの予約のコースが違っている疑いがあるので、
    1) 2 人の今のコースと枠の数
    2) 同じ日・同じ団体の他の人がどのコースで、枠がいくつあるか
    3) その団体に登録されているコース一覧
  を出して、予約画面で選び直すコースを決めるためのもの。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check82.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 4742・4795 のコース確認 (9/3・9/9・9/10) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★2 人の今の予約 (受診日・団体・コース・枠の数・結果が入っている枠の数) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, s.PK_SEQ,
       LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 漢字, LTRIM(RTRIM(ISNULL(g.KANA_SIMEI,''))) AS カナ,
       LTRIM(RTRIM(ISNULL(s.DANTAI_CD1,''))) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS 枠数,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS 結果あり
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA IN (4742, 4795)
ORDER BY 2
"@ 10

Q '--- 2. ★同じ 3 日間の人をコースごとに (人数・枠数・主な検査の枠の有無) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       COUNT(*) AS 人数,
       MIN((SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ)) AS 枠数min,
       MAX((SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ)) AS 枠数max,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '068689') THEN 1 ELSE 0 END) AS HbA1c枠,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069010') THEN 1 ELSE 0 END) AS 血小板枠,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069245') THEN 1 ELSE 0 END) AS 便潜血枠,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067052') THEN 1 ELSE 0 END) AS 眼底枠,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A') THEN 1 ELSE 0 END) AS 胃部枠,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A') THEN 1 ELSE 0 END) AS 心電図枠
FROM T_KENSIN s
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
GROUP BY s.D_KENSIN, d.MEISYO1, s.COURSE_CD, c.MEISYO
ORDER BY 1, 5 DESC
"@ 40

Q '--- 3. 2 人と同じ団体に登録されているコース一覧 (予約画面のコース欄に出る候補) ---' @"
SELECT LTRIM(RTRIM(ISNULL(c.DANTAI_CD1,''))) AS 団体CD, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 団体,
       LTRIM(RTRIM(ISNULL(c.COURSE_CD,''))) AS コース, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       (SELECT COUNT(*) FROM T_KENSIN s2 WHERE s2.F_TORIKESI = 0 AND s2.DANTAI_CD1 = c.DANTAI_CD1 AND s2.COURSE_CD = c.COURSE_CD
          AND s2.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')) AS 今回の人数
FROM T_COURSE1 c
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = c.DANTAI_CD1
WHERE c.DANTAI_CD1 IN (SELECT s.DANTAI_CD1 FROM T_KENSIN s WHERE s.F_TORIKESI = 0
                        AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA IN (4742, 4795))
ORDER BY 1, 3
"@ 40

Q '--- 4. 2 人の今の枠 (項目CD・項目名・結果) : どの検査の枠が有って何が無いかを見る ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 項目名, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS 結果
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA IN (4742, 4795)
ORDER BY 1, 2
"@ 200

W ''
W '=== 読み方 ==='
W '  1: 2 人の今のコースと枠数。4742 は 103 枠・4795 は 15 枠のはず。結果あり が 0 なら、まだ何も入っていないので'
W '     予約画面でコースを変えても消えるものは無い。'
W '  2: 同じ日の他の人が何というコースか。人数が一番多いコース (枠数 100 前後・眼底枠/心電図枠 あり) が'
W '     その団体の普通のコース。2 人のコースがそれと違っていれば、予約画面でそのコースに変える。'
W '     ※ 血小板・HbA1c・便潜血は人によって有無が違う (年齢やオプション) ので、ここが違うだけなら直さなくてよい。'
W '  3: 予約画面のコース欄で選べる候補。2 で決めたコース名をここで確かめる。'
W '  4: 4795 は 15 枠しか無いので、どの検査だけ入っているコースかが分かる (身体計測だけ、など)。'
W ''
W '  直し方 (健診ナビ): 予約者一覧 → その人を開く → コース欄を 2 で決めたコースに変えて保存。'
W '  健診ナビが新しいコースの枠を作り直すので、そのあと start.bat で SRL と 910 を 3→4 でもう一度。'
notepad $out
