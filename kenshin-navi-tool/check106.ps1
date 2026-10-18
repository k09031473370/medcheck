<#
  石屋: 服薬中 (問診の服薬歴) なのに J (治療中) になっていない人 (check106.ps1)
  問診 084001 服薬歴(血圧) / 084002 服薬歴(血糖) / 084003 服薬歴(脂質) と、
  くくり判定 【判定】血圧 600150 / 糖代謝 600340 / 脂質代謝 600320、総合判定 11000 を並べる。
  check106.bat → r_check106.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check106.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 服薬中なのに J でない人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
function Val($cd) { return "MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) END)" }
function Han($cd) { return "MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) END)" }
$P = @"
(SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
        $(Val '084001') AS Q1, $(Val '084002') AS Q2, $(Val '084003') AS Q3,
        $(Han '600150') AS BP, $(Han '600340') AS TOU, $(Han '600320') AS SISITU, $(Han '11000') AS SOGO,
        $(Val '068689') AS A1C, $(Val '068682') AS GLU, $(Val 'R06-0002') AS TG
 FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
 LEFT JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
 WHERE $ISHIYA
 GROUP BY s.PK_SEQ, s.UKE_NO_KENSA, s.D_KENSIN, g.KANJI_SIMEI) p
"@

Q '--- 1. 服薬歴の答えの分布 (どの値が「はい」かを見る) ---' @"
SELECT '血圧' AS 服薬, ISNULL(p.Q1,'(枠なし)') AS 答え, COUNT(*) AS 人数 FROM $P GROUP BY ISNULL(p.Q1,'(枠なし)')
UNION ALL SELECT '血糖', ISNULL(p.Q2,'(枠なし)'), COUNT(*) FROM $P GROUP BY ISNULL(p.Q2,'(枠なし)')
UNION ALL SELECT '脂質', ISNULL(p.Q3,'(枠なし)'), COUNT(*) FROM $P GROUP BY ISNULL(p.Q3,'(枠なし)')
ORDER BY 1, 2
"@ 30

$YES = "IN ('1','はい','有','あり')"
Q '--- 2. ★服薬中 (どれか 1 つでも「はい」) の人と、対応するくくり判定 ---' @"
SELECT p.YMD AS 受診日, p.UKE AS 受付番号, p.NAME AS 氏名,
       CASE WHEN p.Q1 $YES THEN '服薬' ELSE '' END AS 血圧薬, ISNULL(p.BP,'') AS 血圧判定,
       CASE WHEN p.Q2 $YES THEN '服薬' ELSE '' END AS 血糖薬, ISNULL(p.TOU,'') AS 糖代謝, ISNULL(p.A1C,'') AS HbA1c, ISNULL(p.GLU,'') AS 血糖,
       CASE WHEN p.Q3 $YES THEN '服薬' ELSE '' END AS 脂質薬, ISNULL(p.SISITU,'') AS 脂質代謝, ISNULL(p.TG,'') AS 中性脂肪,
       ISNULL(p.SOGO,'') AS 総合
FROM $P
WHERE p.Q1 $YES OR p.Q2 $YES OR p.Q3 $YES
ORDER BY p.YMD, p.UKE
"@ 200

Q '--- 3. 件数: 服薬中で、対応するくくりが J でない人 (受診日別) ---' @"
SELECT p.YMD AS 受診日,
       SUM(CASE WHEN p.Q1 $YES THEN 1 ELSE 0 END) AS 血圧薬,
       SUM(CASE WHEN p.Q1 $YES AND ISNULL(p.BP,'') <> 'J' THEN 1 ELSE 0 END) AS うちJでない,
       SUM(CASE WHEN p.Q2 $YES THEN 1 ELSE 0 END) AS 血糖薬,
       SUM(CASE WHEN p.Q2 $YES AND ISNULL(p.TOU,'') <> 'J' THEN 1 ELSE 0 END) AS うちJでない_,
       SUM(CASE WHEN p.Q3 $YES THEN 1 ELSE 0 END) AS 脂質薬,
       SUM(CASE WHEN p.Q3 $YES AND ISNULL(p.SISITU,'') <> 'J' THEN 1 ELSE 0 END) AS うちJでない__
FROM $P GROUP BY p.YMD ORDER BY 1
"@ 10

Q '--- 4. 参考: 2026 年の全受診者で、くくり判定 血圧・糖代謝・脂質代謝 に J が付いている件数 (このクリニックで J を使うことがあるか) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, COUNT(*) AS J の件数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('600150','600340','600320','11000')
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'J'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)) ORDER BY 1
"@ 10

W ''
W '=== 読み方 ==='
W '  1: 「はい」にあたる値が 1 以外 (例 2) なら、2・3 の条件が合っていないので知らせてください。'
W '  2: 服薬中の人の一覧。血圧薬なのに血圧判定が D/F、血糖薬なのに糖代謝が D/F … が J にしていない人。'
W '  4: 0 件なら、このクリニックは服薬中でも J を使っていない (いつもの運用)。'
notepad $out
