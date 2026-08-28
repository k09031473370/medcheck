<#
  団体グループ(組合)の使われ方を調べる (check18.ps1)
  check18.bat をダブルクリックすると実行され、結果 r_check18.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check17 で分かったこと:
    T_DANTAIGROUP (団体グループ) があり、M_DANTAI.DANTAIGROUP_CD で紐づく。
    → 「団体グループ = 組合 / 団体 = 加盟会社」で登録できる構造。

  ここで確かめること:
    すでにこの仕組みを使っている組合があるか。あれば、その通りに真似るのが一番安全。
    新島の会社が40社ほど登録されているので、それがグループになっているかを見る。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check18.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 団体グループ(組合)の使われ方 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 登録されている団体グループの一覧 (これが「組合」にあたる) ---' @"
SELECT LTRIM(RTRIM(DANTAIGROUP_CD)) AS グループCD,
       DANTAIGROUP_MEISYO AS グループ名,
       DANTAIGROUP_RYAKU  AS 略称,
       DANTAIGROUP_FLG    AS FLG
FROM T_DANTAIGROUP
ORDER BY DANTAIGROUP_CD
"@ 100

Q '--- 2. グループごとに何社ぶら下がっているか ---' @"
SELECT LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,''))) AS グループCD,
       MAX(g.DANTAIGROUP_MEISYO) AS グループ名,
       COUNT(*) AS 加盟社数
FROM M_DANTAI d
LEFT JOIN T_DANTAIGROUP g ON LTRIM(RTRIM(g.DANTAIGROUP_CD)) = LTRIM(RTRIM(d.DANTAIGROUP_CD))
GROUP BY LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,'')))
ORDER BY COUNT(*) DESC
"@ 100

Q '--- 3. 一番大きいグループの中身 (加盟会社の並び方の見本) ---' @"
SELECT TOP 60 LTRIM(RTRIM(d.DANTAI_CD1)) AS 団体CD, d.MEISYO1 AS 会社名,
       LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,''))) AS グループCD, g.DANTAIGROUP_MEISYO AS グループ名
FROM M_DANTAI d
JOIN T_DANTAIGROUP g ON LTRIM(RTRIM(g.DANTAIGROUP_CD)) = LTRIM(RTRIM(d.DANTAIGROUP_CD))
WHERE LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,''))) = (
    SELECT TOP 1 LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,'')))
    FROM M_DANTAI
    WHERE LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,''))) <> ''
    GROUP BY LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,'')))
    ORDER BY COUNT(*) DESC
)
ORDER BY d.DANTAI_CD1
"@ 80

Q '--- 4. 新島の会社はグループになっているか ---' @"
SELECT TOP 60 LTRIM(RTRIM(d.DANTAI_CD1)) AS 団体CD, d.MEISYO1 AS 会社名,
       LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,''))) AS グループCD,
       g.DANTAIGROUP_MEISYO AS グループ名
FROM M_DANTAI d
LEFT JOIN T_DANTAIGROUP g ON LTRIM(RTRIM(g.DANTAIGROUP_CD)) = LTRIM(RTRIM(d.DANTAIGROUP_CD))
WHERE d.MEISYO1 LIKE N'%新島%' OR d.MEISYO1 LIKE N'%式根島%'
ORDER BY d.DANTAI_CD1
"@ 80

Q '--- 5. グループ無し(空)の団体がどれくらいあるか ---' @"
SELECT CASE WHEN LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,''))) = '' THEN N'グループ無し' ELSE N'グループ有り' END AS 区分,
       COUNT(*) AS 団体数
FROM M_DANTAI
GROUP BY CASE WHEN LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,''))) = '' THEN N'グループ無し' ELSE N'グループ有り' END
"@ 10

Q '--- 6. 石材関係の団体がすでに登録されていないか ---' @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, MEISYO1 AS 会社名,
       LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,''))) AS グループCD
FROM M_DANTAI
WHERE MEISYO1 LIKE N'%石材%' OR MEISYO1 LIKE N'%石工%' OR MEISYO1 LIKE N'%墓%'
ORDER BY DANTAI_CD1
"@ 40

Q '--- 7. 団体コードの空き番 (新しく採番するときの参考。使用中の最大値) ---' @"
SELECT MAX(LTRIM(RTRIM(DANTAI_CD1))) AS 団体CD最大, COUNT(*) AS 団体数 FROM M_DANTAI
"@ 5

Q '--- 8. グループ単位で受診実績を集計できるか (直近1年) ---' @"
SELECT TOP 30 LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,''))) AS グループCD,
       MAX(g.DANTAIGROUP_MEISYO) AS グループ名,
       COUNT(DISTINCT s.DANTAI_CD1) AS 会社数,
       COUNT(*) AS 受診件数
FROM T_KENSIN s
JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_DANTAIGROUP g ON LTRIM(RTRIM(g.DANTAIGROUP_CD)) = LTRIM(RTRIM(d.DANTAIGROUP_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2025/08/01'
  AND LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,''))) <> ''
GROUP BY LTRIM(RTRIM(ISNULL(d.DANTAIGROUP_CD,'')))
ORDER BY COUNT(*) DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  1・2 にグループが並んでいる → この仕組みが実際に使われています。同じやり方で石材組合を作れます。'
W '  4 で新島の会社に同じグループCDが付いている → 離島の巡回健診で既にこの形を使っている前例です。'
W '     石材組合もそれに倣うのが一番安全です。'
W '  1 が空だった → 仕組みはあるが未使用。最初の1件になります。'
W '     その場合でも、団体グループを1つ作って加盟会社を紐づける形が素直です。'
W '  8 でグループ単位に集計できていれば、組合への一括請求もツール側で組めます。'

notepad $out
