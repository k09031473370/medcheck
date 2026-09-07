<#
  誤ったキー入力が結果欄に落ちていないか探す (check42.ps1)
  check42.bat をダブルクリックすると実行され、結果 r_check42.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  自動判定の確認を消すための「y」キーが、ダイアログではなく
  結果の入力欄に入ってしまうことがあります。
  8/21 の尿蛋白 3人で実際に起きました ((-) が y に化けていた)。

  ※ 前の版は「英字を含む行」を全部出していましたが、
     600xxx【判定】の結果欄には判定記号(A/B/C)がそのまま入るため、
     正常な行が何千行も出てしまいました。今回はそこを外しています。
#>
[CmdletBinding()]
param(
    [string]$Ymd  = '2026/08/21',
    [string]$From = '2026/08/01'   # ここから今日までを見る
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check42.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

# 判定を入れておく項目。ここは英字(A/B/C…)が入るのが正常なので除外する。
$SkipJudge = @"
      LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '600%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '610%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1000%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1100%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE 'MT00%'
"@

"=== 誤入力の探索 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★★数値の項目なのに英字が混ざっている行 ($Ymd) ---" @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(k.KEKKA)) + ']' AS 結果,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND k.KEKKA LIKE '%[a-zA-Z]%'
  AND $SkipJudge
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN (
      SELECT LTRIM(RTRIM(k2.KOMOKU_CD))
      FROM T_KENSA k2 JOIN T_KENSIN s2 ON s2.PK_SEQ = k2.PK_SEQ
      WHERE s2.D_KENSIN = '$Ymd' AND s2.F_TORIKESI = 0
        AND ISNUMERIC(LTRIM(RTRIM(k2.KEKKA))) = 1
      GROUP BY LTRIM(RTRIM(k2.KOMOKU_CD)))
ORDER BY k.KOMOKU_CD, g.KANA_SIMEI
"@ 60

Q "--- 2. ★★結果CDとマスタの文字が食い違う行 ($Ymd・所見系すべて) ---" @"
SELECT g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'    AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + ISNULL(y.SYOKEN, '') + ']'                AS 本来の文字,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
LEFT JOIN T_SYOKEN2 y ON LTRIM(RTRIM(y.SYOKEN_CD)) = LTRIM(RTRIM(m.SYOKEN_CD))
                     AND LTRIM(RTRIM(y.KEKKA_CD))  = LTRIM(RTRIM(k.KEKKA_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) <> ''
  AND LTRIM(RTRIM(ISNULL(m.SYOKEN_CD,''))) NOT IN ('', '998')
  AND ISNULL(y.SYOKEN,'') <> ''
  AND LTRIM(RTRIM(k.KEKKA)) <> LTRIM(RTRIM(y.SYOKEN))
ORDER BY k.KOMOKU_CD, g.KANA_SIMEI
"@ 60

Q "--- 3. ★所見系の食い違いを $From 以降の全部の日で ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']' AS 結果,
       '[' + ISNULL(y.SYOKEN, '') + ']'             AS 本来の文字,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
LEFT JOIN T_SYOKEN2 y ON LTRIM(RTRIM(y.SYOKEN_CD)) = LTRIM(RTRIM(m.SYOKEN_CD))
                     AND LTRIM(RTRIM(y.KEKKA_CD))  = LTRIM(RTRIM(k.KEKKA_CD))
WHERE s.D_KENSIN >= '$From' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) <> ''
  AND LTRIM(RTRIM(ISNULL(m.SYOKEN_CD,''))) NOT IN ('', '998')
  AND ISNULL(y.SYOKEN,'') <> ''
  AND LTRIM(RTRIM(k.KEKKA)) <> LTRIM(RTRIM(y.SYOKEN))
ORDER BY s.D_KENSIN DESC, k.KOMOKU_CD
"@ 80

Q "--- 4. ★数値の項目に英字が混ざっている行 ($From 以降の全部の日) ---" @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(k.KEKKA)) + ']' AS 結果,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN >= '$From' AND s.F_TORIKESI = 0
  AND k.KEKKA LIKE '%[a-zA-Z]%'
  AND $SkipJudge
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN (
      SELECT LTRIM(RTRIM(k2.KOMOKU_CD))
      FROM T_KENSA k2 JOIN T_KENSIN s2 ON s2.PK_SEQ = k2.PK_SEQ
      WHERE s2.D_KENSIN >= '$From' AND s2.F_TORIKESI = 0
        AND ISNUMERIC(LTRIM(RTRIM(k2.KEKKA))) = 1
      GROUP BY LTRIM(RTRIM(k2.KOMOKU_CD)))
ORDER BY s.D_KENSIN DESC, k.KOMOKU_CD
"@ 80

W ''
W '=== 読み方 ==='
W '  1 は「ふだん数字が入る項目に英字が混ざっている行」です。'
W '     身長が 172.4y になっている、といった取りこぼしを拾います。'
W ''
W '  2 は「結果CDとマスタの文字が食い違う行」です。'
W '     尿蛋白の y 3件はここに出ます。'
W ''
W '  3 と 4 は同じことを8月以降の全部の日で見ます。'
W '     他の日に飛び火していないかの確認です。'
W ''
W '  600xxx【判定】の結果欄には判定記号(A/B/C)が入るのが正常なので、'
W '  そこは最初から外しています。'
W ''
W '  2 に尿蛋白の3件だけ、1・3・4 に何も出なければ、'
W '  誤入力はその3件だけということになります。'

notepad $out
