<#
  誤ったキー入力が結果欄に落ちていないか探す (check42.ps1)
  check42.bat をダブルクリックすると実行され、結果 r_check42.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  自動判定の確認ダイアログを消すための「y」キーが、
  ダイアログではなく結果の入力欄に入ってしまうことがあります。
  8/21 の尿蛋白 3人で実際に起きました (値が (-) から y に化けていた)。

  同じことが他の項目・他の日で起きていないかを探します。
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

"=== 誤入力の探索 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q "--- 1. ★★$Ymd で、結果に英字が混ざっている行 (y / n など) ---" @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(k.KEKKA)) + ']'               AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND k.KEKKA LIKE '%[a-zA-Z]%'
ORDER BY k.KOMOKU_CD, g.KANA_SIMEI
"@ 60

Q "--- 2. ★$From 以降の全部の日で、結果に英字が混ざっている行 ---" @"
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
ORDER BY s.D_KENSIN DESC, k.KOMOKU_CD
"@ 80

Q "--- 3. ★$Ymd の尿検査 3項目の値の一覧 (おかしな値が無いか) ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'    AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       COUNT(*) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('069206','069207','069211','069245','069246')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), k.KEKKA, k.KEKKA_CD
ORDER BY 1, 5 DESC
"@ 40

Q "--- 4. ★$Ymd で、結果CDはあるのに結果の文字が対応表と食い違う行を探す(所見系) ---" @"
SELECT g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'    AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD,
       '[' + ISNULL(y.SYOKEN, '(マスタに無い)') + ']'   AS マスタの所見文,
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
"@ 80

W ''
W '=== 読み方 ==='
W '  1 に出る行が、結果欄に英字(a〜z)が混ざっているものです。'
W '     健診の結果はふつう 数字・(-)(+)・日本語なので、英字はまず出ません。'
W '     「y」だけの行はもちろん、「172.4y」のようにくっついた行も拾います。'
W '     すでに分かっている尿蛋白の3人以外に出たら、それも直す対象です。'
W '     ごくまれに、もともと英字を含む正しい値が出ることがあります。'
W '     その場合は値を見れば区別が付きます。'
W ''
W '  2 は8月以降の全部の日を見ます。他の日にも飛び火していないかの確認です。'
W ''
W '  3 は尿の値の一覧です。(-) (+) 以外の妙な値が混ざっていないかを見ます。'
W ''
W '  4 は所見の項目で、結果CDとマスタの所見文が食い違っている行です。'
W '     ここに出れば、その行も文字だけ書き換わっています。'
W ''
W '  なお、直すのは健診ナビの画面からで大丈夫です。'
W '  値を入れ直して自動判定→登録すれば元に戻ります。'

notepad $out
