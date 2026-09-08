<#
  視力の判定が今どうなっているかを調べる (check47.ps1)
  check47.bat をダブルクリックすると実行され、結果 r_check47.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  片方でも 0.6 以下の人の判定を D にしたい。
  書き込む前に、いまどこにどんな判定が入っているかを確かめます。
    ・視力の項目そのもの (067012 右裸眼 / 067013 左裸眼 / 067018 右矯正 / 067019 左矯正)
    ・【判定】視力 (600040)
  どちらに入れるべきかを、実物を見て決めます。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check47.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 視力の判定の調査 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★★全員の視力と判定 (値のある人だけ) ---' @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ,
       '[' + ISNULL(r1.KEKKA,'') + ']' AS 右裸眼, '[' + ISNULL(r1.HANTEI_KIGO,'') + ']' AS 判定R,
       '[' + ISNULL(l1.KEKKA,'') + ']' AS 左裸眼, '[' + ISNULL(l1.HANTEI_KIGO,'') + ']' AS 判定L,
       '[' + ISNULL(r2.KEKKA,'') + ']' AS 右矯正, '[' + ISNULL(r2.HANTEI_KIGO,'') + ']' AS 判定R2,
       '[' + ISNULL(l2.KEKKA,'') + ']' AS 左矯正, '[' + ISNULL(l2.HANTEI_KIGO,'') + ']' AS 判定L2,
       '[' + ISNULL(j.HANTEI_KIGO,'') + ']' AS 判定_視力,
       s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KENSA r1 ON r1.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(r1.KOMOKU_CD)) = '067012'
LEFT JOIN T_KENSA l1 ON l1.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(l1.KOMOKU_CD)) = '067013'
LEFT JOIN T_KENSA r2 ON r2.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(r2.KOMOKU_CD)) = '067018'
LEFT JOIN T_KENSA l2 ON l2.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(l2.KOMOKU_CD)) = '067019'
LEFT JOIN T_KENSA j  ON j.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD))  = '600040'
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND (LTRIM(RTRIM(ISNULL(r1.KEKKA,''))) NOT IN ('','#')
    OR LTRIM(RTRIM(ISNULL(l1.KEKKA,''))) NOT IN ('','#')
    OR LTRIM(RTRIM(ISNULL(r2.KEKKA,''))) NOT IN ('','#')
    OR LTRIM(RTRIM(ISNULL(l2.KEKKA,''))) NOT IN ('','#'))
ORDER BY g.KANA_SIMEI
"@ 95

Q '--- 2. ★【判定】視力 (600040) の分布 ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'       AS 結果, COUNT(*) AS 人数
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) = '600040'
GROUP BY k.HANTEI_KIGO, k.KEKKA
ORDER BY 3 DESC
"@ 20

Q '--- 3. ★視力の項目そのものに判定が付いているか ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 値あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('067012','067013','067018','067019','600040')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1
"@ 20

Q '--- 4. 視力の基準値マスタ (どの値でどの判定になる設定か) ---' @"
SELECT * FROM T_KIJUN2 WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('067012','067013','067018','067019')
ORDER BY KOMOKU_CD, KIJUN_CD, RENBAN
"@ 60

W ''
W '=== 読み方 ==='
W '  1 が本体です。片方でも 0.6 以下の13人がどんな判定になっているかを見てください。'
W '     (内山裕美子 / ｶｰﾀｰｸﾘｽﾄﾌｧｰ / 川島ことみ / 河野良市 / 河野勢也 /'
W '      田部井麗子 / 田村州子 / 冨安達郎 / 西納圭子 / 秦正明 / 村島太 /'
W '      安田和哉 / 山本稔)'
W ''
W '  3 で「視力の項目そのもの」と「【判定】視力」のどちらに判定が入る作りかが分かります。'
W '     判定を D にしたいのがどちらなのかを、これを見て決めます。'
W ''
W '  4 は基準値マスタです。0.6 以下がどの判定になる設定かが分かります。'
W '     もし今の設定と違う判定を入れたいなら、マスタを直すか、'
W '     こちらのツールで上書きするかの選択になります。'
W '     ※ マスタを直さずに上書きすると、次に自動判定を回したときに'
W '        元に戻ってしまう可能性があります。'

notepad $out
