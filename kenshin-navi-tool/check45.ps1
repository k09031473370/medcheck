<#
  判定が健診ナビの記号に直ったかを確かめる (check45.ps1)
  check45.bat をダブルクリックすると実行され、結果 r_check45.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  東振協の判定と健診ナビの判定記号は別物です。
    東振協 C (要観察)     → 健診ナビ D
    東振協 D (要精密検査) → 健診ナビ F
  取込をやり直したあと、ちゃんと D・F になっているかを見ます。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check45.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 判定の変換の確認 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★★判定記号の分布 (C が残っていないか) ---' @"
SELECT CASE LTRIM(RTRIM(k.KOMOKU_CD))
            WHEN '077010A' THEN N'胸部X線'
            WHEN '077300A' THEN N'胃部X線'
            WHEN '067112A' THEN N'心電図'
            WHEN '017101A' THEN N'他覚所見' END AS 検査,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定, COUNT(*) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','077300A','067112A','017101A')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), k.HANTEI_KIGO
ORDER BY 1, 3 DESC
"@ 40

Q '--- 2. ★D または F の人の一覧 (16人・18件になるはず) ---' @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ,
       CASE LTRIM(RTRIM(k.KOMOKU_CD))
            WHEN '077010A' THEN N'胸部X線'
            WHEN '077300A' THEN N'胃部X線'
            WHEN '067112A' THEN N'心電図'
            WHEN '017101A' THEN N'他覚所見' END AS 検査,
       LEFT(k.KEKKA, 24) AS 所見,
       '[' + LTRIM(RTRIM(k.HANTEI_KIGO)) + ']' AS 判定,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','077300A','067112A','017101A')
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('D','F')
ORDER BY 3, g.KANA_SIMEI
"@ 40

Q '--- 3. ★C が残っている人 (ここに出たら変換が効いていません) ---' @"
SELECT g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       LEFT(k.KEKKA, 24) AS 所見,
       '[' + LTRIM(RTRIM(k.HANTEI_KIGO)) + ']' AS 判定,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','077300A','067112A','017101A')
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'C'
ORDER BY 2, g.KANA_SIMEI
"@ 40

Q '--- 4. その16人の各科の判定 (自動判定を回し直したあとに揃うはず) ---' @"
SELECT g.KANJI_SIMEI AS 氏名,
       '[' + ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                      WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '600130'), '') + ']' AS 判定_心電図,
       '[' + ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                      WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '600160'), '') + ']' AS 判定_消化器,
       '[' + ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                      WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '600010'), '') + ']' AS 判定_内科診察,
       '[' + ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) FROM T_KENSA a
                      WHERE a.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD)) = '11000'), '') + ']' AS 総合判定,
       CONVERT(varchar(20), s.D_JIDOHANTEI, 120) AS 自動判定日,
       s.PK_SEQ
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
               AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','077300A','067112A','017101A')
               AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('D','F'))
ORDER BY g.KANA_SIMEI
"@ 30

W ''
W '=== 読み方 ==='
W '  1 に [C] が出ていなければ、変換は効いています。'
W '     出るはずの数はこれです。'
W '       胸部X線  A 86 / B 2'
W '       胃部X線  A 57 / B 6 / D 3 / F 5'
W '       心電図   A 66 / B 16 / D 5 / F 1'
W '       他覚所見 A 84 / D 4'
W ''
W '  2 が16人・18件なら狙いどおりです。'
W ''
W '  3 に名前が出たら変換が効いていません。書込をやり直す必要があります。'
W ''
W '  4 はそのあとの作業用です。この16人の自動判定を回し直すと、'
W '     各科の判定と総合判定が新しい判定に合わせて出し直されます。'

notepad $out
