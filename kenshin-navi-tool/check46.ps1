<#
  尿蛋白(+)で判定が D1 になっている人を出す (check46.ps1)
  check46.bat をダブルクリックすると実行され、結果 r_check46.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  基準値マスタ(069206)の (+) には C と D しか登録されていないのに、
  自動判定のあと D1 が付いた人がいました。誰なのかを確かめます。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check46.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 尿蛋白の判定 D1 の調査 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★★尿蛋白の判定が A 以外の人 ---' @"
SELECT g.KANJI_SIMEI AS 氏名, g.KANA_SIMEI AS カナ,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'       AS 尿蛋白,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']'    AS 結果CD,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       '[' + LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), k.HANTEI_CD),''))) + ']' AS 判定CD,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069206'
  AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) NOT IN ('', 'A')
ORDER BY g.KANA_SIMEI
"@ 20

Q '--- 2. ★その人の腎臓まわりの値 (D1 になる理由の手がかり) ---' @"
SELECT g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'       AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND s.PK_SEQ IN (SELECT k2.PK_SEQ FROM T_KENSA k2 JOIN T_KENSIN s2 ON s2.PK_SEQ = k2.PK_SEQ
                    WHERE s2.D_KENSIN = '$Ymd' AND s2.F_TORIKESI = 0
                      AND LTRIM(RTRIM(k2.KOMOKU_CD)) = '069206'
                      AND LTRIM(RTRIM(ISNULL(k2.HANTEI_KIGO,''))) NOT IN ('', 'A'))
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN
      ('069206','069207','069211','068673','068674','068675','600220','600070','11000','10000')
ORDER BY k.KOMOKU_CD
"@ 30

Q '--- 3. D1 という判定記号が過去にどこで使われているか ---' @"
SELECT TOP 30 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']' AS 結果,
       s.PK_SEQ
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'D1'
ORDER BY s.D_KENSIN DESC
"@ 35

Q '--- 4. 判定記号のマスタに D1 は登録されているか ---' @"
SELECT * FROM T_HANTEIS
"@ 60

W ''
W '=== 読み方 ==='
W '  1 でその人の名前が分かります。'
W ''
W '  2 でその人の腎臓まわりの値が見えます。'
W '     尿蛋白だけでなく、クレアチニンや尿潜血も合わせて'
W '     判定が上がる作りなら、D1 は自動判定が出した正常な値です。'
W ''
W '  3 で、D1 が過去にも使われているかが分かります。'
W '     ほかの日にもあるなら、この健診ナビで普通に使われている記号です。'
W ''
W '  4 は判定記号のマスタです。D1 が登録されていれば、'
W '     少なくとも「あり得ない記号」ではないということになります。'

notepad $out
