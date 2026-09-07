<#
  過去の巡回健診の日は受付番号がどうなっているかを調べる (check39.ps1)
  check39.bat をダブルクリックすると実行され、結果 r_check39.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  8/21 は 98人全員の受付番号が空で、T_KANJA_G にも受付の行がありませんでした。
  これが「受付処理をしていないから」なのか、
  「巡回健診では元からそうなのか」を確かめます。

  過去の巡回の日と見くらべれば分かります。
#>
[CmdletBinding()]
param(
    [string]$From = '2026/04/01'   # ここから今日までを見る
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check39.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 巡回の日の受付番号の調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★受診日ごとの受付番号の状態 (人数の多い日 = 巡回) ---' @"
SELECT x.受診日, x.団体, COUNT(*) AS 人数,
       SUM(CASE WHEN x.受付No = ''  THEN 1 ELSE 0 END) AS 受付No空,
       SUM(CASE WHEN x.受付行数 = 0 THEN 1 ELSE 0 END) AS 受付行なし,
       MIN(CASE WHEN x.受付No = '' THEN NULL ELSE x.受付N END) AS 受付No最小,
       MAX(x.受付N) AS 受付No最大
FROM (
  SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
         LEFT(ISNULL(d.MEISYO1, '(院内)'), 16) AS 団体,
         LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), ''))) AS 受付No,
         s.UKE_NO_KENSA AS 受付N,
         (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付行数
  FROM T_KENSIN s
  LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
  WHERE s.D_KENSIN >= '$From' AND s.F_TORIKESI = 0
) x
GROUP BY x.受診日, x.団体
HAVING COUNT(*) >= 10
ORDER BY x.受診日 DESC
"@ 80

Q '--- 2. ★人数の多い日(巡回)の中身を1日だけ見る: 直近で受付Noが入っている団体の日 ---' @"
SELECT TOP 25 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       ISNULL(d.MEISYO1,'') AS 団体,
       g.KANJI_SIMEI AS 氏名,
       '[' + ISNULL(s.SEQ1,'') + ']' AS SEQ1,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), '') + ']' AS 受付No,
       (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付行数
FROM T_KENSIN s
LEFT JOIN M_DANTAI d  ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_KOJIN1 g  ON g.KOJIN_ID   = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND s.UKE_NO_KENSA IS NOT NULL
  AND s.DANTAI_CD1 IS NOT NULL AND s.DANTAI_CD1 <> ''
  AND s.D_KENSIN = (
      SELECT MAX(t.D_KENSIN) FROM T_KENSIN t
      WHERE t.F_TORIKESI = 0 AND t.UKE_NO_KENSA IS NOT NULL
        AND t.DANTAI_CD1 IS NOT NULL AND t.DANTAI_CD1 <> ''
        AND t.D_KENSIN < '2026/08/21')
ORDER BY s.UKE_NO_KENSA
"@ 30

Q '--- 3. 城西学園の過去の受診日 (同じ団体の去年はどうだったか) ---' @"
SELECT x.受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN x.受付No = ''  THEN 1 ELSE 0 END) AS 受付No空,
       SUM(CASE WHEN x.受付行数 = 0 THEN 1 ELSE 0 END) AS 受付行なし,
       MAX(x.受付N) AS 受付No最大
FROM (
  SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
         LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA), ''))) AS 受付No,
         s.UKE_NO_KENSA AS 受付N,
         (SELECT COUNT(*) FROM T_KANJA_G q WHERE q.PK_SEQ = s.PK_SEQ) AS 受付行数
  FROM T_KENSIN s
  WHERE s.F_TORIKESI = 0
    AND s.DANTAI_CD1 = (SELECT TOP 1 t.DANTAI_CD1 FROM T_KENSIN t
                         WHERE t.D_KENSIN = '2026/08/21' AND t.F_TORIKESI = 0)
) x
GROUP BY x.受診日
ORDER BY x.受診日 DESC
"@ 30

W ''
W '=== 読み方 ==='
W '  1 で、8/21 以外の人数の多い日 (巡回) を見てください。'
W '     どの日も「受付No空」が 0 なら → 巡回でも受付番号は入るのが普通。'
W '     8/21 だけ 98 なら、8/21 の受付処理が通っていないということです。'
W ''
W '     逆に、巡回の日はどこも「受付No空 = 人数」なら → 巡回では元から空。'
W '     その場合は無理に入れない方がいいです。'
W ''
W '  3 は同じ城西学園の過去の日です。去年どうだったかが、一番の答えになります。'

notepad $out
