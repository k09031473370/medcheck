<#
  1 人分の判定・コメントを書き出す (check101.ps1)
  自動判定を 1 人だけかける前と後に実行して、何が変わったか (消えたか) を見比べる用。
  check101.bat をダブルクリック → 受付番号を聞かれる → r_check101_<受付番号>_<日時>.txt がメモ帳で開く。
  DBは読むだけで、一切変更しません。
#>
[CmdletBinding()]
param([string]$UkeNo)
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
if (-not $UkeNo) { $UkeNo = Read-Host '受付番号 (例 4714)' }
if ($UkeNo -notmatch '^\d+$') { Write-Host '受付番号は数字で入れてください。'; pause; return }
$out  = Join-Path $dir ("r_check101_{0}_{1}.txt" -f $UkeNo, (Get-Date -Format 'yyyyMMdd_HHmmss'))
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 受付番号 $UkeNo の判定とコメント $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$WHO = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA = $UkeNo"

Q '--- 1. 本人 ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, g.KANJI_SIMEI AS 氏名, LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, s.PK_SEQ
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID WHERE $WHO
"@ 5

Q '--- 2. ★判定が付いている項目 すべて (所見・くくり・総合判定を含む) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, LEFT(ISNULL(k.KEKKA,''), 40) AS 結果, ISNULL(k.KEKKA_CD,'') AS 結果CD, ISNULL(k.HANTEI_KIGO,'') AS 判定, ISNULL(k.HL,'') AS HL
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $WHO AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
ORDER BY k.KOMOKU_CD
"@ 300

Q '--- 3. ★気にしている項目 (判定が空でも出す): 胸部X線 / 胃部 / 心電図 / 診察 / 眼底 / HbA1c / 血小板 / eGFR / 血糖 / 中性脂肪 / くくり / 総合判定 ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名, LEFT(ISNULL(k.KEKKA,''), 40) AS 結果, ISNULL(k.KEKKA_CD,'') AS 結果CD, ISNULL(k.HANTEI_KIGO,'') AS 判定, ISNULL(k.HL,'') AS HL
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $WHO AND (LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','077010B','077300A','077300B','067112A','067112B','067112C','017101A',
   '068689','069010','068676','068682','H30-0011','H30-0012','H30-0013','R06-0002','068702','R06-0003','001282','600320','600340','600310','600220','11000')
   OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '0670[5-9]%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE 'H30-005%')
ORDER BY k.KOMOKU_CD
"@ 100

Q '--- 4. ★医師指示事項 (11001〜11009) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 枠, ISNULL(k.KEKKA_CD,'') AS 番号, ISNULL(k.KEKKA,'') AS 文章
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $WHO AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('11001','11002','11003','11004','11005','11006','11007','11008','11009')
ORDER BY k.KOMOKU_CD
"@ 20

W ''
W '=== 使い方 ==='
W '  自動判定をかける前に 1 回、かけた後に 1 回この check を押し、2 つのメモ帳を見比べる。'
W '  2 で消えた行 (判定が空になった項目)、3 で空腹時血糖・空腹時中性脂肪に値が写ったか、4 で文章が入れ替わったか/消えたか、を見る。'
notepad $out
