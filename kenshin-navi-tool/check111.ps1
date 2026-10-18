<#
  石屋: 総合判定(編集) 11000 が自動判定の総合判定 10000 と違う人と、その理由になった判定 (check111.ps1)
  HbA1c・血糖・糖代謝 / 血小板・血液一般 / eGFR・腎機能 / 中性脂肪・脂質代謝 / 眼底 を並べる。
  check111.bat → r_check111.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check111.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 総合判定(編集) が自動判定と違う人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
function Han($cd) { return "MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) END)" }
function Val($cd) { return "MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) END)" }
$P = @"
(SELECT s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
        $(Han '10000') AS AUTO, $(Han '11000') AS EDIT,
        $(Han '068689') AS A1C, $(Han '068682') AS GLU, $(Han '600340') AS TOU,
        $(Han '069010') AS PLT, $(Han '600310') AS BLD,
        $(Han '068676') AS EGFR, $(Han '600220') AS JIN,
        $(Han 'R06-0002') AS TG, $(Han '600320') AS SISITU,
        $(Han '600060') AS EYE, $(Val '600060') AS EYEV
 FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
 LEFT JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
 WHERE $ISHIYA
 GROUP BY s.UKE_NO_KENSA, s.D_KENSIN, g.KANJI_SIMEI) p
"@

Q '--- 1. 総合判定(編集) 11000 と 自動判定 10000 が違う人 ---' @"
SELECT p.YMD AS 受診日, p.UKE AS 受付番号, p.NAME AS 氏名, ISNULL(p.AUTO,'') AS 自動総合, ISNULL(p.EDIT,'') AS 編集総合,
       ISNULL(p.A1C,'') AS HbA1c, ISNULL(p.GLU,'') AS 血糖, ISNULL(p.TOU,'') AS 糖代謝,
       ISNULL(p.PLT,'') AS 血小板, ISNULL(p.BLD,'') AS 血液一般,
       ISNULL(p.EGFR,'') AS eGFR, ISNULL(p.JIN,'') AS 腎機能,
       ISNULL(p.TG,'') AS 中性脂肪, ISNULL(p.SISITU,'') AS 脂質代謝,
       ISNULL(p.EYE,'') AS 眼底, ISNULL(p.EYEV,'') AS 眼底の文字
FROM $P
WHERE ISNULL(p.AUTO,'') <> ISNULL(p.EDIT,'')
ORDER BY p.YMD, p.UKE
"@ 200

W ''
W '=== 読み方 ==='
W '  自動総合 = 自動判定が出した総合判定 (10000)、編集総合 = 結果票に出る総合判定 (11000)。'
W '  編集総合と同じ文字が右のどの欄にあるかで、上げた理由が分かる (例: 眼底 F → 総合 F)。'
notepad $out
