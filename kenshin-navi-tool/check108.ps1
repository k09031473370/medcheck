<#
  J (治療中) は既往歴 (T_KIOU) と「病名 → 判定」の対応表 (T_KIOU_HANTEI) から付くかを確かめる (check108.ps1)
  check108.bat → r_check108.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check108.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== J と既往歴 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"

Q '--- 1. ★病名 → 判定 の対応表 T_KIOU_HANTEI (どの病名がどの項目を J/E にするか) ---' @"
SELECT h.RENBAN AS 連番, LTRIM(RTRIM(ISNULL(h.BYOMEI,''))) AS 病名, LTRIM(RTRIM(ISNULL(h.KOMOKU_CD,''))) AS 項目CD, m.MEISYO1 AS 項目名,
       LTRIM(RTRIM(ISNULL(h.SYOKEN_CD,''))) AS 所見CD, LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) AS 判定, h.K_FLG AS フラグ, h.KIJUN_PK
FROM T_KIOU_HANTEI h LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(h.KOMOKU_CD))
ORDER BY h.RENBAN
"@ 120

Q '--- 2. T_KIOU の列 ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型 FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_KIOU' ORDER BY c.ORDINAL_POSITION
"@ 30

Q '--- 3. ★見本: 最近 血圧が J の人 3 人の既往歴 (本人分) ---' @"
SELECT TOP 15 s.UKE_NO_KENSA AS 受付番号, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, k.RENBAN AS 連番,
       LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) AS 病名, LTRIM(RTRIM(ISNULL(k.BYOMEI_CD,''))) AS 病名CD,
       LTRIM(RTRIM(ISNULL(CAST(k.CHIRYO AS varchar(10)),''))) AS 治療, LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) AS 治療名, LTRIM(RTRIM(ISNULL(CAST(k.ZOKUGARA AS varchar(10)),''))) AS 続柄
FROM T_KENSIN s JOIN T_KIOU k ON k.KOJIN_ID = s.KOJIN_ID
WHERE s.PK_SEQ IN (SELECT TOP 3 s2.PK_SEQ FROM T_KENSIN s2 WHERE s2.F_TORIKESI = 0 AND s2.D_KENSIN >= '2026/09/01'
                   AND EXISTS (SELECT 1 FROM T_KENSA j WHERE j.PK_SEQ = s2.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD)) = '600150' AND LTRIM(RTRIM(ISNULL(j.HANTEI_KIGO,''))) = 'J')
                   ORDER BY s2.D_KENSIN DESC)
ORDER BY 1, 3
"@ 30

Q '--- 4. ★石屋 199 人の既往歴 (本人分) の中身: 病名 × 治療 × 人数 ---' @"
SELECT LTRIM(RTRIM(ISNULL(k.BYOMEI,''))) AS 病名, LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) AS 治療名, COUNT(DISTINCT s.PK_SEQ) AS 人数
FROM T_KENSIN s JOIN T_KIOU k ON k.KOJIN_ID = s.KOJIN_ID
WHERE $ISHIYA AND ISNULL(CAST(k.ZOKUGARA AS varchar(10)),'0') IN ('0','')
GROUP BY LTRIM(RTRIM(ISNULL(k.BYOMEI,''))), LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,'')))
ORDER BY 3 DESC
"@ 40

Q '--- 5. 治療 (CHIRYO) の値と名前の組み合わせ (2026 年の受診者全体) ---' @"
SELECT LTRIM(RTRIM(ISNULL(CAST(k.CHIRYO AS varchar(10)),''))) AS 治療, LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,''))) AS 治療名, COUNT(*) AS 件数
FROM T_KIOU k WHERE k.KOJIN_ID IN (SELECT s.KOJIN_ID FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01')
GROUP BY LTRIM(RTRIM(ISNULL(CAST(k.CHIRYO AS varchar(10)),''))), LTRIM(RTRIM(ISNULL(k.CHIRYOMEI,'')))
ORDER BY 3 DESC
"@ 20

W ''
W '=== 読み方 ==='
W '  1: 高血圧・糖尿病・脂質異常症 などの病名が 血圧・糖代謝・脂質代謝 の項目を J にする行があれば、J は既往歴から付いている。'
W '  3: J の人には 高血圧 (治療中) のような既往歴が入っているはず。'
W '  4: 石屋は「特記事項なし」か空ばかりなら、J が付かない理由はこれ。'
notepad $out
