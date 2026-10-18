<#
  石屋: 協会判定 (610xxx) が くくり判定 (600xxx) と合っているか (check105.ps1)
  協会けんぽへのデータは協会判定 (指導区分) を使う。自動判定は協会判定を計算し直すが、ツールで直したくくり判定 (HbA1c・eGFR・眼底) を
  拾うかは分からないので、自動判定のあとにこれで比べる。check105.bat → r_check105.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check105.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 協会判定とくくり判定 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"

Q '--- 1. 協会判定 (610xxx) とくくり判定 (600xxx) の項目と、石屋の枠の数 ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))) AS 石屋の枠
FROM T_KOMOKU m WHERE LTRIM(RTRIM(m.KOMOKU_CD)) LIKE '61%' OR LTRIM(RTRIM(m.KOMOKU_CD)) LIKE '60%'
GROUP BY LTRIM(RTRIM(m.KOMOKU_CD)) ORDER BY 1
"@ 80

Q '--- 2. 見本 5 人の 協会判定・くくり判定 (結果と判定) ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 結果, MAX(LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))) AS 判定
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND s.UKE_NO_KENSA IN (4633, 4648, 4667, 4691, 4714)
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '61%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '60%')
  AND (LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' OR LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '')
GROUP BY s.UKE_NO_KENSA, LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1, 2
"@ 200

Q '--- 3. ★眼底: 【判定】眼底 600060 × 【協会判定】眼底指導区分 610100 の人数 (受診日別) ---' @"
SELECT t.YMD AS 受診日, t.GAN AS 眼底, t.KYO AS 協会眼底指導区分, COUNT(*) AS 人数 FROM (
  SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD,
         MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '600060' THEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) END) AS GAN,
         MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '610100' THEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + '/' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) END) AS KYO
  FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
  WHERE $ISHIYA GROUP BY s.PK_SEQ, s.D_KENSIN) t
GROUP BY t.YMD, t.GAN, t.KYO ORDER BY 1, 2, 3
"@ 60

Q '--- 4. 協会判定ごとの 結果/判定 の分布 (受診日別。指導区分の文字の種類を知るため) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + '/' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) AS 結果_判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '61%'
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), LTRIM(RTRIM(k.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + '/' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))
ORDER BY 1, 2, 4
"@ 200

W ''
W '=== 読み方 ==='
W '  1: 協会判定の項目と、どのくくりに対応するか (名前で見る)。'
W '  2: 見本の人で、協会判定がくくり判定 (ツールで直した HbA1c・eGFR・眼底) を拾っているか。'
W '  3: 眼底 F の人の 協会眼底指導区分 が「異常なし」側なら、自動判定は眼底の業者判定を拾っていない。'
notepad $out
