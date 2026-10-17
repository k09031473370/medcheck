<#
  石屋: 血糖・中性脂肪の振り分け先の枠と値 (check104.ps1)
  結果票は 空腹時血糖 / 随時血糖 / 空腹時中性脂肪 / 随時中性脂肪 を印字し、未判別は印字しない (2026/10/18 に 9/3 随時の人と 9/9 の人のプレビューで確認)。
  振り分け先の枠が 9/9・9/10 の人にあるか (ツールで値を写せるか)、値と判定がどこに入っているかを見る。
  check104.bat をダブルクリック → r_check104.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check104.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 血糖・中性脂肪の振り分け先 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$CDS = "('001282','068682','H30-0011','H30-0012','H30-0013','R06-0002','068702','R06-0003')"

Q '--- 1. ★受診日 × 項目: 枠 / 値あり / 判定あり / H・L 印あり ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, MAX(LTRIM(RTRIM(ISNULL(m.RYAKUSYO1,'')))) AS 略称,
       COUNT(*) AS 枠,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HL,''))) <> '' THEN 1 ELSE 0 END) AS 印あり
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $CDS
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1, 2
"@ 40

Q '--- 2. 見本: 9/3 で振り分け済みの人 (4631 空腹 / 4634 随時 / 4632 食直後) と 未振り分けの人 (4671 / 4691) の値・判定・印 ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       MAX(LTRIM(RTRIM(ISNULL(k.KEKKA,'')))) AS 値, MAX(LTRIM(RTRIM(ISNULL(k.KEKKA_CD,'')))) AS 結果CD,
       MAX(LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))) AS 判定, MAX(LTRIM(RTRIM(ISNULL(k.HL,'')))) AS 印
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND s.UKE_NO_KENSA IN (4631, 4634, 4632, 4671, 4691) AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $CDS
GROUP BY s.UKE_NO_KENSA, LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1, 2
"@ 60

Q '--- 3. 振り分け先の枠が無い人の数 (受診日別。ツールで写すなら枠が要る) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN NOT EXISTS (SELECT 1 FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = 'H30-0011') THEN 1 ELSE 0 END) AS 空腹時血糖の枠なし,
       SUM(CASE WHEN NOT EXISTS (SELECT 1 FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = '068702') THEN 1 ELSE 0 END) AS 空腹時中性脂肪の枠なし,
       SUM(CASE WHEN NOT EXISTS (SELECT 1 FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = 'H30-0012') THEN 1 ELSE 0 END) AS 随時血糖の枠なし,
       SUM(CASE WHEN NOT EXISTS (SELECT 1 FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = 'R06-0003') THEN 1 ELSE 0 END) AS 随時中性脂肪の枠なし
FROM T_KENSIN s WHERE $ISHIYA
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111) ORDER BY 1
"@ 10

W ''
W '=== 読み方 ==='
W '  1: 9/3 は 空腹時/随時/食直後 に値と判定が入っているはず。9/9・9/10 は 未判別にしか値が無いはず。'
W '  2: 振り分け済みの人で、未判別の値が残っているか (写す) / 消えているか (移す)、印がどこに付くかを見る。'
W '  3: 枠なしが 0 なら、ツールで 未判別 → 空腹時 に値・判定・印を写せる。枠なしが多ければ、自動判定のかけ直しが必要。'
notepad $out
