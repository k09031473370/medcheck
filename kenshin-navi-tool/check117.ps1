<#
  福生商工会: オプションを申し込んだ人の人数と項目 (check117.ps1)
  オプションは 料金明細 T_RYOUKIN に OP で始まる項目CD (OPJ014 乳腺超音波 など) で入る (check74・75 で確認)。
  あわせて、検査の枠 T_KENSA にオプションの行 (OP〜) ができているかも数える。
  check117.bat → r_check117.txt。DBは読むだけ。
  1〜3 は人数だけ (氏名なし)。4 は 1 人ずつの一覧で氏名が入るので、院内だけで使う (外に貼らない)。
#>
[CmdletBinding()]
param([string]$Name = '福生', [string]$From = '2026/10/01', [string]$To = '2026/10/31')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check117.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== $Name のオプション申込 ($From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$WHO = @"
s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'
  AND s.DANTAI_CD1 IN (SELECT d.DANTAI_CD1 FROM T_DANTAI1 d WHERE d.MEISYO1 LIKE N'%$Name%')
"@
$OPT = "LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) LIKE 'OP%'"

Q '--- 0. 対象の団体と、受診日ごとの人数 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(s.DANTAI_CD1)) AS 団体CD,
       MAX(LTRIM(RTRIM(ISNULL(d.MEISYO1,'')))) AS 団体名, COUNT(*) AS 受診者
FROM T_KENSIN s LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE $WHO
GROUP BY s.D_KENSIN, s.DANTAI_CD1 ORDER BY 1, 2
"@ 40

Q '--- 1. ★オプションの項目ごとの人数 (料金明細から。追加削除 の印ごと) ---' @"
SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS オプション,
       ISNULL(CONVERT(varchar(10), r.F_ADD_DEL),'') AS 追加削除, COUNT(DISTINCT r.PK_SEQ) AS 人数, COUNT(*) AS 行数
FROM T_RYOUKIN r
JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
WHERE $WHO AND $OPT
GROUP BY LTRIM(RTRIM(r.KOMOKU_CD)), r.F_ADD_DEL
ORDER BY 4 DESC, 1
"@ 80

Q '--- 2. ★オプションを 1 つ以上申し込んだ人の数 (受診日ごと) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 受診者,
       SUM(CASE WHEN EXISTS (SELECT 1 FROM T_RYOUKIN r WHERE r.PK_SEQ = s.PK_SEQ AND $OPT) THEN 1 ELSE 0 END) AS オプションあり,
       SUM(CASE WHEN (SELECT COUNT(DISTINCT LTRIM(RTRIM(r.KOMOKU_CD))) FROM T_RYOUKIN r WHERE r.PK_SEQ = s.PK_SEQ AND $OPT) >= 2 THEN 1 ELSE 0 END) AS うち2つ以上
FROM T_KENSIN s
WHERE $WHO
GROUP BY s.D_KENSIN ORDER BY 1
"@ 20

Q '--- 3. 確かめ: 検査の枠 (T_KENSA) にあるオプションの行の人数 (1 と合っていれば、枠もできている) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(LTRIM(RTRIM(ISNULL(m.MEISYO1,'')))) AS オプション, COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $WHO AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE 'OP%'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)) ORDER BY 3 DESC, 1
"@ 80

Q '--- 4. 1 人ずつ (受診日・受付番号・氏名・オプション)。★氏名入り。院内だけで使う ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       STUFF((SELECT DISTINCT N'、' + LTRIM(RTRIM(ISNULL(m.MEISYO1, r.KOMOKU_CD)))
              FROM T_RYOUKIN r LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(r.KOMOKU_CD))
              WHERE r.PK_SEQ = s.PK_SEQ AND $OPT FOR XML PATH('')), 1, 1, '') AS オプション
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE $WHO AND EXISTS (SELECT 1 FROM T_RYOUKIN r WHERE r.PK_SEQ = s.PK_SEQ AND $OPT)
ORDER BY 1, LEN(LTRIM(RTRIM(s.UKE_NO_KENSA))), s.UKE_NO_KENSA
"@ 500

W ''
W '=== 読み方 ==='
W '  1 が「何を何人が申し込んだか」。追加削除 の印が 2 種類あれば、取り消した分が混ざっている (どちらが取消かは 0 の団体ページで確かめる)。'
W '  2 が「オプションを申し込んだ人の数」。'
W '  3 の人数が 1 と違えば、料金明細と検査の枠がずれている (枠が無いと結果が入れられない)。'
W '  4 は氏名入りなので、外に出さない。'
notepad $out
