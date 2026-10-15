<#
  心電図・胃部X線・胸部X線の「判定」の入り方を調べる (check88.ps1)
  check88.bat をダブルクリックすると実行され、結果 r_check88.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  健診ナビの自動判定で心電図・胃部の判定が付かない、との報告。
    1) 健診ナビが実際に使っている判定記号 (A/B/C/D/E/F/G/J、C12 や D2 のような装置コードが入っているか)
    2) 所見マスタ ZK011(心電図)・ZK031(胃部)・ZK021(胸部) に判定記号が入っているか
    3) 石屋の心電図・胃部・胸部の行に、所見と判定がどれだけ入っているか
    4) 判定区分のマスタ表があるか
  を見て、ツールで読影医の判定を書くときの文字を決める。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check88.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 判定記号の入り方 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★2026 年の結果で使われている判定記号 (項目種別ごと: 心電図 067112A / 胃部 077300A / 胸部 077010A / 診察 017101A / それ以外) ---' @"
SELECT CASE LTRIM(RTRIM(k.KOMOKU_CD)) WHEN '067112A' THEN N'心電図' WHEN '077300A' THEN N'胃部X線' WHEN '077010A' THEN N'胸部X線' WHEN '017101A' THEN N'診察' ELSE N'その他' END AS 種別,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定記号, COUNT(*) AS 件数,
       MIN(CONVERT(varchar(10), s.D_KENSIN, 111)) AS 初出, MAX(CONVERT(varchar(10), s.D_KENSIN, 111)) AS 最新
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN >= '2026/01/01' AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> ''
GROUP BY CASE LTRIM(RTRIM(k.KOMOKU_CD)) WHEN '067112A' THEN N'心電図' WHEN '077300A' THEN N'胃部X線' WHEN '077010A' THEN N'胸部X線' WHEN '017101A' THEN N'診察' ELSE N'その他' END,
         LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,'')))
ORDER BY 1, 2
"@ 80

Q '--- 2. ★所見マスタの判定記号の有無 (ZK011 心電図 / ZK031 胃部 / ZK021 胸部 / SHIN 診察) ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS 所見区分, COUNT(*) AS 所見の数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定記号あり,
       MIN(CASE WHEN LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) <> '' THEN LTRIM(RTRIM(KEKKA_CD)) + '=' + LTRIM(RTRIM(HANTEI_KIGO)) END) AS 例
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) IN ('ZK011','ZK031','ZK021','ZK030','SHIN')
GROUP BY LTRIM(RTRIM(SYOKEN_CD)) ORDER BY 1
"@ 10

Q '--- 3. 石屋で使われている心電図・胃部の所見コードと、マスタの判定記号 ---' @"
SELECT CASE LTRIM(RTRIM(k.KOMOKU_CD)) WHEN '067112A' THEN N'心電図' ELSE N'胃部X線' END AS 種別,
       LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS 結果CD, LTRIM(RTRIM(ISNULL(k.KEKKA,''))) AS 所見文,
       COUNT(*) AS 人数, SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり,
       MIN('[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']') AS 判定の例,
       MIN('[' + LTRIM(RTRIM(ISNULL(m.HANTEI_KIGO,''))) + ']') AS マスタの判定
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_SYOKEN2 m ON LTRIM(RTRIM(m.SYOKEN_CD)) = CASE LTRIM(RTRIM(k.KOMOKU_CD)) WHEN '067112A' THEN 'ZK011' ELSE 'ZK031' END
                      AND LTRIM(RTRIM(m.KEKKA_CD)) = LTRIM(RTRIM(k.KEKKA_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('067112A','077300A') AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY k.KOMOKU_CD, k.KEKKA_CD, k.KEKKA ORDER BY 1, 4 DESC
"@ 60

Q '--- 4. 石屋 受診日ごと: 心電図・胃部・胸部の 所見あり / 判定あり の人数 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 心電図所見,
       SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A' AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 心電図判定,
       SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 胃部所見,
       SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '077300A' AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 胃部判定,
       SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A' AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 胸部所見,
       SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A' AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 胸部判定
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
GROUP BY s.D_KENSIN ORDER BY 1
"@ 10

Q '--- 5. 石屋の胸部X線の判定記号の内訳 (いま入っている文字) ---' @"
SELECT '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定記号, COUNT(*) AS 人数
FROM T_KENSA k JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831
  AND LTRIM(RTRIM(k.KOMOKU_CD)) = '077010A'
GROUP BY LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) ORDER BY 1
"@ 10

Q '--- 6. 判定区分のマスタらしい表 (名前に HANTEI を含む表と列) ---' @"
SELECT t.TABLE_NAME AS 表, STUFF((SELECT ', ' + c.COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = t.TABLE_NAME ORDER BY c.ORDINAL_POSITION FOR XML PATH('')), 1, 2, '') AS 列
FROM INFORMATION_SCHEMA.TABLES t
WHERE t.TABLE_TYPE = 'BASE TABLE' AND (t.TABLE_NAME LIKE '%HANTEI%' OR t.TABLE_NAME LIKE '%HANTE%' OR t.TABLE_NAME LIKE '%KUBUN%')
ORDER BY 1
"@ 20

W ''
W '=== 読み方 ==='
W '  1: 健診ナビが帳票などで実際に使っている判定記号。A/B/C/D/E/F/G/J 以外 (C12, D2 など) が出ていれば、それも通る文字。'
W '  2・3: 所見マスタに判定記号が無い種別は、自動判定では判定が付かない。読影医/装置の判定をツールで書く必要がある。'
W '  4: 石屋で所見はあるのに判定が無い人数。5: 胸部X線にいま入っている判定の文字 (C/D は健診ナビの 8 段階では 要注意/要観察 なので直す)。'
notepad $out
