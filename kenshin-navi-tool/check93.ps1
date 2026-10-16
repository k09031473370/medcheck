<#
  健診ナビの「検査追加」が、どの表に行を作っているかを調べる (check93.ps1)
  check93.bat をダブルクリックすると実行され、結果 r_check93.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  ツールで枠 (T_KENSA の行) を足せるかを判断するため、
    1) PK_SEQ と KOMOKU_CD の両方の列を持つ表 (= 人×項目 の行を持つ表) を全部挙げる
    2) 昨日 eGFR (068676) の枠を足した人 1 人 (4831) について、その表それぞれに 068676 の行が何件あるか
    3) 同じ人の T_KENSA の 068676 の行と、元からあった項目 (068674 ｸﾚｱﾁﾆﾝ) の行を全列並べて、どの列が違うか
  を見る。T_KENSA 以外にも行ができていれば、T_KENSA だけ足すのは危ない (= 健診ナビの画面で足す)。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check93.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 検査追加が触る表の調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. PK_SEQ と KOMOKU_CD の両方を持つ表 ---' @"
SELECT a.TABLE_NAME AS 表,
       (SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = a.TABLE_NAME) AS 列数
FROM INFORMATION_SCHEMA.COLUMNS a
JOIN INFORMATION_SCHEMA.COLUMNS b ON b.TABLE_NAME = a.TABLE_NAME AND b.COLUMN_NAME = 'KOMOKU_CD'
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = a.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE a.COLUMN_NAME = 'PK_SEQ'
ORDER BY 1
"@ 40

Q '--- 2. 4831 (9/10 鳥居) の PK_SEQ ---' @"
SELECT s.PK_SEQ, s.UKE_NO_KENSA, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日
FROM T_KENSIN s WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '2026/09/10' AND s.UKE_NO_KENSA = 4831
"@ 5

# 1 の表ごとに、4831 の 068676 (昨日足した枠) と 068674 (元からの枠) の行数を数える SQL を組み立てる
$tbls = & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql @"
SELECT a.TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS a
JOIN INFORMATION_SCHEMA.COLUMNS b ON b.TABLE_NAME = a.TABLE_NAME AND b.COLUMN_NAME = 'KOMOKU_CD'
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = a.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE a.COLUMN_NAME = 'PK_SEQ' ORDER BY 1
"@ -MaxRows 100 2>$null | Where-Object { $_ -match '^\s*(T_|M_)[A-Z0-9_]+\s*$' } | ForEach-Object { $_.Trim() }
$parts = @()
foreach ($t in $tbls) {
    $parts += "SELECT '$t' AS 表, SUM(CASE WHEN LTRIM(RTRIM(KOMOKU_CD)) = '068676' THEN 1 ELSE 0 END) AS eGFRの行, SUM(CASE WHEN LTRIM(RTRIM(KOMOKU_CD)) = '068674' THEN 1 ELSE 0 END) AS ｸﾚｱﾁﾆﾝの行, COUNT(*) AS この人の全行 FROM $t WHERE PK_SEQ = (SELECT TOP 1 PK_SEQ FROM T_KENSIN WHERE F_TORIKESI = 0 AND D_KENSIN = '2026/09/10' AND UKE_NO_KENSA = 4831)"
}
if ($parts.Count -gt 0) {
    Q '--- 3. ★4831 について、表ごとの 068676 (昨日足した eGFR) と 068674 (元からあるｸﾚｱﾁﾆﾝ) の行数 ---' ($parts -join "`nUNION ALL`n") 40
} else {
    W ''; W '--- 3. 表の一覧が取れなかったので省略 ---'
}

Q '--- 4. 4831 の T_KENSA: 068676 と 068674 の行を全列 (違う列を見る) ---' @"
SELECT * FROM T_KENSA
WHERE PK_SEQ = (SELECT TOP 1 PK_SEQ FROM T_KENSIN WHERE F_TORIKESI = 0 AND D_KENSIN = '2026/09/10' AND UKE_NO_KENSA = 4831)
  AND LTRIM(RTRIM(KOMOKU_CD)) IN ('068676','068674','068689','069010')
ORDER BY KOMOKU_CD
"@ 10

Q '--- 5. T_KENSA の列一覧 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, IS_NULLABLE AS NULL可, COLUMN_DEFAULT AS 既定値
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_KENSA' ORDER BY ORDINAL_POSITION
"@ 80

W ''
W '=== 読み方 ==='
W '  3 で eGFR の行が T_KENSA 以外の表にもあれば、検査追加は複数の表を更新している → ツールで T_KENSA だけ足すのは危ない。'
W '     T_KENSA にしか無ければ、同じコースの人の行をコピーして枠を作るツールが作れる (4・5 で列の形を見て決める)。'
notepad $out
