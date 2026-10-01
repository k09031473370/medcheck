<#
  セットの中身の定義を読む (check75.ps1)
  check75.bat をダブルクリックすると実行され、結果 r_check75.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check74 で分かったこと
    オプションを付けると T_KENSA に2種類の行ができる
      ・オプションの親項目そのもの      例 OPJ014【乳腺超音波検査】
      ・その中身の検査項目              例 017210【乳腺超音波】+ 017210A/B/C
    件数は T_RYOUKIN とぴったり一致 (OPJ014 20件 ↔ 017210 20行)
    T_KOMOKU で オプション項目はすべて K_SET = 1 (セット項目)
    T_SET1 / T_SET2 / M_OPTION_GRP というテーブルがある

  ここで読みたいもの
    どのセットにどの検査が含まれるかの定義。
    これが読めれば、オプションを付けたとき何の行を作ればよいかが確定する。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check75.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
$OPTS = "'OPJ004','OPJ009','OPJ010','OPJ011','OPJ014','OPJ015','OPJ017','OPJ018','OPJ019','OPJ020','OP0008'"
"=== セットの中身の定義 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_SET1 / T_SET2 / M_OPTION_GRP の列と件数 ---' @"
SELECT o.name AS テーブル, c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c
JOIN sys.objects o ON o.object_id = c.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE o.name IN ('T_SET1','T_SET2','M_OPTION_GRP')
ORDER BY o.name, c.column_id
"@ 40

Q '--- 2. T_SET1 の中身 (全部) ---' @"
SELECT TOP 100 * FROM T_SET1
"@ 120

Q '--- 3. ★T_SET2 の中身 (オプションのセットだけ) ---' @"
SELECT TOP 200 * FROM T_SET2
"@ 220

Q '--- 4. M_OPTION_GRP の中身 ---' @"
SELECT TOP 40 * FROM M_OPTION_GRP
"@ 50

Q '--- 5. T_KOMOKU で セット=1 の項目 (オプションの親) 全部 ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分, OPTION_GRP AS オプション群,
       LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見CD
FROM T_KOMOKU WHERE LTRIM(RTRIM(ISNULL(K_SET,''))) = '1'
ORDER BY KOMOKU_CD
"@ 200

Q '--- 6. ★9月の福生で、オプションごとに実際にできた検査行 (1人だけを例に) ---' @"
SELECT LTRIM(RTRIM(r.KOMOKU_CD)) AS オプション, r.PK_SEQ,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,''))) AS 親項目CD
FROM T_RYOUKIN r
JOIN T_KENSA k ON k.PK_SEQ = r.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE r.PK_SEQ IN (SELECT TOP 1 PK_SEQ FROM T_RYOUKIN WHERE LTRIM(RTRIM(KOMOKU_CD)) = 'OPJ014')
  AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) <> ''
  AND (LTRIM(RTRIM(k.OYA_KOMOKU_CD)) <> '' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE 'OP%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '0172%')
ORDER BY 1, 3
"@ 80

Q '--- 7. ★OYA_KOMOKU_CD の使われ方 (子項目が親をどう指しているか) ---' @"
SELECT LTRIM(RTRIM(k.OYA_KOMOKU_CD)) AS 親項目CD, LTRIM(RTRIM(ISNULL(p.MEISYO1,''))) AS 親の名前,
       COUNT(DISTINCT LTRIM(RTRIM(k.KOMOKU_CD))) AS 子の種類, COUNT(*) AS 行数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_KOMOKU p ON LTRIM(RTRIM(p.KOMOKU_CD)) = LTRIM(RTRIM(k.OYA_KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/09/01' AND '2026/09/30'
  AND d.MEISYO1 LIKE N'%福生%' AND LTRIM(RTRIM(ISNULL(k.OYA_KOMOKU_CD,''))) <> ''
GROUP BY k.OYA_KOMOKU_CD, p.MEISYO1 ORDER BY COUNT(*) DESC
"@ 80

Q '--- 8. SEQ1 / SEQ2 の決まり方 (行を作るとき何を入れるか) ---' @"
SELECT TOP 20 k.PK_SEQ, LTRIM(RTRIM(k.SEQ1)) AS SEQ1, LTRIM(RTRIM(k.SEQ2)) AS SEQ2,
       COUNT(*) AS 行数, MIN(LTRIM(RTRIM(k.KOMOKU_CD))) AS 項目CD例
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '2026/09/01' AND '2026/09/30' AND d.MEISYO1 LIKE N'%福生%'
GROUP BY k.PK_SEQ, k.SEQ1, k.SEQ2 ORDER BY k.PK_SEQ
"@ 30

W ''
W '=== 読み方 ==='
W '  2・3 にセットと項目の対応が入っていれば、それが答え。'
W '  オプションを付けたとき作るべき検査行が確定する。'
W '  6 は1人を例にして、OPJ014 を受けた人に実際どんな行があるかを見る。'
W '  7 の OYA_KOMOKU_CD で、子項目が親をどう指しているかが分かる。'
W '  8 は行を新しく作るときに SEQ1/SEQ2 に何を入れるかの手がかり。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
