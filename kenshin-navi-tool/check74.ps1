<#
  オプションを追加したとき、健診ナビが何の行を作るのかを調べる (check74.ps1)
  check74.bat をダブルクリックすると実行され、結果 r_check74.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  なぜ調べるか
    結果取込 (form_import.ps1 の Commit-Plan) は T_KENSA を UPDATE するだけで、
    行が無ければエラーになる。つまりオプションや胃部追加の結果を載せるには、
    先に T_KENSA にその検査の行が無いといけない。
    健診ナビでオプションを付けると何の行ができるのかを、9月の実績から確かめる。

  9月の福生3社の T_RYOUKIN 実績
    OPJ014 乳腺超音波検査 20件 / OPJ015 子宮細胞診 18件 / OPJ010 便潜血2日法 5件
    OPJ004 ABC健診 3件 / OPJ009 脳梗塞リスク 2件 / OPJ017,018,019,020 各1件
#>
[CmdletBinding()]
param([string]$Name = '福生', [string]$From = '2026/09/01', [string]$To = '2026/09/30')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check74.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== オプションを付けると何の行ができるか ($Name $From〜$To) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 9月の福生で OPJ014(乳腺超音波) を受けた人
$OPT = @"
(SELECT DISTINCT r.PK_SEQ, LTRIM(RTRIM(r.KOMOKU_CD)) AS OPT_CD
 FROM T_RYOUKIN r
 JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
 JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
 WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To'
   AND d.MEISYO1 LIKE N'%$Name%' AND LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,''))) <> '') o
"@

Q '--- 1. T_RYOUKIN の F_ADD_DEL の値 (追加か削除かの印) ---' @"
SELECT LTRIM(RTRIM(ISNULL(r.KOMOKU_CD,'(空)'))) AS 項目CD, r.F_ADD_DEL AS 追加削除,
       COUNT(*) AS 件数, MIN(r.REN) AS 連番最小, MAX(r.REN) AS 連番最大,
       MIN(LTRIM(RTRIM(ISNULL(r.TEKIYOU,'')))) AS 摘要
FROM T_RYOUKIN r
JOIN T_KENSIN s ON s.PK_SEQ = r.PK_SEQ
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND d.MEISYO1 LIKE N'%$Name%'
GROUP BY r.KOMOKU_CD, r.F_ADD_DEL ORDER BY 1, 2
"@ 60

Q '--- 2. ★オプションを受けた人の T_KENSA に、コースに無い検査が入っているか ---' @"
SELECT o.OPT_CD AS オプション, LTRIM(RTRIM(k.KOMOKU_CD)) AS 入っている項目CD,
       LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 検査項目, COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM $OPT
JOIN T_KENSIN s ON s.PK_SEQ = o.PK_SEQ
JOIN T_KENSA k ON k.PK_SEQ = o.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE NOT EXISTS (SELECT 1 FROM T_COURSE2 c2
                  WHERE c2.DANTAI_CD1 = s.DANTAI_CD1 AND c2.COURSE_CD = s.COURSE_CD
                    AND LTRIM(RTRIM(c2.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD)))
  AND LEFT(LTRIM(RTRIM(k.KOMOKU_CD)), 3) NOT IN ('084', '083', '085')
  AND LTRIM(RTRIM(k.KOMOKU_CD)) NOT LIKE '1001%'
GROUP BY o.OPT_CD, k.KOMOKU_CD, m.MEISYO1
ORDER BY o.OPT_CD, COUNT(DISTINCT k.PK_SEQ) DESC
"@ 200

Q '--- 3. オプションの項目CDそのもの (OPJ014 等) が T_KENSA にあるか ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(m.MEISYO1,''))) AS 検査項目,
       COUNT(*) AS 行数
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND d.MEISYO1 LIKE N'%$Name%'
  AND (LTRIM(RTRIM(k.KOMOKU_CD)) LIKE 'OP%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '0172%'
       OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '0777%' OR LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '0694%')
GROUP BY k.KOMOKU_CD, m.MEISYO1 ORDER BY 1
"@ 80

Q '--- 4. T_KENSA の列 (行を作るとき何を入れるか) ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ, c.is_nullable AS NULL可
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KENSA') ORDER BY c.column_id
"@ 60

Q '--- 5. T_KENSA の1行の実物 (何が入っているか。結果は伏せません。1人分だけ) ---' @"
SELECT TOP 8 k.PK_SEQ, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, k.*
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN BETWEEN '$From' AND '$To' AND d.MEISYO1 LIKE N'%$Name%'
ORDER BY k.PK_SEQ, k.KOMOKU_CD
"@ 20

Q '--- 6. ★オプションの親項目が何の子項目に展開されるか (T_KOMOKU の作り) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, LTRIM(RTRIM(ISNULL(MEISYO1,''))) AS 検査項目,
       LTRIM(RTRIM(ISNULL(K_KOMOKU,''))) AS 項目区分, LTRIM(RTRIM(ISNULL(K_SET,''))) AS セット,
       LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見CD, OPTION_GRP AS オプション群
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN
      ('OPJ004','OPJ009','OPJ010','OPJ011','OPJ014','OPJ015','OPJ017','OPJ018','OPJ019','OPJ020','OP0008')
ORDER BY KOMOKU_CD
"@ 40

Q '--- 7. セット項目の中身がどこかに定義されているか (SET らしいテーブルをさがす) ---' @"
SELECT o.name AS テーブル, COUNT(*) AS 列数
FROM sys.objects o JOIN sys.columns c ON c.object_id = o.object_id
WHERE o.type = 'U' AND (o.name LIKE '%SET%' OR o.name LIKE '%OPTION%')
GROUP BY o.name ORDER BY o.name
"@ 40

W ''
W '=== 読み方 ==='
W '  2 にオプションごとの検査項目が出れば、健診ナビはオプションを付けたときに'
W '    T_KENSA へその検査の行を作っている。→ 私のツールも同じ行を作ればよい。'
W '  2 が空なら、T_KENSA は増えていない。結果取込のときに行が無くて困るので、'
W '    別の方法 (コースを変える・健診ナビの画面で付ける) を考える。'
W '  3 で OPJ014 のような親項目そのものが T_KENSA にあるかが分かる。'
W '  6・7 でセットがどう展開されるかの手がかりをさがす。'
W '  ※ 読むだけです。何も書いていません。'
notepad $out
