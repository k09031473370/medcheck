<#
  石屋: 自動判定をかけた人に 判定コメント (医師指示事項 11001〜11009) が入っているかを見る (check99.ps1)
  check99.bat をダブルクリックすると実行され、結果 r_check99.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  医師指示事項 = 総合判定N(編集) 11001〜11009 の KEKKA (文章) と KEKKA_CD (定型文の番号)。自動判定が判定に応じて入れる。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check99.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 判定コメント (医師指示事項) の入り具合 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$SLOTS  = "('11001','11002','11003','11004','11005','11006','11007','11008','11009')"
$RANK = "CASE {0} WHEN 'A' THEN 1 WHEN 'B' THEN 2 WHEN 'C' THEN 3 WHEN 'D' THEN 4 WHEN 'E' THEN 5 WHEN 'F' THEN 6 WHEN 'G' THEN 7 WHEN 'J' THEN 7 ELSE 0 END"
$R_A1C = $RANK -f 'b.A1C'; $R_PLT = $RANK -f 'b.PLT'; $R_EGF = $RANK -f 'b.EGF'
$R_TOU = $RANK -f 'b.TOU'; $R_KET = $RANK -f 'b.KET'; $R_JIN = $RANK -f 'b.JIN'; $R_SG = $RANK -f 'b.SOGO'
$R_MAX = "(CASE WHEN $R_A1C >= $R_PLT AND $R_A1C >= $R_EGF THEN $R_A1C WHEN $R_PLT >= $R_EGF THEN $R_PLT ELSE $R_EGF END)"

Q '--- 1. 受診日ごと: 総合判定あり / コメントが 1 つ以上ある人 / コメント数の平均・最大 ---' @"
SELECT p.YMD AS 受診日, COUNT(*) AS 人数, SUM(CASE WHEN p.SOGO <> '' THEN 1 ELSE 0 END) AS 総合判定あり,
       SUM(CASE WHEN p.NCOM > 0 THEN 1 ELSE 0 END) AS コメントあり, SUM(CASE WHEN p.SOGO <> '' AND p.NCOM = 0 THEN 1 ELSE 0 END) AS 判定ありコメントなし,
       CAST(AVG(CAST(p.NCOM AS float)) AS decimal(4,1)) AS 平均コメント数, MAX(p.NCOM) AS 最大
FROM (SELECT s.PK_SEQ, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD,
             ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '11000'),'') AS SOGO,
             (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS NCOM
      FROM T_KENSIN s WHERE $ISHIYA) p
GROUP BY p.YMD ORDER BY 1
"@ 10

Q '--- 2. 使われている定型文 (番号と文章の先頭) と人数 ---' @"
SELECT LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS 番号, LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 60) AS 文章, COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
GROUP BY LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))), LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 60)
ORDER BY 3 DESC, 1
"@ 80

Q '--- 3. 総合判定はあるのにコメントが 1 つも無い人 ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, g.KANJI_SIMEI AS 氏名,
       (SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '11000') AS 総合判定
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE $ISHIYA
  AND ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '11000'),'') <> ''
  AND NOT EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '')
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@ 200

Q '--- 4. ★check98 で くくり/総合判定 を直す人 (36 人) に、今どんなコメントが入っているか (糖代謝・腎機能・血液のコメントが要るか判断用) ---' @"
SELECT b.UKE AS 受付番号, b.YMD AS 受診日, b.KANJI_SIMEI AS 氏名,
       CASE WHEN $R_A1C > $R_TOU THEN N'糖' ELSE '' END + CASE WHEN $R_PLT > $R_KET THEN N'血' ELSE '' END + CASE WHEN $R_EGF > $R_JIN THEN N'腎' ELSE '' END + CASE WHEN $R_MAX > $R_SG THEN N'総' ELSE '' END AS 直す欄,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 枠, LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS 番号, LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 50) AS 文章
FROM (SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, g.KANJI_SIMEI,
             LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) AS A1C, LTRIM(RTRIM(ISNULL(p.HANTEI_KIGO,''))) AS PLT, LTRIM(RTRIM(ISNULL(e.HANTEI_KIGO,''))) AS EGF,
             LTRIM(RTRIM(ISNULL(t.HANTEI_KIGO,''))) AS TOU, LTRIM(RTRIM(ISNULL(kt.HANTEI_KIGO,''))) AS KET, LTRIM(RTRIM(ISNULL(j.HANTEI_KIGO,''))) AS JIN,
             LTRIM(RTRIM(ISNULL(sg.HANTEI_KIGO,''))) AS SOGO
      FROM T_KENSIN s
      LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
      LEFT JOIN T_KENSA a  ON a.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(a.KOMOKU_CD))  = '068689'
      LEFT JOIN T_KENSA p  ON p.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(p.KOMOKU_CD))  = '069010'
      LEFT JOIN T_KENSA e  ON e.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(e.KOMOKU_CD))  = '068676'
      LEFT JOIN T_KENSA t  ON t.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(t.KOMOKU_CD))  = '600340'
      LEFT JOIN T_KENSA kt ON kt.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(kt.KOMOKU_CD)) = '600310'
      LEFT JOIN T_KENSA j  ON j.PK_SEQ  = s.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD))  = '600220'
      LEFT JOIN T_KENSA sg ON sg.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(sg.KOMOKU_CD)) = '11000'
      WHERE $ISHIYA AND LTRIM(RTRIM(ISNULL(sg.HANTEI_KIGO,''))) <> '') b
LEFT JOIN T_KENSA k ON k.PK_SEQ = b.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
WHERE $R_A1C > $R_TOU OR $R_PLT > $R_KET OR $R_EGF > $R_JIN OR $R_MAX > $R_SG
ORDER BY b.YMD, b.UKE, k.KOMOKU_CD
"@ 300

Q '--- 5. 参考: 糖代謝・腎機能・血液 に関係しそうな定型文が石屋で使われているか (文章に 糖 / HbA1c / 腎 / eGFR / 血小板 を含むもの) ---' @"
SELECT LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) AS 番号, LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 70) AS 文章, COUNT(DISTINCT k.PK_SEQ) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS
  AND (k.KEKKA LIKE N'%糖%' OR k.KEKKA LIKE N'%HbA1c%' OR k.KEKKA LIKE N'%腎%' OR k.KEKKA LIKE N'%eGFR%' OR k.KEKKA LIKE N'%血小板%' OR k.KEKKA LIKE N'%血液%')
GROUP BY LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))), LEFT(LTRIM(RTRIM(ISNULL(k.KEKKA,''))), 70)
ORDER BY 3 DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  1: 自動判定済み (総合判定あり) の人にコメントが入っているか。「判定ありコメントなし」が 0 なら全員に何かしら入っている。'
W '  2: どの定型文が使われているか。4: くくりを直す 36 人の今のコメント。糖代謝 D なのに糖のコメントが無い、などが分かる。'
W '  5: 糖・腎・血液のコメントが自動判定でどう付いているか。これが分かれば、くくりを上げた人に同じ番号の文章を足せる (医師指示事項追加 と同じ方法)。'
notepad $out
