<#
  石屋: 医師指示事項に【中性脂肪】の文と【脂質代謝】(コレステロール) の文が両方ある人 (check116.ps1)
  中性脂肪の文は kukuri_comment が足したもの。脂質代謝の文は健診ナビの自動判定が入れたもの。
  脂質代謝の文が LDL・HDL の話なら別の内容 (両方あってよい)。脂質代謝の文にも中性脂肪が書いてあれば重複。
  check116.bat → r_check116.txt。DBは読むだけ。氏名は出さない。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check116.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 中性脂肪の文と脂質代謝の文が両方ある人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$SLOTS = "'11001','11002','11003','11004','11005','11006','11007','11008','11009','11010','11011','11012','11013','11014','11015','11016','11017','11018','11019','11020'"
$TG    = "c.KEKKA LIKE N'%中性脂肪%'"
$CHO   = "(c.KEKKA LIKE N'%脂質%' OR c.KEKKA LIKE N'%コレステロール%' OR c.KEKKA LIKE N'%ｺﾚｽﾃﾛｰﾙ%' OR c.KEKKA LIKE N'%ＬＤＬ%' OR c.KEKKA LIKE N'%ＨＤＬ%' OR c.KEKKA LIKE N'%LDL%' OR c.KEKKA LIKE N'%HDL%')"
$BOTH  = @"
EXISTS (SELECT 1 FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN ($SLOTS) AND $TG)
  AND EXISTS (SELECT 1 FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN ($SLOTS) AND $CHO AND NOT ($TG))
"@
function J($cd) { "(SELECT MAX(LTRIM(RTRIM(ISNULL(x.HANTEI_KIGO,'')))) FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = '$cd')" }

Q '--- 1. 人数 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, COUNT(*) AS 人数
FROM T_KENSIN s WHERE $ISHIYA AND $BOTH
GROUP BY s.D_KENSIN ORDER BY 1
"@ 10

Q '--- 2. 判定 (脂質代謝のくくり / 中性脂肪 / LDL / HDL / 総コレステロール) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号,
       $(J '600320') AS 脂質代謝,
       COALESCE(NULLIF($(J '068702'),''), NULLIF($(J 'R06-0003'),''), $(J 'R06-0002')) AS 中性脂肪,
       $(J '068709') AS LDL, $(J '068708') AS HDL, $(J '068705') AS 総ｺﾚｽﾃﾛｰﾙ
FROM T_KENSIN s WHERE $ISHIYA AND $BOTH
ORDER BY 1, 2
"@ 300

Q '--- 3. 医師指示事項の脂質の文 (枠番号・文) ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(c.KOMOKU_CD)) AS 枠,
       CASE WHEN $TG AND $CHO THEN N'★両方の言葉' WHEN $TG THEN N'中性脂肪' ELSE N'脂質代謝' END AS 種類,
       LEFT(LTRIM(RTRIM(c.KEKKA)), 80) AS 文
FROM T_KENSIN s JOIN T_KENSA c ON c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN ($SLOTS)
WHERE $ISHIYA AND $BOTH AND ($TG OR $CHO)
ORDER BY 1, 2
"@ 600

Q '--- 4. 脂質代謝の文の言い回しごとの人数 ---' @"
SELECT LEFT(LTRIM(RTRIM(c.KEKKA)), 80) AS 文, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA c ON c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN ($SLOTS)
WHERE $ISHIYA AND $BOTH AND $CHO AND NOT ($TG)
GROUP BY LEFT(LTRIM(RTRIM(c.KEKKA)), 80) ORDER BY 2 DESC
"@ 100

W ''
W '=== 読み方 ==='
W '  4 の文が LDL・HDL・コレステロールの話だけなら、中性脂肪の文とは別の内容 (両方あってよい)。'
W '  4 の文が「脂質代謝 高値です」のように何の値か書いていないときは、2 の判定で見る:'
W '    中性脂肪だけ C 以上で LDL・HDL が A/B → 同じことを 2 回書いている (どちらかを消す)'
W '    LDL か HDL も C 以上 → 別の内容 (両方あってよい)'
W '  3 で「★両方の言葉」の文がある人は、その文 1 つで中性脂肪も書いてあるので、【中性脂肪】の文は重複。'
notepad $out
