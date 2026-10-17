<#
  石屋: 眼底の所見と判定がどこまで入っているか (check102.ps1)
  check102.bat をダブルクリックすると実行され、結果 r_check102.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  眼底の項目: 所見右 067052 / 所見左 067053 (GANTEI コード or 文字) / KW 067054,067055 / ScheieH 067056,067057 / ScheieS 067058,067059
             / WM H30-0051,0052 / Davis H30-0053,0054 / セット【眼底検査(左右別)】067050
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check102.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 眼底の所見と判定 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$GAN = "('067050','067052','067053','067054','067055','067056','067057','067058','067059','H30-0051','H30-0052','H30-0053','H30-0054')"

Q '--- 1. ★受診日 × 眼底の項目ごと: 枠 / 値あり / 判定あり ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 枠, SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS 値あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) <> '' THEN 1 ELSE 0 END) AS 判定あり,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = '' THEN 1 ELSE 0 END) AS 値ありで判定なし
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $GAN
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1, 2
"@ 60

Q '--- 2. 眼底所見 右 (067052) の値と判定の組み合わせ (何が入っていて、どう判定されているか) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LEFT(ISNULL(k.KEKKA,''), 30) AS 所見, ISNULL(k.KEKKA_CD,'') AS 結果CD, ISNULL(k.HANTEI_KIGO,'') AS 判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067052'
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), LEFT(ISNULL(k.KEKKA,''), 30), ISNULL(k.KEKKA_CD,''), ISNULL(k.HANTEI_KIGO,'')
ORDER BY 1, 5 DESC
"@ 80

Q '--- 3. 眼底所見 左 (067053) の値と判定の組み合わせ ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, LEFT(ISNULL(k.KEKKA,''), 30) AS 所見, ISNULL(k.KEKKA_CD,'') AS 結果CD, ISNULL(k.HANTEI_KIGO,'') AS 判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067053'
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), LEFT(ISNULL(k.KEKKA,''), 30), ISNULL(k.KEKKA_CD,''), ISNULL(k.HANTEI_KIGO,'')
ORDER BY 1, 5 DESC
"@ 80

Q '--- 4. ★眼底の所見 (右か左) に値があるのに、右も左も判定が空の人 (先生に判定してもらう人) ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, g.KANJI_SIMEI AS 氏名,
       LEFT(ISNULL(r.KEKKA,''), 30) AS 右所見, ISNULL(r.HANTEI_KIGO,'') AS 右判定, LEFT(ISNULL(l.KEKKA,''), 30) AS 左所見, ISNULL(l.HANTEI_KIGO,'') AS 左判定,
       ISNULL((SELECT TOP 1 ISNULL(x.HANTEI_KIGO,'') FROM T_KENSA x WHERE x.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(x.KOMOKU_CD)) = '11000'),'') AS 総合判定
FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KENSA r ON r.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(r.KOMOKU_CD)) = '067052'
LEFT JOIN T_KENSA l ON l.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(l.KOMOKU_CD)) = '067053'
WHERE $ISHIYA
  AND (LTRIM(RTRIM(ISNULL(r.KEKKA,''))) <> '' OR LTRIM(RTRIM(ISNULL(l.KEKKA,''))) <> '')
  AND LTRIM(RTRIM(ISNULL(r.HANTEI_KIGO,''))) = '' AND LTRIM(RTRIM(ISNULL(l.HANTEI_KIGO,''))) = ''
ORDER BY s.D_KENSIN, s.UKE_NO_KENSA
"@ 250

Q '--- 5. 眼底の所見マスタ GANTEI に判定記号があるか (自動判定で付くかどうかの元) ---' @"
SELECT y.KEKKA_CD AS 結果CD, y.SYOKEN AS 所見, ISNULL(y.HANTEI_KIGO,'') AS 判定
FROM T_SYOKEN2 y WHERE LTRIM(RTRIM(y.SYOKEN_CD)) = 'GANTEI' ORDER BY y.KEKKA_CD
"@ 80

Q '--- 6. ★【判定】眼底 600060 (結果票の「判定_眼底」に出る欄): 受診日 × 判定 × 総合判定 の人数 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, ISNULL(k.HANTEI_KIGO,'') AS 眼底判定, ISNULL(z.HANTEI_KIGO,'') AS 総合判定, COUNT(*) AS 人数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '600060'
LEFT JOIN T_KENSA z ON z.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(z.KOMOKU_CD)) = '11000'
WHERE $ISHIYA
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), ISNULL(k.HANTEI_KIGO,''), ISNULL(z.HANTEI_KIGO,'')
ORDER BY 1, 2, 3
"@ 60

Q '--- 7. 眼底に関係する項目のマスタ (600060 以外にも「眼底」の名前の判定欄があるか) と 石屋の枠の数 ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       (SELECT COUNT(*) FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
         WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))) AS 石屋の枠
FROM T_KOMOKU m WHERE m.MEISYO1 LIKE '%眼底%' ORDER BY 1
"@ 60

Q '--- 8. 医師指示事項 (11001〜11009) に入っている「眼」の文章 (石屋 3 日分 + 2026 年の全員): 眼底のコメントの定型文を知るため ---' @"
SELECT LTRIM(RTRIM(k.KEKKA)) AS 文章, ISNULL(k.KEKKA_CD,'') AS 番号,
       SUM(CASE WHEN s.UKE_NO_KENSA BETWEEN 4631 AND 4831 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') THEN 1 ELSE 0 END) AS 石屋,
       COUNT(*) AS 全体
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('11001','11002','11003','11004','11005','11006','11007','11008','11009')
  AND k.KEKKA LIKE '%眼%'
GROUP BY LTRIM(RTRIM(k.KEKKA)), ISNULL(k.KEKKA_CD,'')
ORDER BY 4 DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  1: 受診日ごとに眼底の各項目に判定が付いているか。「値ありで判定なし」が 0 でなければ、その項目は判定が付いていない。'
W '  4: 先生に判定してもらう人の一覧。所見が文字 (A09：… のような) の人はここに出る想定。'
W '  5: GANTEI の所見に判定記号が無ければ、コード入力の人も自動判定で判定は付かない。'
W '  6: 600060 が結果票に出る眼底の判定。業者の判定 (9/3: A39 B9 C2 D6 F1 G1 / 9/9・9/10: A100 B10 C2 D27 G2) と合っていなければ、'
W '     mapping_gantei_hantei_93.csv / mapping_gantei_hantei_bkup123.csv で業者の判定を 600060 に入れる (HANDOFF 参照)。'
W '  8: 眼底のコメントの定型文。kukuri_comment に眼底の規則を足すときの文章に使う。'
notepad $out
