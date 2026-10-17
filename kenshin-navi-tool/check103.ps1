<#
  石屋 199 人: 1 人ずつの補正状況と、外部点検で挙がった漏れの確認 (check103.ps1)
  check103.bat をダブルクリックすると実行され、結果 r_check103.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check103.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋 199 人: 補正状況と漏れの確認 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$SLOTS  = "('11001','11002','11003','11004','11005','11006','11007','11008','11009')"
$RANK   = "CASE {0} WHEN 'A' THEN 1 WHEN 'B' THEN 2 WHEN 'C' THEN 3 WHEN 'D' THEN 4 WHEN 'E' THEN 5 WHEN 'F' THEN 6 WHEN 'G' THEN 7 WHEN 'J' THEN 7 ELSE 0 END"
function Han($cd) { return "MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) END)" }
function Val($cd) { return "MAX(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) END)" }
function Cnt($cd) { return "SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd' THEN 1 ELSE 0 END)" }
# 1 人 1 行の材料 (派生テーブル)。関数名 H は組込みの別名 h (Get-History) とぶつかるので Han/Val/Cnt
$P = @"
(SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME,
        $(Han '11000') AS SOGO, $(Cnt '600060') AS GAN_N, $(Han '600060') AS GAN,
        $(Han 'R06-0002') AS TG, $(Han '068702') AS TG_K, $(Han 'R06-0003') AS TG_Z, $(Val 'R06-0002') AS TG_V,
        $(Han '068689') AS A1C, $(Han '069010') AS PLT, $(Val '069010') AS PLT_V, $(Han '068676') AS EGF,
        $(Han '068682') AS GLU, $(Val '068682') AS GLU_V, $(Han '600340') AS TOU, $(Han '600310') AS KETU,
        SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '' THEN 1 ELSE 0 END) AS CMT,
        SUM(CASE WHEN LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS THEN 1 ELSE 0 END) AS CMT_N,
        $(Val '001282') AS SAIKETU
 FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
 LEFT JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
 WHERE $ISHIYA
 GROUP BY s.PK_SEQ, s.UKE_NO_KENSA, s.D_KENSIN, g.KANJI_SIMEI) p
"@

Q '--- 1. ★199 人の状態 (受診日・受付番号順) ---' @"
SELECT p.YMD AS 受診日, p.UKE AS 受付番号, p.NAME AS 氏名, ISNULL(p.SOGO,'') AS 総合,
       CASE WHEN p.GAN_N = 0 THEN '枠なし' ELSE ISNULL(p.GAN,'') END AS 眼底,
       ISNULL(p.TG,'') + CASE WHEN ISNULL(p.TG_K,'') + ISNULL(p.TG_Z,'') <> '' THEN '(振分' + ISNULL(p.TG_K,'') + ISNULL(p.TG_Z,'') + ')' ELSE '' END AS 中性脂肪,
       ISNULL(p.GLU,'') AS 血糖, ISNULL(p.A1C,'') AS A1c, ISNULL(p.PLT,'') AS 血小板, ISNULL(p.EGF,'') AS eGFR, ISNULL(p.TOU,'') AS 糖代謝,
       CAST(p.CMT AS varchar(2)) + '/' + CAST(p.CMT_N AS varchar(2)) AS コメント, ISNULL(p.SAIKETU,'') AS 採血時間,
       CASE WHEN ISNULL(p.SOGO,'') = '' THEN '自動判定待ち'
            WHEN p.GAN_N = 0 THEN '眼底の枠なし'
            WHEN ISNULL(p.GAN,'') = '' THEN '眼底未入力(先生確認)'
            WHEN ISNULL(p.TG,'') + ISNULL(p.TG_K,'') + ISNULL(p.TG_Z,'') = '' THEN '中性脂肪の判定なし'
            WHEN ISNULL(p.A1C,'') = '' OR ISNULL(p.PLT,'') = '' OR ISNULL(p.EGF,'') = '' THEN 'A1c/血小板/eGFR の判定なし'
            WHEN p.CMT_N > 0 AND p.CMT = p.CMT_N THEN 'コメント枠満杯(要確認)'
            ELSE '済' END AS 状態
FROM $P
ORDER BY p.YMD, p.UKE
"@ 250

Q '--- 2. 状態ごとの人数 (受診日別) ---' @"
SELECT t.YMD AS 受診日, t.ST AS 状態, COUNT(*) AS 人数 FROM (
  SELECT p.YMD,
       CASE WHEN ISNULL(p.SOGO,'') = '' THEN '自動判定待ち'
            WHEN p.GAN_N = 0 THEN '眼底の枠なし'
            WHEN ISNULL(p.GAN,'') = '' THEN '眼底未入力(先生確認)'
            WHEN ISNULL(p.TG,'') + ISNULL(p.TG_K,'') + ISNULL(p.TG_Z,'') = '' THEN '中性脂肪の判定なし'
            WHEN ISNULL(p.A1C,'') = '' OR ISNULL(p.PLT,'') = '' OR ISNULL(p.EGF,'') = '' THEN 'A1c/血小板/eGFR の判定なし'
            WHEN p.CMT_N > 0 AND p.CMT = p.CMT_N THEN 'コメント枠満杯(要確認)'
            ELSE '済' END AS ST
  FROM $P) t
GROUP BY t.YMD, t.ST ORDER BY 1, 2
"@ 40

Q '--- 3. ★血糖(未判別) 068682: 値があるのに判定が無い人 / 血糖の判定が糖代謝より重い人 (補正の漏れ候補) ---' @"
SELECT p.YMD AS 受診日, p.UKE AS 受付番号, p.NAME AS 氏名, ISNULL(p.GLU_V,'') AS 血糖値, ISNULL(p.GLU,'') AS 血糖判定, ISNULL(p.TOU,'') AS 糖代謝, ISNULL(p.SOGO,'') AS 総合,
       ISNULL(p.SAIKETU,'') AS 採血時間,
       CASE WHEN ISNULL(p.GLU_V,'') <> '' AND ISNULL(p.GLU,'') = '' THEN '値あり判定なし' ELSE '血糖の方が重い' END AS 理由
FROM $P
WHERE ISNULL(p.SOGO,'') <> ''
  AND ((ISNULL(p.GLU_V,'') <> '' AND ISNULL(p.GLU,'') = '')
       OR ($($RANK -f "ISNULL(p.GLU,'')") > $($RANK -f "ISNULL(p.TOU,'')")))
ORDER BY p.YMD, p.UKE
"@ 60

Q '--- 4. 血糖の判定の分布 (受診日 × 判定) ---' @"
SELECT p.YMD AS 受診日, ISNULL(p.GLU,'') AS 血糖判定, COUNT(*) AS 人数 FROM $P GROUP BY p.YMD, ISNULL(p.GLU,'') ORDER BY 1, 2
"@ 40

Q '--- 5. ★血小板が B 以外の異常 (D/E/F/G) の人と、血液一般・血の字を含むコメント ---' @"
SELECT p.YMD AS 受診日, p.UKE AS 受付番号, p.NAME AS 氏名, ISNULL(p.PLT_V,'') AS 血小板, ISNULL(p.PLT,'') AS 判定, ISNULL(p.KETU,'') AS 血液一般, ISNULL(p.SOGO,'') AS 総合,
       ISNULL((SELECT TOP 1 LEFT(LTRIM(RTRIM(c.KEKKA)), 40) FROM T_KENSA c WHERE c.PK_SEQ = p.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN $SLOTS AND (c.KEKKA LIKE '%血小板%' OR c.KEKKA LIKE '%血液%' OR c.KEKKA LIKE '%貧血%')),'(なし)') AS 血液のコメント
FROM $P
WHERE ISNULL(p.PLT,'') IN ('D','E','F','G')
ORDER BY p.YMD, p.UKE
"@ 40

Q '--- 6. 2026 年に使われた「血小板」を含むコメントの定型文 (血小板のコメントを足すときの文章候補) ---' @"
SELECT TOP 15 LTRIM(RTRIM(k.KEKKA)) AS 文章, ISNULL(k.KEKKA_CD,'') AS 番号, COUNT(*) AS 件数
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN >= '2026/01/01' AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS AND k.KEKKA LIKE '%血小板%'
GROUP BY LTRIM(RTRIM(k.KEKKA)), ISNULL(k.KEKKA_CD,'') ORDER BY 3 DESC
"@ 20

Q '--- 7. ★J (治療中) が付いている欄 (J は G と同じ重さで扱っているので、J の総合判定の人は他の項目が F でも上がらない) ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE $ISHIYA AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) = 'J'
GROUP BY CONVERT(varchar(10), s.D_KENSIN, 111), s.UKE_NO_KENSA, LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1, 2, 3
"@ 60

Q '--- 8. ★コメントの「既にあり」誤認の候補: 糖代謝/糖尿病/腎機能/中性脂肪/眼底 を含み、かつ「異常なし」「治療」を含む文 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(k.KOMOKU_CD)) AS 枠, LEFT(LTRIM(RTRIM(k.KEKKA)), 60) AS 文章
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE $ISHIYA AND LTRIM(RTRIM(k.KOMOKU_CD)) IN $SLOTS
  AND (k.KEKKA LIKE '%糖代謝%' OR k.KEKKA LIKE '%糖尿病%' OR k.KEKKA LIKE '%腎機能%' OR k.KEKKA LIKE '%中性脂肪%' OR k.KEKKA LIKE '%眼底%')
  AND (k.KEKKA LIKE '%異常なし%' OR k.KEKKA LIKE '%異常はなし%' OR k.KEKKA LIKE '%治療%')
ORDER BY 1, 2, 3
"@ 60

Q '--- 9. 採血時間 001282 の分布 (1 = 食後10時間以上 / 2 = 3.5〜10 時間 / 3 = 3.5 時間未満) ---' @"
SELECT p.YMD AS 受診日, ISNULL(p.SAIKETU,'') AS 採血時間, COUNT(*) AS 人数 FROM $P GROUP BY p.YMD, ISNULL(p.SAIKETU,'') ORDER BY 1, 2
"@ 20

W ''
W '=== 読み方 ==='
W '  1・2: 状態が「済」以外の人が残り。9/10 の「自動判定待ち」「眼底の枠なし」は 35 人の自動判定のあとに流し直せば消える想定。'
W '  3: 0 行なら血糖の判定は自動判定で付いていて、糖代謝にも入っている。行が出たら kukuri_fix で直す (血糖も対象に入れた)。'
W '  5: 血小板 F の人に血液のコメントが無ければ、6 の定型文を使って足す。'
W '  7: J の総合判定の人は、他の項目が F でも総合判定が上がっていない。先生に見てもらう。'
W '  8: ここに出た文は、コメント追加のツールが「同じ種類の文が既にある」と見て足さなかった可能性がある。'
W '  9: 9/9・9/10 は全員 1 にした (業者ファイルでは食後 3.5 時間・1 時間の人や空欄の人がいた)。'
notepad $out
