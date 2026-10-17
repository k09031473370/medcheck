<#
  石屋: 中性脂肪と血糖の「振り分け」と判定が、受診日ごとにどこまで付いているか (check100.ps1)
  check100.bat をダブルクリックすると実行され、結果 r_check100.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  背景: 中性脂肪(未判別) R06-0002 と 血糖(未判別) 068682 は、自動判定が採血時間 001282 を見て
        空腹時中性脂肪 068702 / 随時中性脂肪 R06-0003、空腹時血糖 H30-0011 / 随時血糖 H30-0012 に写してから判定する。
        9/9・9/10 は採血時間を自動判定のあとに入れたので、写っていない可能性が高い。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check100.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: 中性脂肪・血糖の振り分けと判定 (受診日ごと) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
function Val($cd) { return "ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.KEKKA,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd'),'')" }
function Han($cd) { return "ISNULL((SELECT TOP 1 LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '$cd'),'')" }

Q '--- 1. ★受診日ごと: 中性脂肪 (未判別の値 / 空腹時に写った人 / 随時に写った人 / 判定が付いている人) ---' @"
SELECT p.YMD AS 受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN p.TG <> '' THEN 1 ELSE 0 END) AS 中性脂肪の値あり,
       SUM(CASE WHEN p.TG_H <> '' THEN 1 ELSE 0 END) AS 未判別に判定あり,
       SUM(CASE WHEN p.TGF <> '' THEN 1 ELSE 0 END) AS 空腹時に値あり,
       SUM(CASE WHEN p.TGF_H <> '' THEN 1 ELSE 0 END) AS 空腹時に判定あり,
       SUM(CASE WHEN p.TGZ <> '' THEN 1 ELSE 0 END) AS 随時に値あり,
       SUM(CASE WHEN p.TGZ_H <> '' THEN 1 ELSE 0 END) AS 随時に判定あり,
       SUM(CASE WHEN p.TG_H = '' AND p.TGF_H = '' AND p.TGZ_H = '' THEN 1 ELSE 0 END) AS [★中性脂肪の判定なし]
FROM (SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD,
             $(Val 'R06-0002') AS TG, $(Han 'R06-0002') AS TG_H, $(Val '068702') AS TGF, $(Han '068702') AS TGF_H, $(Val 'R06-0003') AS TGZ, $(Han 'R06-0003') AS TGZ_H
      FROM T_KENSIN s WHERE $ISHIYA) p
GROUP BY p.YMD ORDER BY 1
"@ 10

Q '--- 2. ★受診日ごと: 血糖 (未判別の値 / 空腹時に写った人 / 随時・食直後に写った人 / 判定) ---' @"
SELECT p.YMD AS 受診日, COUNT(*) AS 人数,
       SUM(CASE WHEN p.BS <> '' THEN 1 ELSE 0 END) AS 血糖の値あり,
       SUM(CASE WHEN p.BS_H <> '' THEN 1 ELSE 0 END) AS 未判別に判定あり,
       SUM(CASE WHEN p.BSF <> '' THEN 1 ELSE 0 END) AS 空腹時に値あり,
       SUM(CASE WHEN p.BSZ <> '' OR p.BSS <> '' THEN 1 ELSE 0 END) AS 随時食直後に値あり,
       SUM(CASE WHEN p.BSF = '' AND p.BSZ = '' AND p.BSS = '' THEN 1 ELSE 0 END) AS [★振り分けなし]
FROM (SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD,
             $(Val '068682') AS BS, $(Han '068682') AS BS_H, $(Val 'H30-0011') AS BSF, $(Val 'H30-0012') AS BSZ, $(Val 'H30-0013') AS BSS
      FROM T_KENSIN s WHERE $ISHIYA) p
GROUP BY p.YMD ORDER BY 1
"@ 10

Q '--- 3. 脂質代謝のくくり (600320) の判定分布 (受診日ごと)。中性脂肪が判定されていなければ、中性脂肪だけ高い人の脂質代謝が A のまま ---' @"
SELECT p.YMD AS 受診日, p.SHI AS 脂質代謝, COUNT(*) AS 人数
FROM (SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, $(Han '600320') AS SHI FROM T_KENSIN s WHERE $ISHIYA) p
GROUP BY p.YMD, p.SHI ORDER BY 1, 2
"@ 30

Q '--- 4. ★中性脂肪の判定が無い人のうち、値が 150 以上 (判定が付けば B 以上になる人) ---' @"
SELECT p.YMD AS 受診日, p.UKE AS 受付番号, p.NAME AS 氏名, p.TG AS 中性脂肪,
       CASE WHEN TRY_CONVERT(float, p.TG) >= 500 THEN 'F' WHEN TRY_CONVERT(float, p.TG) >= 300 THEN 'D' WHEN TRY_CONVERT(float, p.TG) >= 150 THEN 'B' ELSE 'A' END AS 空腹時基準なら,
       p.SHI AS 今の脂質代謝, p.SOGO AS 今の総合判定
FROM (SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS YMD, s.UKE_NO_KENSA AS UKE, g.KANJI_SIMEI AS NAME,
             $(Val 'R06-0002') AS TG, $(Han 'R06-0002') AS TG_H, $(Han '068702') AS TGF_H, $(Han 'R06-0003') AS TGZ_H, $(Han '600320') AS SHI, $(Han '11000') AS SOGO
      FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID WHERE $ISHIYA) p
WHERE p.TG_H = '' AND p.TGF_H = '' AND p.TGZ_H = '' AND TRY_CONVERT(float, p.TG) >= 150
ORDER BY p.YMD, TRY_CONVERT(float, p.TG) DESC
"@ 200

W ''
W '=== 読み方 ==='
W '  1: 「★中性脂肪の判定なし」が 0 でない受診日は、中性脂肪が判定に入っていない。9/9 がそうなら、自動判定のかけ直しか、別の手当てが要る。'
W '  2: 「★振り分けなし」は空腹時血糖に写っていない人。血糖(未判別) 自体には判定が付いているので、結果票にどちらが出るかで要否が決まる。'
W '  4: 中性脂肪 150 以上で判定が無い人。この人たちは脂質代謝・総合判定・コメントにも中性脂肪が反映されていない。'
notepad $out
