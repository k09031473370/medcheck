﻿<#
  団体名の【】による束ね方を調べる (check19.ps1)
  check19.bat をダブルクリックすると実行され、結果 r_check19.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check18 で分かったこと:
    T_DANTAIGROUP は 0件。団体グループの仕組みはあるが使われていない。
    そのかわり「【福生】細野石材」のように、名前の頭に【】が付いた団体があった。

  ここで確かめること:
    このクリニックは【】で団体を束ねているのか。
    束ねているなら、石材組合も同じやり方に合わせるのが安全。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check19.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 団体名の【】による束ね方 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 【】が付いた団体は何社あるか。頭の【】ごとの社数 ---' @"
SELECT SUBSTRING(MEISYO1, 1, CHARINDEX(N'】', MEISYO1)) AS 頭の記号, COUNT(*) AS 社数
FROM M_DANTAI
WHERE MEISYO1 LIKE N'【%】%'
GROUP BY SUBSTRING(MEISYO1, 1, CHARINDEX(N'】', MEISYO1))
ORDER BY COUNT(*) DESC
"@ 80

Q '--- 2. 【】付きの団体の一覧 (どういう束ね方をしているか) ---' @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, MEISYO1 AS 団体名,
       LTRIM(RTRIM(ISNULL(RYAKUSYO1,''))) AS 略称
FROM M_DANTAI
WHERE MEISYO1 LIKE N'【%】%'
ORDER BY MEISYO1
"@ 200

Q '--- 3. 細野石材の周りの団体コード (連番でまとまっていないか) ---' @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, MEISYO1 AS 団体名
FROM M_DANTAI
WHERE DANTAI_CD1 BETWEEN '0000000300' AND '0000000330'
ORDER BY DANTAI_CD1
"@ 60

Q '--- 4. 細野石材の受診実績 (いつ・何人・どのコースで来ているか) ---' @"
SELECT TOP 40 s.D_KENSIN AS 受診日, COUNT(*) AS 人数,
       MAX(c1.MEISYO) AS コース例
FROM T_KENSIN s
LEFT JOIN T_COURSE1 c1 ON c1.COURSE_CD = s.COURSE_CD AND c1.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.DANTAI_CD1 = '0000000311' AND s.F_TORIKESI = 0
GROUP BY s.D_KENSIN
ORDER BY s.D_KENSIN DESC
"@ 50

Q '--- 5. 同じ日に来ている他の団体 (組合でまとめて来ているなら同日に並ぶはず) ---' @"
SELECT TOP 60 s.D_KENSIN AS 受診日, LTRIM(RTRIM(s.DANTAI_CD1)) AS 団体CD,
       d.MEISYO1 AS 団体名, COUNT(*) AS 人数
FROM T_KENSIN s
JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0
  AND s.D_KENSIN IN (SELECT DISTINCT D_KENSIN FROM T_KENSIN
                     WHERE DANTAI_CD1 = '0000000311' AND F_TORIKESI = 0)
GROUP BY s.D_KENSIN, LTRIM(RTRIM(s.DANTAI_CD1)), d.MEISYO1
ORDER BY s.D_KENSIN DESC, COUNT(*) DESC
"@ 80

Q '--- 6. 迷子のグループCDを持つ団体はどれか ---' @"
SELECT LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, MEISYO1 AS 団体名,
       LTRIM(RTRIM(DANTAIGROUP_CD)) AS グループCD
FROM M_DANTAI
WHERE LTRIM(RTRIM(ISNULL(DANTAIGROUP_CD,''))) <> ''
"@ 20

Q '--- 7. 請求は団体ごとか、まとめてか (請求データの持ち方) ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'T_SEIKYU1'
ORDER BY ORDINAL_POSITION
"@ 100

W ''
W '=== 読み方 ==='
W '  1・2 で【】が何十社にも付いている'
W '      → このクリニックは名前の【】で束ねる運用をしています。'
W '        石材組合も【石材組合】〇〇石材 という名前にするのが、既存に合っていて安全です。'
W '  1 が数件しかない'
W '      → 【】は単なるメモ書きです。団体グループ(T_DANTAIGROUP)を使う方が筋が良い。'
W '  5 で細野石材と同じ日に他の石材屋が並んでいる'
W '      → すでに組合単位で巡回健診をしています。その並びをそのまま増やす形になります。'

notepad $out
