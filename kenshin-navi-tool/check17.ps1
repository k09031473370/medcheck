﻿<#
  組合の下に加盟会社をぶら下げる枠があるか調べる (check17.ps1)
  check17.bat をダブルクリックすると実行され、結果 r_check17.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  やりたいこと: 石材組合の健診。組合に小さい会社が何社も加盟していて、
  「同じ組合の中の会社」として登録したい。

  健診ナビには 団体マスタ(M_DANTAI) のほかに T_DANTAI1 / T_DANTAI2 / T_BUSYO があり、
  団体コードが DANTAI_CD1 という名前なので、2段以上の階層を持てる可能性が高い。
  実際にどう使えるのか、すでに同じ使い方をしている団体が無いかを調べる。

  ※ 自宅の練習用DBでも 3〜6 は見られます。1・2・7 はクリニックのPCで。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check17.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 組合と加盟会社の枠 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 団体マスタ(M_DANTAI)の列 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'M_DANTAI'
ORDER BY ORDINAL_POSITION
"@ 200

Q '--- 2. 団体コードが2段になっていないか (CD2・親・上位 らしい列を全テーブルから探す) ---' @"
SELECT TABLE_NAME AS テーブル, COLUMN_NAME AS 列, DATA_TYPE AS 型
FROM INFORMATION_SCHEMA.COLUMNS
WHERE COLUMN_NAME LIKE '%DANTAI%'
   OR COLUMN_NAME LIKE '%KUMIAI%' OR COLUMN_NAME LIKE '%OYA%'
   OR COLUMN_NAME LIKE '%GROUP%'  OR COLUMN_NAME LIKE '%JOI%'
ORDER BY TABLE_NAME, ORDINAL_POSITION
"@ 200

Q '--- 3. T_DANTAI1 の列 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_DANTAI1'
ORDER BY ORDINAL_POSITION
"@ 200

Q '--- 4. T_DANTAI2 の列 (2段目にあたる可能性) ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_DANTAI2'
ORDER BY ORDINAL_POSITION
"@ 200

Q '--- 5. 部署マスタ T_BUSYO の列 (加盟会社を部署として持たせる手もあるため) ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_BUSYO'
ORDER BY ORDINAL_POSITION
"@ 200

Q '--- 6. すでに「組合」「協会」「連合」などの団体が登録されていないか ---' @"
SELECT TOP 40 LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, MEISYO1 AS 団体名
FROM M_DANTAI
WHERE MEISYO1 LIKE N'%組合%' OR MEISYO1 LIKE N'%協会%'
   OR MEISYO1 LIKE N'%連合%' OR MEISYO1 LIKE N'%会%'
ORDER BY DANTAI_CD1
"@ 60

Q '--- 7. 1つの団体に部署がいくつぶら下がっているか (多い順) ---' @"
SELECT TOP 20 LTRIM(RTRIM(b.DANTAI_CD1)) AS 団体CD, d.MEISYO1 AS 団体名, COUNT(*) AS 部署数
FROM T_BUSYO b LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = b.DANTAI_CD1
GROUP BY LTRIM(RTRIM(b.DANTAI_CD1)), d.MEISYO1
ORDER BY COUNT(*) DESC
"@ 30

Q '--- 8. 部署の中身の例 (会社名らしいものが入っているか) ---' @"
SELECT TOP 40 LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(BUSYO_CD)) AS 部署CD, *
FROM T_BUSYO
ORDER BY DANTAI_CD1, BUSYO_CD
"@ 40

Q '--- 9. 受診データ側で、団体・部署をどう持っているか ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'T_KENSIN'
  AND (COLUMN_NAME LIKE '%DANTAI%' OR COLUMN_NAME LIKE '%BUSYO%' OR COLUMN_NAME LIKE '%KENPO%')
ORDER BY ORDINAL_POSITION
"@ 60

Q '--- 10. 個人マスタ側の所属の持ち方 ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'T_KOJIN1'
  AND (COLUMN_NAME LIKE '%DANTAI%' OR COLUMN_NAME LIKE '%BUSYO%' OR COLUMN_NAME LIKE '%SYOZOKU%')
ORDER BY ORDINAL_POSITION
"@ 60

W ''
W '=== 読み方 ==='
W '  2 に DANTAI_CD2 のような列があった'
W '      → 団体そのものが2段構造。組合=1段目、加盟会社=2段目 で登録できます。これが本命。'
W '  7 で1つの団体にたくさん部署がぶら下がっている団体がある'
W '      → 「団体=組合 / 部署=加盟会社」という使い方が、このクリニックでの前例です。'
W '        8 でその中身が会社名になっていれば確定です。'
W '  どちらも無い'
W '      → 加盟会社ごとに団体を作り、団体コードの付け方で組合をまとめる形になります'
W '        (例: 組合が 0000000200 なら、加盟会社を 0000000201, 0000000202 …)。'
W '        請求や集計はツール側でまとめられます。'

notepad $out
