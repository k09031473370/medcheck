<#
  事業所管理番号が社員番号かどうかを確かめる (check31.ps1)
  check31.bat をダブルクリックすると実行され、結果 r_check31.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  東振協ファイルの3列目「事業所管理番号」は、88人が全員違う番号を持っている
  (100320 / 500512 / 200439 / 100508 …)。事業所ごとの番号なら全員同じはずなので、
  実際には人を識別する番号 = 社員番号ではないかと考えられる。

  ここで確かめること:
    健診ナビ側に、城西学園の98人の社員番号(T_KOJIN1.KOJIN_NO)が
    すでに入っているか。入っていれば、予約取込のときに登録された番号なので、
    ファイルの番号と一致するはず。
    空なら、今回はじめて分かる情報ということになる。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check31.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 事業所管理番号は社員番号か $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★本命: 城西学園(8/21)98人の社員番号が入っているか ---' @"
SELECT CASE WHEN LTRIM(RTRIM(ISNULL(k.KOJIN_NO,''))) = '' THEN N'社員番号なし' ELSE N'社員番号あり' END AS 区分,
       COUNT(*) AS 人数
FROM T_KENSIN s LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
GROUP BY CASE WHEN LTRIM(RTRIM(ISNULL(k.KOJIN_NO,''))) = '' THEN N'社員番号なし' ELSE N'社員番号あり' END
"@ 10

Q '--- 2. 実際の中身 (先頭40人。ファイルの番号と見比べる) ---' @"
SELECT TOP 40 k.KANJI_SIMEI AS 漢字氏名, k.KANA_SIMEI AS カナ氏名,
       '[' + LTRIM(RTRIM(ISNULL(k.KOJIN_NO,''))) + ']' AS 社員番号,
       s.PK_SEQ
FROM T_KENSIN s LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
ORDER BY k.KANA_SIMEI
"@ 60

Q '--- 3. 他の団体では社員番号を使っているか (運用の前例) ---' @"
SELECT TOP 20 LTRIM(RTRIM(d.DANTAI_CD1)) AS 団体CD, MAX(dm.MEISYO1) AS 団体名,
       COUNT(*) AS 人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(d.KOJIN_NO,''))) <> '' THEN 1 ELSE 0 END) AS 社員番号あり
FROM T_KOJIN1 d LEFT JOIN M_DANTAI dm ON dm.DANTAI_CD1 = d.DANTAI_CD1
GROUP BY LTRIM(RTRIM(d.DANTAI_CD1))
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(d.KOJIN_NO,''))) <> '' THEN 1 ELSE 0 END) > 0
ORDER BY COUNT(*) DESC
"@ 30

Q '--- 4. 社員番号の桁数の傾向 (東振協の6桁と合うか) ---' @"
SELECT LEN(LTRIM(RTRIM(KOJIN_NO))) AS 桁数, COUNT(*) AS 人数,
       MIN(LTRIM(RTRIM(KOJIN_NO))) AS 例1, MAX(LTRIM(RTRIM(KOJIN_NO))) AS 例2
FROM T_KOJIN1
WHERE LTRIM(RTRIM(ISNULL(KOJIN_NO,''))) <> ''
GROUP BY LEN(LTRIM(RTRIM(KOJIN_NO)))
ORDER BY COUNT(*) DESC
"@ 30

Q '--- 5. 城西学園の団体コードと、その団体の人の社員番号 ---' @"
SELECT TOP 20 LTRIM(RTRIM(k.DANTAI_CD1)) AS 団体CD, k.KANJI_SIMEI AS 氏名,
       '[' + LTRIM(RTRIM(ISNULL(k.KOJIN_NO,''))) + ']' AS 社員番号
FROM T_KOJIN1 k
WHERE k.DANTAI_CD1 = '0000000370'
ORDER BY k.KANA_SIMEI
"@ 30

W ''
W '=== 読み方 ==='
W '  1 で「社員番号あり」が98人 → 予約取込のときに登録済み。'
W '     2 の中身がファイルの番号(100320 など)と同じなら、この列は社員番号で確定。'
W '     同じなら取り込む必要はありません(すでに入っているため)。'
W '  1 で「社員番号なし」が98人 → 予約取込のときには無かった番号です。'
W '     今回はじめて分かる情報なので、入れる価値があります。'
W '     ただし社員番号は結果票にも出る項目なので、入れるかどうかは業務判断です。'
W '  3・4 は他の団体の運用です。社員番号を使っている団体が多く、桁数も6桁が'
W '     中心なら、東振協の番号をそのまま入れて問題ない可能性が高いです。'

notepad $out
