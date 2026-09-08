<#
  既往歴と自覚症状の入れ場所を調べる (check48.ps1)
  check48.bat をダブルクリックすると実行され、結果 r_check48.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  東振協のCSVには既往歴(39列)と自覚症状(13列)が入っているが、まだ取り込んでいない。
  結果報告書には枠がある (既往歴4枠・自覚症状4枠) ので、健診ナビ側の受け皿を探す。

    自覚症状 … 017107A〜J の10枠と見られる。選択肢の一覧を確認する。
    既往歴   … 帳票の差込が **既往歴_状況1_1 と形が違う。
                検査結果(T_KENSA)ではなく別のテーブルの可能性が高い。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21'
)

$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check48.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 既往歴・自覚症状の入れ場所の調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# ---------------- 自覚症状 ----------------
Q '--- 1. ★自覚症状の項目 (017107 まわり) ---' @"
SELECT LTRIM(RTRIM(m.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       '[' + LTRIM(RTRIM(ISNULL(m.SYOKEN_CD,''))) + ']' AS 所見リスト
FROM T_KOMOKU m
WHERE LTRIM(RTRIM(m.KOMOKU_CD)) LIKE '017107%'
   OR m.MEISYO1 LIKE N'%自覚症状%'
ORDER BY 1
"@ 40

Q "--- 2. ★$Ymd に自覚症状の枠があるか ---" @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       COUNT(*) AS 枠の数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#') THEN 1 ELSE 0 END) AS 値あり
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '017107%'
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD))
ORDER BY 1
"@ 20

Q '--- 3. ★自覚症状の選択肢の一覧 (CSVの症状名と突き合わせる材料) ---' @"
SELECT LTRIM(RTRIM(y.KEKKA_CD)) AS 結果CD, y.SYOKEN AS 症状,
       '[' + LTRIM(RTRIM(ISNULL(y.HANTEI_KIGO,''))) + ']' AS 判定
FROM T_SYOKEN2 y
WHERE LTRIM(RTRIM(y.SYOKEN_CD)) = (
      SELECT TOP 1 LTRIM(RTRIM(m.SYOKEN_CD)) FROM T_KOMOKU m
      WHERE LTRIM(RTRIM(m.KOMOKU_CD)) = '017107A')
ORDER BY LEN(LTRIM(RTRIM(y.KEKKA_CD))), LTRIM(RTRIM(y.KEKKA_CD))
"@ 80

Q '--- 4. 自覚症状が実際に入っている過去の例 ---' @"
SELECT TOP 20 CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日,
       g.KANJI_SIMEI AS 氏名,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA,''))) + ']'    AS 結果,
       '[' + LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) + ']' AS 結果CD
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE LTRIM(RTRIM(k.KOMOKU_CD)) LIKE '017107%'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
ORDER BY s.D_KENSIN DESC
"@ 25

# ---------------- 既往歴 ----------------
Q '--- 5. ★★既往歴らしきテーブルを探す ---' @"
SELECT c.TABLE_NAME AS テーブル, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME LIKE '%KIOU%' OR c.TABLE_NAME LIKE '%KAZOKU%'
   OR c.TABLE_NAME LIKE '%BYOUMEI%' OR c.TABLE_NAME LIKE '%JOUKYOU%'
ORDER BY c.TABLE_NAME, c.ORDINAL_POSITION
"@ 80

Q '--- 6. ★既往歴の中身の例 (T_KIOU_HANTEI) ---' @"
SELECT TOP 20 * FROM T_KIOU_HANTEI
"@ 25

Q '--- 7. 既往歴の病名マスタらしきもの ---' @"
SELECT c.TABLE_NAME AS テーブル, COUNT(*) AS 列数
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME LIKE 'M_%' AND (c.COLUMN_NAME LIKE '%KIOU%' OR c.COLUMN_NAME LIKE '%BYOU%')
GROUP BY c.TABLE_NAME
ORDER BY 1
"@ 40

Q "--- 8. $Ymd の人が既往歴のテーブルに行を持っているか ---" @"
SELECT COUNT(*) AS 既往歴の行数
FROM T_KIOU_HANTEI k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
"@ 10

W ''
W '=== 読み方 ==='
W '  1〜3 で自覚症状の受け皿が分かります。'
W '     3 に症状の一覧が出れば、CSVの13の症状名と突き合わせて対応表が作れます。'
W '     2 の「枠の数」が88なら、そのまま書き込めます。'
W ''
W '  4 に過去の実例が出れば、どういう形で入れるのが正しいかが分かります。'
W ''
W '  5〜8 が既往歴です。T_KIOU_HANTEI がそれらしいので中身を見ます。'
W '     検査結果(T_KENSA)とは別のテーブルなら、書き込み方も別に作る必要があります。'
W '     8 が 0 なら、8/21 の人には既往歴の行がまだ無いということです。'
W ''
W '  ※ 既往歴は健診ナビの作りが分からないうちは書き込みません。'
W '     まず実物を見てから、作れるかどうかを判断します。'

notepad $out
