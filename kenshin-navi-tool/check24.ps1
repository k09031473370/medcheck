<#
  ツールが書いた行と実データの行を全列で突き合わせる (check24.ps1)
  check24.bat をダブルクリックすると実行され、結果 r_check24.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  気づいたこと:
    結果入力画面で、ツールが取り込んだ値(秋葉さん)と実データの値とで文字色が違う。
    健診ナビは結果の「入力元」を色で区別していると思われる。

  つまり:
    実データの行にはあって、ツールが書いた行には入っていない列がある。
    ツールは KEKKA / KEKKA_CD / HANTEI_KIGO の3列しか書いていない。
    自動判定がその足りない列を Substring で切り出しているなら、
    空のまま = 落ちる、で今回のエラーの説明がつく。

  ここでやること:
    T_KENSA の全列について、
      ・秋葉さん(ツールが書いた行)の中身
      ・同じ項目の実データの行の中身
    を並べて、どの列が空のままかを洗い出す。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check24.txt'

# テスト受診者 秋葉達也 2026/08/21 受付6
$PK = 2005094

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== ツールが書いた行と実データの行の全列比較 (PK_SEQ $PK) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_KENSA の全列 (ツールが書いているのは KEKKA / KEKKA_CD / HANTEI_KIGO の3つだけ) ---' @"
SELECT ORDINAL_POSITION AS 順, COLUMN_NAME AS 列, DATA_TYPE AS 型,
       CHARACTER_MAXIMUM_LENGTH AS 長さ, IS_NULLABLE AS NULL可
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'T_KENSA'
ORDER BY ORDINAL_POSITION
"@ 100

Q '--- 2. 身長(001211) 秋葉さんの行 vs 実データの行 (縦に並べて見比べる) ---' @"
SELECT N'★ツール' AS 区分, * FROM T_KENSA WHERE PK_SEQ = $PK AND LTRIM(RTRIM(KOMOKU_CD)) = '001211'
UNION ALL
SELECT TOP 3 N'実データ', * FROM T_KENSA
WHERE LTRIM(RTRIM(KOMOKU_CD)) = '001211' AND PK_SEQ <> $PK
  AND LTRIM(RTRIM(ISNULL(KEKKA,''))) <> ''
ORDER BY 区分
"@ 20

Q '--- 3. 血圧(001253) 秋葉さんの行 vs 実データの行 ---' @"
SELECT N'★ツール' AS 区分, * FROM T_KENSA WHERE PK_SEQ = $PK AND LTRIM(RTRIM(KOMOKU_CD)) = '001253'
UNION ALL
SELECT TOP 3 N'実データ', * FROM T_KENSA
WHERE LTRIM(RTRIM(KOMOKU_CD)) = '001253' AND PK_SEQ <> $PK
  AND LTRIM(RTRIM(ISNULL(KEKKA,''))) <> ''
ORDER BY 区分
"@ 20

Q '--- 4. AST(068636) 秋葉さんの行 vs 実データの行 (血液は取込元が違うので別に見る) ---' @"
SELECT N'★ツール' AS 区分, * FROM T_KENSA WHERE PK_SEQ = $PK AND LTRIM(RTRIM(KOMOKU_CD)) = '068636'
UNION ALL
SELECT TOP 3 N'実データ', * FROM T_KENSA
WHERE LTRIM(RTRIM(KOMOKU_CD)) = '068636' AND PK_SEQ <> $PK
  AND LTRIM(RTRIM(ISNULL(KEKKA,''))) <> ''
ORDER BY 区分
"@ 20

Q '--- 5. 平野さん(08/20 院内 SHTA2)の身長の行 (文字色が違って見えた側の実物) ---' @"
SELECT TOP 5 k.*
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '2026/08/20'
  AND LTRIM(RTRIM(s.COURSE_CD)) = 'SHTA2'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) = '001211'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
"@ 20

Q '--- 6. 受診そのもの(T_KENSIN)の比較: 秋葉さん vs 08/20院内の人 ---' @"
SELECT N'★秋葉' AS 区分, * FROM T_KENSIN WHERE PK_SEQ = $PK
UNION ALL
SELECT TOP 2 N'08/20院内', * FROM T_KENSIN
WHERE D_KENSIN = '2026/08/20' AND LTRIM(RTRIM(COURSE_CD)) = 'SHTA2' AND F_TORIKESI = 0
ORDER BY 区分
"@ 20

Q '--- 7. T_KENSA の列ごとに、実データで中身が入っている率 (空でない件数が多い列 = 本来入るべき列) ---' @"
SELECT COUNT(*) AS 実データ行数,
       SUM(CASE WHEN KEKKA_CD    IS NOT NULL AND LTRIM(RTRIM(KEKKA_CD))    <> '' THEN 1 ELSE 0 END) AS KEKKA_CD入り,
       SUM(CASE WHEN HANTEI_KIGO IS NOT NULL AND LTRIM(RTRIM(HANTEI_KIGO)) <> '' THEN 1 ELSE 0 END) AS HANTEI_KIGO入り
FROM T_KENSA k
JOIN T_KENSIN s ON s.PK_SEQ = k.PK_SEQ
WHERE s.D_KENSIN = '2026/08/20'
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
"@ 10

W ''
W '=== 読み方 ==='
W '  2〜5 が本命です。★ツール の行と 実データ の行を横に見比べて、'
W '  実データ側だけ中身が入っている列を探してください。'
W '  よくあるのは次のような列です:'
W '      入力区分 / 入力元 (手入力か取込かを表す。文字色の正体はたぶんこれ)'
W '      登録日時 / 更新日時 / 担当者'
W '      機器コード / 検査センターコード'
W '      単位 / 基準値の控え'
W '  そういう列が見つかったら、それが文字色の違いの正体であり、'
W '  同時に自動判定が Substring で落ちている原因の可能性が高いです。'
W ''
W '  6 は受診そのものの比較です。ステータス系の列(結果確定フラグなど)が'
W '  実データ側だけ立っていれば、報告書画面が 0件 だった理由もここで説明がつきます。'
W ''
W '  r_check24.txt をそのまま送ってください。'
W '  足りない列が分かれば、ツール側でその列も一緒に書くように直します。'
W '  (実データと同じ形で書けば、文字色も揃い、自動判定も通るはずです)'

notepad $out
