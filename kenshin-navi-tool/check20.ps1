<#
  自動判定がエラーになる原因の項目を探す (check20.ps1)
  check20.bat をダブルクリックすると実行され、結果 r_check20.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  起きたこと:
    テスト取込した秋葉さん(PK_SEQ 2005094)で「自動判定」を押すと
    「インデックスおよび長さは文字列内の場所を参照しなければなりません。パラメータ名: length」
    という .NET のエラーが出る。

  これは健診ナビが結果の文字列を Substring で切り出そうとして、
  想定より短い・形の違う値に当たったときに出るエラー。
  つまり、取り込んだ値のどれかの「形」が実データと違う。

  ここでやること:
    秋葉さんの各項目の値を、他の受診者の実際の値と突き合わせ、
    「長さが実データの範囲から外れている」「実データでは必ずコードが入るのに空」
    といった不一致を機械的に洗い出す。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check20.txt'

# テスト受診者 秋葉達也 2026/08/21 受付6
$PK = 2005094

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 自動判定エラーの原因項目さがし (PK_SEQ $PK) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 疑い順: 値の長さが、他の受診者の実データの範囲(最短〜最長)から外れている項目 ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD,
       MAX(m.MEISYO1)            AS 項目名,
       MAX(a.KEKKA)              AS テストの値,
       MAX(LEN(a.KEKKA))         AS 長さ,
       MIN(LEN(k.KEKKA))         AS 実データ最短,
       MAX(LEN(k.KEKKA))         AS 実データ最長,
       MAX(k.KEKKA)              AS 実データの例
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
JOIN T_KENSA k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
             AND k.PK_SEQ <> a.PK_SEQ
             AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
GROUP BY LTRIM(RTRIM(a.KOMOKU_CD))
HAVING MAX(LEN(a.KEKKA)) > MAX(LEN(k.KEKKA))
    OR MAX(LEN(a.KEKKA)) < MIN(LEN(k.KEKKA))
ORDER BY 項目CD
"@ 100

Q '--- 2. 実データではほぼ必ず結果CD(KEKKA_CD)が入るのに、テストでは空の項目 ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD,
       MAX(m.MEISYO1)            AS 項目名,
       MAX(a.KEKKA)              AS テストの値,
       MAX(LTRIM(RTRIM(ISNULL(a.KEKKA_CD,'')))) AS テストのCD,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) <> '' THEN 1 ELSE 0 END) AS 実データでCD入り,
       COUNT(*) AS 実データ件数,
       MAX(k.KEKKA_CD) AS 実データCDの例
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
JOIN T_KENSA k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
             AND k.PK_SEQ <> a.PK_SEQ
             AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
  AND LTRIM(RTRIM(ISNULL(a.KEKKA_CD,''))) = ''
GROUP BY LTRIM(RTRIM(a.KOMOKU_CD))
HAVING SUM(CASE WHEN LTRIM(RTRIM(ISNULL(k.KEKKA_CD,''))) <> '' THEN 1 ELSE 0 END) * 10 > COUNT(*) * 9
ORDER BY 項目CD
"@ 100

Q '--- 3. テストの値に全角文字・コロン・スペースが入っている項目 (実データと形が違いやすい) ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD,
       m.MEISYO1                 AS 項目名,
       a.KEKKA                   AS テストの値,
       LEN(a.KEKKA)              AS 長さ
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
  AND (a.KEKKA LIKE N'%[０-９：　]%'
       OR a.KEKKA LIKE '%:%'
       OR a.KEKKA LIKE '% %')
ORDER BY 項目CD
"@ 100

Q '--- 4. 参考: テストで値が入っている全項目の一覧 ---' @"
SELECT LTRIM(RTRIM(a.KOMOKU_CD)) AS 項目CD,
       m.MEISYO1  AS 項目名,
       a.KEKKA    AS 値,
       LTRIM(RTRIM(ISNULL(a.KEKKA_CD,'')))    AS 結果CD,
       LTRIM(RTRIM(ISNULL(a.HANTEI_KIGO,''))) AS 判定
FROM T_KENSA a
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(a.KOMOKU_CD))
WHERE a.PK_SEQ = $PK
  AND LTRIM(RTRIM(ISNULL(a.KEKKA,''))) <> ''
ORDER BY 項目CD
"@ 200

W ''
W '=== 読み方 ==='
W '  1 に出た項目が最有力です。値の長さが実データの範囲から外れている項目は、'
W '     自動判定が Substring で切り出すときにこのエラーを起こします。'
W '     (例: 実データは「1.0」なのにテストは「0.1」のような桁違い、'
W '          実データは「130」なのにテストは「7」など)'
W '  2 に出た項目は、実データでは結果CDが必ず入るのにテストでは空のもの。'
W '     判定処理がCD側を切り出す作りだと、ここでも同じエラーになります。'
W '  3 は全角・コロン・スペース入りの値。時刻などは形式違いに注意。'
W '  1〜3 が全部空なら、原因はテストデータではなく元々の枠の可能性が高いので、'
W '     r_check20.txt を送ってください。次を考えます。'
W ''
W '  原因の項目が分かったら、直し方は2つ:'
W '    a) 結果入力画面でその項目の値だけ実データらしい値に手で直して、もう一度 自動判定'
W '    b) どうせテストなので、ここで止めて 元に戻す.bat で片付ける'

notepad $out
