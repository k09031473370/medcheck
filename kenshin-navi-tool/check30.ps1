<#
  選択肢と所見の中身を書き出す (check30.ps1)
  check30.bat をダブルクリックすると実行され、結果 r_check30.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  check28 で、項目ごとに使う「リスト名」が分かった。
    問診 : ｴ / ﾃ / ﾆ / TYO / NYOU1 / TEISEI / NINSIN / KITUEN /
           YAKU01〜03 / INSYU_SYU / INSYU_RYO / TOKUMON11 / TOKUMON14
    所見 : SHIN(内科診察) / ZK011(心電図) / ZK020・ZK021(胸部X線) /
           ZK030・ZK031(胃部X線)

  ここではその中身(正しい書き方)を全部書き出す。
  東振協ファイルの書き方と突き合わせて、変換表を仕上げるために使う。

  ※ check28 の調査1はSQLの列名を間違えていた。
     T_SYOKEN2 の列は SYOKEN_CD / KEKKA_CD / SYOKEN / HANTEI_KIGO。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check30.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 選択肢と所見の中身 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★問診の選択肢 (はい/いいえ以外のもの) ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS リスト, LTRIM(RTRIM(KEKKA_CD)) AS 結果CD,
       SYOKEN AS 選択肢, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) IN
      ('KITUEN','INSYU_SYU','INSYU_RYO','TOKUMON11','TOKUMON14',
       'YAKU01','YAKU02','YAKU03','NINSIN','TYO','NYOU1','TEISEI')
ORDER BY SYOKEN_CD, KEKKA_CD
"@ 200

Q '--- 2. ★1文字のリスト (ｴ = はい/いいえ系、ﾃ = 食べる速度、ﾆ = 生活習慣の改善) ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS リスト, LTRIM(RTRIM(KEKKA_CD)) AS 結果CD,
       SYOKEN AS 選択肢, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) IN (N'ｴ', N'ﾃ', N'ﾆ')
ORDER BY SYOKEN_CD, KEKKA_CD
"@ 100

Q '--- 3. ★所見: 内科診察(SHIN) と 心電図(ZK011) ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS リスト, LTRIM(RTRIM(KEKKA_CD)) AS 結果CD,
       SYOKEN AS 所見, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) IN ('SHIN','ZK011')
ORDER BY SYOKEN_CD, KEKKA_CD
"@ 400

Q '--- 4. ★所見: 胸部X線の所見(ZK021) ---' @"
SELECT LTRIM(RTRIM(KEKKA_CD)) AS 結果CD, SYOKEN AS 所見,
       LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK021'
ORDER BY KEKKA_CD
"@ 300

Q '--- 5. ★所見: 胃部X線の部位(ZK030)と所見(ZK031) ---' @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS リスト, LTRIM(RTRIM(KEKKA_CD)) AS 結果CD,
       SYOKEN AS 所見, LTRIM(RTRIM(ISNULL(HANTEI_KIGO,''))) AS 判定
FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) IN ('ZK030','ZK031')
ORDER BY SYOKEN_CD, KEKKA_CD
"@ 400

Q '--- 6. 自覚症状(017107A)が使うリスト名と、その中身 ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名,
       LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見リスト
FROM T_KOMOKU WHERE LTRIM(RTRIM(KOMOKU_CD)) LIKE '017107%'
ORDER BY KOMOKU_CD
"@ 30

Q '--- 7. 自覚症状の選択肢 (MN010 など。上のリスト名と合わせて見る) ---' @"
SELECT LTRIM(RTRIM(s.SYOKEN_CD)) AS リスト, LTRIM(RTRIM(s.KEKKA_CD)) AS 結果CD, s.SYOKEN AS 選択肢
FROM T_SYOKEN2 s
WHERE LTRIM(RTRIM(s.SYOKEN_CD)) = (
    SELECT TOP 1 LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) FROM T_KOMOKU
    WHERE LTRIM(RTRIM(KOMOKU_CD)) = '017107A')
ORDER BY s.KEKKA_CD
"@ 200

Q '--- 8. 胸部X線の部位(ZK020) ---' @"
SELECT LTRIM(RTRIM(KEKKA_CD)) AS 結果CD, SYOKEN AS 部位
FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) = 'ZK020'
ORDER BY KEKKA_CD
"@ 200

W ''
W '=== 読み方 ==='
W '  1・2 が問診の正しい書き方です。東振協ファイルの書き方と見比べます。'
W '     例: ファイル「１合未満」に対して 1 に「1合未満」とあれば、'
W '         変換表に 1合未満 への読み替えを1行足すだけで入るようになります。'
W '  3〜5・8 が所見の正しい書き方です。'
W '     ファイルには「瘢痕像」「心陰影拡大あり」「慢性胃炎」「平低Ｔ波」などが出てきます。'
W '     同じ意味の所見が別の書き方で載っていれば、それに読み替えます。'
W '     載っていなければ、その所見は取り込めません(担当者に相談)。'
W '  6・7 は自覚症状です。健診ナビは10枠あるので、ファイルの13項目のうち'
W '     10個まで入れられます。選択肢の並びを見て、対応を決めます。'

notepad $out
