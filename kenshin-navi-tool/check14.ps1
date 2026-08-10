<#
  東振協の胸部X線コードを一括入力できるか調べる (check14.ps1)
  check14.bat をダブルクリックすると実行され、結果 r_check14.txt がメモ帳で開きます。
  DBもファイルも読むだけで、一切変更しません。

  やりたいこと: 連名入力で部位と所見を毎回選ぶのをやめ、
  東振協指定の5文字コードを一覧で流し込めるようにしたい。

  必要なのは「5文字コード → 部位CD + 所見CD」の変換表。
  そのために、健診ナビ側の選択肢マスタと、東振協まわりのファイルを調べる。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check14.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 東振協の胸部X線コード 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 胸部X線の「部位」の選択肢 (ZK020) ---' @"
SELECT KEKKA_CD AS 結果CD, SYOKEN AS 表示, HANTEI_KIGO AS 判定
FROM T_SYOKEN2 WHERE SYOKEN_CD = 'ZK020' ORDER BY KEKKA_CD
"@ 200

Q '--- 2. 胸部X線の「所見」の選択肢 (ZK021) ---' @"
SELECT KEKKA_CD AS 結果CD, SYOKEN AS 表示, HANTEI_KIGO AS 判定
FROM T_SYOKEN2 WHERE SYOKEN_CD = 'ZK021' ORDER BY KEKKA_CD
"@ 300

Q '--- 3. 所見リストの一覧 (東振協用の別リストが無いか) ---' @"
SELECT s.SYOKEN_CD AS リスト, COUNT(*) AS 選択肢数, MIN(s.SYOKEN) AS 例
FROM T_SYOKEN2 s GROUP BY s.SYOKEN_CD ORDER BY s.SYOKEN_CD
"@ 100

Q '--- 4. 胸部X線まわりの項目 (連名入力で使う項目CDを特定する) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名, SYOKEN_CD AS 選択肢リスト
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%胸部%' OR MEISYO1 LIKE N'%ﾚﾝﾄｹﾞﾝ%' OR MEISYO1 LIKE N'%レントゲン%'
ORDER BY KOMOKU_CD
"@ 60

Q '--- 5. 東振協コースの人に今なにが入っているか (直近の実データ) ---' @"
SELECT TOP 40 s.D_KENSIN AS 受診日, s.UKE_NO_KENSA AS 受付番号, s.COURSE_CD AS コース,
       LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, k.KEKKA AS 結果, k.KEKKA_CD AS 結果CD
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
WHERE s.F_TORIKESI = 0 AND s.COURSE_CD LIKE 'T%'
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('077010A','077010B','077010Z','077300A')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> ''
ORDER BY s.D_KENSIN DESC, s.UKE_NO_KENSA
"@ 50

Q '--- 6. 東振協らしいマスタ/テーブル ---' @"
SELECT DISTINCT t.TABLE_NAME AS テーブル
FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME
WHERE c.COLUMN_NAME LIKE '%TOUSIN%' OR c.COLUMN_NAME LIKE '%TOSIN%'
   OR t.TABLE_NAME LIKE '%TOUSIN%' OR t.TABLE_NAME LIKE '%TOSIN%'
ORDER BY t.TABLE_NAME
"@ 40

# ---- 東振協まわりのファイル ----
W ''
W '--- 7. 東振協まわりのフォルダの中身 (コード表があるかもしれない) ---'
foreach ($p in @('\\KNSV\KenshinNavi\東振協H30', '\\KNSV\KenshinNavi\vbR_連名帳票',
                 '\\KNSV\KenshinNavi\CSV', '\\KNSV\KenshinNavi\設定dat')) {
    W ''
    W "[$p]"
    if (-not (Test-Path $p)) { W '  見つかりません'; continue }
    try {
        Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue |
            Sort-Object Name | Select-Object -First 40 |
            ForEach-Object { W ('  {0}  ({1} KB)' -f $_.Name, [math]::Round($_.Length/1KB,1)) }
    } catch { W ("  読めませんでした: {0}" -f $_.Exception.Message) }
}

W ''
W '--- 8. 上のフォルダにあるCSV/DATの中身を少しだけ表示 (コード表か確認する) ---'
foreach ($p in @('\\KNSV\KenshinNavi\東振協H30')) {
    if (-not (Test-Path $p)) { continue }
    foreach ($f in (Get-ChildItem -Path $p -File -Include '*.csv','*.dat','*.txt' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 5)) {
        W ''
        W ("[{0}]" -f $f.FullName)
        try {
            Get-Content -Path $f.FullName -TotalCount 8 -Encoding Default |
                ForEach-Object { W "  $_" }
        } catch { W '  読めませんでした' }
    }
}

W ''
W '※ 5文字コードの実物を数個(できれば所見名つきで)教えてもらえると、変換表を作れます。'
W '=== 完了。この内容をチャットに貼り付けてください ==='
notepad $out
