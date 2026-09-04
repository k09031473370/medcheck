<#
  問診の選択肢を健診ナビから書き出す (check28.ps1)
  check28.bat をダブルクリックすると実行され、結果 r_check28.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  なぜ必要か:
    東振協データ送信フォーマットの問診は「１合未満」「意思あり(6ｹ月以内)」のように
    書き方が独特で、健診ナビの選択肢とは綴りが違う。
    健診ナビ側の正しい選択肢が分かれば、value_map.csv に変換を書いて取り込める。

  ここで書き出すのは、東振協フォーマットで使う問診項目の選択肢一覧。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check28.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 問診の選択肢 書き出し $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

# 東振協フォーマットで使う問診・尿・聴力・便潜血の項目
$CDS = @"
'069206','069207','069211',
'067132','067133','067134','067135',
'069245','069246',
'083002','083003',
'084001','084002','084003','084004','084005','084006','084007','084008',
'084009','084010','084011','084012','084014','084015','084017','084018',
'084019','084020','084021',
'H30-0061','H30-0062','R06-0004'
"@ -replace "`r?`n", ' '

Q '--- 1. ★本命: 項目ごとの選択肢 (これが健診ナビの正しい書き方) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, m.MEISYO1 AS 項目名,
       LTRIM(RTRIM(s.SYOKEN_CD2)) AS 選択CD, s.MEISYO AS 選択肢
FROM T_KOMOKU m
JOIN T_SYOKEN2 s ON LTRIM(RTRIM(s.SYOKEN_CD)) = LTRIM(RTRIM(m.SYOKEN_CD))
JOIN (SELECT KOMOKU_CD FROM T_KOMOKU WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ($CDS)) k
     ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(m.KOMOKU_CD))
ORDER BY k.KOMOKU_CD, s.SYOKEN_CD2
"@ 400

Q '--- 2. 上が空だったとき用: 項目と所見グループの対応 ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名,
       LTRIM(RTRIM(ISNULL(SYOKEN_CD,''))) AS 所見グループ
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ($CDS)
ORDER BY KOMOKU_CD
"@ 60

Q '--- 3. 未確定の項目コードを探す (尿素窒素・総蛋白・総ビリルビン・ペプシノゲン・ピロリ・PSA) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%尿素窒素%' OR MEISYO1 LIKE N'%BUN%'
   OR MEISYO1 LIKE N'%総蛋白%'   OR MEISYO1 LIKE N'%総ﾀﾝﾊﾟｸ%'
   OR MEISYO1 LIKE N'%ﾋﾞﾘﾙﾋﾞﾝ%' OR MEISYO1 LIKE N'%ビリルビン%'
   OR MEISYO1 LIKE N'%ﾍﾟﾌﾟｼ%'   OR MEISYO1 LIKE N'%ペプシ%'
   OR MEISYO1 LIKE N'%ﾋﾟﾛﾘ%'    OR MEISYO1 LIKE N'%ピロリ%'
   OR MEISYO1 LIKE N'%PSA%'     OR MEISYO1 LIKE N'%前立腺%'
   OR MEISYO1 LIKE N'%会話%'
ORDER BY KOMOKU_CD
"@ 100

Q '--- 4. 既往歴・自覚症状の枠 (東振協は20病名ぶん送ってくる) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%既往%' OR MEISYO1 LIKE N'%自覚症状%'
ORDER BY KOMOKU_CD
"@ 60

Q '--- 5. 8/21 の受診者98人の氏名 (ファイルの88人と突き合わせるため) ---' @"
SELECT s.PK_SEQ, k.KANJI_SIMEI AS 漢字氏名, k.KANA_SIMEI AS カナ氏名,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),'') + ']' AS 受付番号
FROM T_KENSIN s LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
ORDER BY k.KANA_SIMEI
"@ 120

W ''
W '=== 読み方 ==='
W '  1 が本命です。項目ごとに健診ナビの正しい選択肢が並びます。'
W '     例: 084019 (飲酒量) に「1合未満」と出れば、'
W '         ファイルの「１合未満」(全角1) を value_map.csv でそれに変換します。'
W '  1 が空なら 2 を見てください。所見グループの持ち方が違う可能性があります。'
W '  3 で項目コードが見つかれば、尿素窒素・総蛋白・総ビリルビンなども取り込めるようになります。'
W '  5 はファイルの88人と健診ナビの98人の突き合わせ用です。'
W '     氏名の書き方(スペースの有無など)が違うと氏名照合が外れるので、そこも見ます。'

notepad $out
