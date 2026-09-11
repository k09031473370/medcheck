<#
  身体計測 (BMI・腹囲) の判定設定を確かめる (check58.ps1)
  check58.bat をダブルクリックすると実行され、結果 r_check58.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  8/21の結果では 身体計測は A と C しか出ていない。
  ほかの項目は「学会のC → 当院のD」で読み替えているのに、身体計測だけ C のまま。
  設定を見て、意図的なのか行が足りないのかを確かめる。

  学会の基準 (参考)
    BMI  A 18.5〜24.9 / C 18.4以下・25.0〜29.9 / D 30.0以上
    腹囲  A 男85cm未満・女90cm未満 / C それ以上
#>
[CmdletBinding()]
param([string]$Ymd = '2026/08/21')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check58.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 身体計測の判定設定 ($Ymd) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★★BMI(001217)・腹囲(001215) の基準値 (基準値1) ---' @"
SELECT KIJUN_CD, LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, RENBAN,
       SEIBETU AS 性別, HANTEI_KIGO AS 判定, SAISYO AS 下限, SAIDAI AS 上限,
       HL, HYOJI_YO AS 表示, NENREI_F AS 年齢下, NENREI_T AS 年齢上
FROM T_KIJUN1
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('001215','001217','001211','001212','001213','001214','600020')
ORDER BY KOMOKU_CD, KIJUN_CD, SEIBETU, RENBAN
"@ 80

Q '--- 2. ★8/21 の BMI・腹囲の値と判定 (性別つき) ---' @"
SELECT g.KANJI_SIMEI AS 氏名, g.SEIBETU AS 性別,
       '[' + ISNULL(b.KEKKA,'') + ']' AS BMI,  '[' + ISNULL(b.HANTEI_KIGO,'') + ']' AS 判定BMI,
       '[' + ISNULL(f.KEKKA,'') + ']' AS 腹囲, '[' + ISNULL(f.HANTEI_KIGO,'') + ']' AS 判定腹囲,
       '[' + ISNULL(j.HANTEI_KIGO,'') + ']' AS 判定_身体計測
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KENSA b ON b.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(b.KOMOKU_CD)) = '001217'
LEFT JOIN T_KENSA f ON f.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(f.KOMOKU_CD)) = '001215'
LEFT JOIN T_KENSA j ON j.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(j.KOMOKU_CD)) = '600020'
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(ISNULL(b.KEKKA,''))) NOT IN ('','#')
ORDER BY CONVERT(float, b.KEKKA) DESC
"@ 95

Q '--- 3. BMI・腹囲の判定分布 (性別ごと) ---' @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名,
       g.SEIBETU AS 性別,
       '[' + LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) + ']' AS 判定,
       COUNT(*) AS 人数,
       MIN(CONVERT(float, k.KEKKA)) AS 最小, MAX(CONVERT(float, k.KEKKA)) AS 最大
FROM T_KENSA k
JOIN T_KENSIN s      ON s.PK_SEQ   = k.PK_SEQ
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
  AND LTRIM(RTRIM(k.KOMOKU_CD)) IN ('001215','001217')
  AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) NOT IN ('','#')
  AND ISNUMERIC(k.KEKKA) = 1
GROUP BY LTRIM(RTRIM(k.KOMOKU_CD)), g.SEIBETU, k.HANTEI_KIGO
ORDER BY 1, 3, 4
"@ 40

Q '--- 4. 参考: 判定記号の意味 (T_HANTEIS) ---' @"
SELECT * FROM T_HANTEIS ORDER BY 1
"@ 30

Q '--- 5. 参考: 基準値1 で C を使っている項目 (身体計測だけか) ---' @"
SELECT LTRIM(RTRIM(j.KOMOKU_CD)) AS 項目CD, MAX(m.MEISYO1) AS 項目名, COUNT(*) AS C行数
FROM T_KIJUN1 j
LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(j.KOMOKU_CD))
WHERE j.KIJUN_CD = '1' AND LTRIM(RTRIM(j.HANTEI_KIGO)) = 'C'
GROUP BY LTRIM(RTRIM(j.KOMOKU_CD))
ORDER BY 1
"@ 80

W ''
W '=== 読み方 ==='
W '  1 が本体です。BMI・腹囲の行に D や F があるか、性別で分かれているかを見てください。'
W '     BMI 30.0以上 の行が無ければ、肥満度が高い人でも C 止まりということです。'
W '  2 で 8/21 の実際の値と判定が分かります。BMI の高い順に並んでいます。'
W '  3 で性別ごとの分布が見えます。腹囲が男女で分かれているかの確認に使えます。'
W '  5 で「C を使っているのは身体計測だけなのか」が分かります。'
W '     ほかの項目にも C があれば、C を残すのは意図的な運用ということになります。'
notepad $out
