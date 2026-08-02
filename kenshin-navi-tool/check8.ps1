<#
  社員番号・事業所名の置き場所を決めるための調査 (check8.ps1)
  check8.bat をダブルクリックすると実行され、結果 r_check8.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  目的: リアンのExcelには 1列目=社員番号 / 5列目=団体名 / 13列目=部署名 が入っている。
        これを健診ナビのどの欄に入れれば結果票に出せるかを決めたい。
        「今すでに健診ナビが使っている欄」を潰さないよう、空いている欄を探す。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check8.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

# 1. 候補になる欄が、健診ナビで今どれくらい使われているか
#    (件数が0なら空いている。多ければ健診ナビが使っているので触らない)
$sql1 = @"
SELECT '個人1.KOJIN_NO'   AS 欄, COUNT(*) AS 入っている件数 FROM T_KOJIN1 WHERE LTRIM(RTRIM(ISNULL(KOJIN_NO,'')))   <> ''
UNION ALL SELECT '個人1.KARUTE_NO',    COUNT(*) FROM T_KOJIN1 WHERE LTRIM(RTRIM(ISNULL(KARUTE_NO,'')))    <> ''
UNION ALL SELECT '個人1.TECHO_NO',     COUNT(*) FROM T_KOJIN1 WHERE LTRIM(RTRIM(ISNULL(TECHO_NO,'')))     <> ''
UNION ALL SELECT '個人2.SYUSSEKI_NO',  COUNT(*) FROM T_KOJIN2 WHERE LTRIM(RTRIM(ISNULL(SYUSSEKI_NO,'')))  <> ''
UNION ALL SELECT '個人2.DIVISION',     COUNT(*) FROM T_KOJIN2 WHERE LTRIM(RTRIM(ISNULL(DIVISION,'')))     <> ''
UNION ALL SELECT '個人2.GYOSYU',       COUNT(*) FROM T_KOJIN2 WHERE LTRIM(RTRIM(ISNULL(GYOSYU,'')))       <> ''
UNION ALL SELECT '個人2.CLASS',        COUNT(*) FROM T_KOJIN2 WHERE LTRIM(RTRIM(ISNULL(CLASS,'')))        <> ''
UNION ALL SELECT '個人2.GAKUNEN',      COUNT(*) FROM T_KOJIN2 WHERE LTRIM(RTRIM(ISNULL(GAKUNEN,'')))      <> ''
UNION ALL SELECT '個人2.S_UKE_NO',     COUNT(*) FROM T_KOJIN2 WHERE LTRIM(RTRIM(ISNULL(S_UKE_NO,'')))     <> ''
UNION ALL SELECT '受診.FREE_CODE',     COUNT(*) FROM T_KENSIN WHERE LTRIM(RTRIM(ISNULL(FREE_CODE,'')))    <> ''
UNION ALL SELECT '受診.BUSYO_CD',      COUNT(*) FROM T_KENSIN WHERE LTRIM(RTRIM(ISNULL(BUSYO_CD,'')))     <> ''
UNION ALL SELECT '受診.S_UKE_NO',      COUNT(*) FROM T_KENSIN WHERE LTRIM(RTRIM(ISNULL(S_UKE_NO,'')))     <> ''
UNION ALL SELECT '受診.JUSIN_KEN_NO',  COUNT(*) FROM T_KENSIN WHERE LTRIM(RTRIM(ISNULL(JUSIN_KEN_NO,''))) <> ''
UNION ALL SELECT '受診.OCR_CODE',      COUNT(*) FROM T_KENSIN WHERE LTRIM(RTRIM(ISNULL(OCR_CODE,'')))     <> ''
UNION ALL SELECT '受診.BIKO',          COUNT(*) FROM T_KENSIN WHERE LTRIM(RTRIM(ISNULL(BIKO,'')))         <> ''
"@

# 2. 7/12のリアンの人で、候補の欄に今なにが入っているか
$sql2 = @"
SELECT TOP 12 s.UKE_NO_KENSA AS 受付番号, j.NAME_KANJI AS 氏名,
       s.DANTAI_CD1 AS 事業所CD, s.COURSE_CD AS コース,
       j.KOJIN_NO AS 個人1_KOJIN_NO, j.KARUTE_NO AS 個人1_KARUTE_NO, j.TECHO_NO AS 個人1_TECHO_NO,
       j2.SYUSSEKI_NO AS 個人2_SYUSSEKI_NO, j2.DIVISION AS 個人2_DIVISION, j2.GYOSYU AS 個人2_GYOSYU,
       s.FREE_CODE AS 受診_FREE_CODE, s.BUSYO_CD AS 受診_BUSYO_CD
FROM T_KENSIN s
JOIN T_KOJIN1 j  ON j.KOJIN_ID  = s.KOJIN_ID
LEFT JOIN T_KOJIN2 j2 ON j2.KOJIN_ID = s.KOJIN_ID
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '2026/07/12'
ORDER BY s.UKE_NO_KENSA
"@

# 3. 事業所(団体)の名前がどのテーブルに入っているか
$sql3 = @"
SELECT TOP 40 t.TABLE_NAME AS TBL, c.COLUMN_NAME AS COL
FROM INFORMATION_SCHEMA.TABLES t
JOIN INFORMATION_SCHEMA.COLUMNS c ON c.TABLE_NAME = t.TABLE_NAME
WHERE t.TABLE_NAME LIKE '%DANTAI%'
ORDER BY t.TABLE_NAME, c.ORDINAL_POSITION
"@

# 4. 部署マスタらしいものがあるか
$sql4 = @"
SELECT TOP 40 t.TABLE_NAME AS TBL, c.COLUMN_NAME AS COL
FROM INFORMATION_SCHEMA.TABLES t
JOIN INFORMATION_SCHEMA.COLUMNS c ON c.TABLE_NAME = t.TABLE_NAME
WHERE t.TABLE_NAME LIKE '%BUSYO%' OR t.TABLE_NAME LIKE '%BUSHO%'
ORDER BY t.TABLE_NAME, c.ORDINAL_POSITION
"@

"=== 社員番号・事業所名の置き場所 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
"" | Out-File $out -Append -Encoding Default
"--- 1. 候補の欄が今どれくらい使われているか (0件なら空いている) ---" | Out-File $out -Append -Encoding Default
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql1 -MaxRows 40 *>&1 | Out-File $out -Append -Encoding Default
"" | Out-File $out -Append -Encoding Default
"--- 2. 7/12 リアンの人に今なにが入っているか ---" | Out-File $out -Append -Encoding Default
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql2 -MaxRows 20 *>&1 | Out-File $out -Append -Encoding Default
"" | Out-File $out -Append -Encoding Default
"--- 3. 事業所(団体)マスタの列 ---" | Out-File $out -Append -Encoding Default
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql3 -MaxRows 60 *>&1 | Out-File $out -Append -Encoding Default
"" | Out-File $out -Append -Encoding Default
"--- 4. 部署マスタらしいテーブル ---" | Out-File $out -Append -Encoding Default
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql4 -MaxRows 60 *>&1 | Out-File $out -Append -Encoding Default
"" | Out-File $out -Append -Encoding Default
"※ 1で件数が0の欄が「空いている欄」です。そこへ社員番号を入れれば健診ナビの動きを邪魔しません。" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
