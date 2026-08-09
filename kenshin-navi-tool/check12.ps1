<#
  サンテック請求の材料あつめ (check12.ps1)
  check12.bat をダブルクリックすると実行され、結果 r_check12.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  目的: 2026/06/26 実施のサンテック健診を、会社へ一括請求したい。
        請求一覧Excelを作るために、DBに何が入っているかを確かめる。
        - 対象者(氏名・年齢・コース)
        - コースやオプションの料金がマスタに入っているか
        - バリウム(胃部X線)・PSA の実施が結果から分かるか
        - 健診ナビ自身の請求テーブル(T_RYOUKIN/T_SEIKYU2)に金額が入っているか
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check12.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== サンテック請求の材料 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. サンテックの団体登録 ---' @"
SELECT DANTAI_CD1 AS 事業所CD, MEISYO1 AS 事業所名, KENPO_CD AS 健保CD
FROM M_DANTAI
WHERE MEISYO1 LIKE N'%サンテック%' OR KANA_MEISYO1 LIKE N'%サンテック%' OR RYAKUSYO1 LIKE N'%サンテック%'
"@ 10

Q '--- 2. 2026/06/26 の受診者 (全員・団体つき) ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, j.KANJI_SIMEI AS 氏名, s.N_NENREI AS 年齢,
       d.MEISYO1 AS 事業所名, s.COURSE_CD AS コースCD, c.MEISYO AS コース名,
       s.F_TORIKESI AS 取消
FROM T_KENSIN s
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.D_KENSIN = '2026/06/26'
ORDER BY s.UKE_NO_KENSA
"@ 80

Q '--- 3. サンテックのコースと料金 (T_COURSE1/T_COURSE3) ---' @"
SELECT c1.COURSE_CD AS コースCD, c1.MEISYO AS コース名,
       c3.KENPO_RYOUKIN AS 健保料金, c3.DANTAI_RYOUKIN AS 団体料金, c3.KOJIN_RYOUKIN AS 個人料金
FROM T_COURSE1 c1
LEFT JOIN T_COURSE3 c3 ON c3.COURSE_CD = c1.COURSE_CD AND c3.DANTAI_CD1 = c1.DANTAI_CD1
WHERE c1.DANTAI_CD1 IN (SELECT DANTAI_CD1 FROM M_DANTAI WHERE MEISYO1 LIKE N'%サンテック%')
ORDER BY c1.COURSE_CD
"@ 30

Q '--- 4. サンテックのオプション料金 (M_OPTION。PSAや胃部X線の加減算がここにあるか) ---' @"
SELECT o.COURSE_CD AS コースCD, o.KOMOKU_CD AS 項目CD, k.MEISYO1 AS 項目名,
       o.K_ADD_DEL AS 追加削除区分, o.KENPO_RYOUKIN AS 健保料金,
       o.DANTAI_RYOUKIN AS 団体料金, o.KOJIN_RYOUKIN AS 個人料金
FROM M_OPTION o
LEFT JOIN T_KOMOKU k ON LTRIM(RTRIM(k.KOMOKU_CD)) = LTRIM(RTRIM(o.KOMOKU_CD))
WHERE o.DANTAI_CD1 IN (SELECT DANTAI_CD1 FROM M_DANTAI WHERE MEISYO1 LIKE N'%サンテック%')
ORDER BY o.COURSE_CD, o.KOMOKU_CD
"@ 60

Q '--- 5. 胃部X線・PSA の項目コード (実施判定に使う) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1 AS 項目名
FROM T_KOMOKU
WHERE MEISYO1 LIKE N'%PSA%' OR MEISYO1 LIKE N'%前立腺%'
   OR MEISYO1 LIKE N'%胃部%' OR MEISYO1 LIKE N'%バリウム%' OR MEISYO1 LIKE N'%胃検査%'
ORDER BY KOMOKU_CD
"@ 40

Q '--- 6. 6/26の人の 胃・PSA の枠と結果 (誰が実施したか) ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(k.KOMOKU_CD)) AS 項目CD,
       km.MEISYO1 AS 項目名, k.KEKKA AS 結果
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ
LEFT JOIN T_KOMOKU km ON LTRIM(RTRIM(km.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE s.D_KENSIN = '2026/06/26' AND s.F_TORIKESI = 0
  AND (km.MEISYO1 LIKE N'%PSA%' OR km.MEISYO1 LIKE N'%前立腺%' OR km.MEISYO1 LIKE N'%胃%')
ORDER BY s.UKE_NO_KENSA, k.KOMOKU_CD
"@ 120

Q '--- 7. 健診ナビの料金テーブル T_RYOUKIN の列 ---' @"
SELECT COLUMN_NAME AS COL, DATA_TYPE AS TYPE
FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_RYOUKIN'
ORDER BY ORDINAL_POSITION
"@ 60

Q '--- 8. 6/26の人の T_RYOUKIN (金額が既に入っているか) ---' @"
SELECT TOP 40 r.*
FROM T_RYOUKIN r
WHERE r.PK_SEQ IN (SELECT PK_SEQ FROM T_KENSIN WHERE D_KENSIN = '2026/06/26' AND F_TORIKESI = 0)
"@ 45

Q '--- 9. 6/26の人の T_KENSIN の料金・請求まわりの列 ---' @"
SELECT s.UKE_NO_KENSA AS 受付番号, s.COURSE_CD AS コース,
       s.TOKUTEI_GOUKEI AS 特定合計, s.TOKUTEI_KOJIN_FUTAN AS 特定個人負担,
       s.TOKUTEI_HOKENSYA_FUTAN AS 特定保険者負担, s.SIHARA_KBN AS 支払区分,
       s.F_SEIKYUU AS 請求F, s.F_SEIKYUU_K AS 請求FK, s.F_SEIKYUU_D AS 請求FD, s.F_SEIKYUU_S AS 請求FS
FROM T_KENSIN s
WHERE s.D_KENSIN = '2026/06/26' AND s.F_TORIKESI = 0
ORDER BY s.UKE_NO_KENSA
"@ 80

W ''
W '=== 完了。この内容をチャットに貼り付けてください ==='
notepad $out
