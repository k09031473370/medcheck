<#
  心電図を「全員撮っているか」を調べる (check7.ps1)
  check7.bat をダブルクリックすると実行され、結果 r_check7.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  目的: リアンのExcelには「心電図を撮ったか」を示す列が無い。
        過去の実績から、枠がある人は全員撮っているのかを確かめる。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check7.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

# リアンは受付番号4000番台。過去1年ぶんを日ごとに集計する
$sql = @"
SELECT s.D_KENSIN AS 受診日,
       COUNT(*) AS 心電図の枠がある人数,
       SUM(CASE WHEN k.KEKKA IS NOT NULL AND LTRIM(RTRIM(k.KEKKA)) <> '' THEN 1 ELSE 0 END) AS 所見が入っている人数,
       SUM(CASE WHEN k.KEKKA IS NULL OR LTRIM(RTRIM(k.KEKKA)) = '' THEN 1 ELSE 0 END) AS 空欄の人数
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A'
WHERE s.F_TORIKESI = 0
  AND s.D_KENSIN >= '2025/08/01'
  AND LTRIM(RTRIM(s.UKE_NO_KENSA)) LIKE '4%'
GROUP BY s.D_KENSIN
ORDER BY s.D_KENSIN DESC
"@

$sql2 = @"
SELECT s.D_KENSIN AS 受診日, s.UKE_NO_KENSA AS 受付番号, k.KEKKA AS 心電図所見
FROM T_KENSIN s
JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '067112A'
WHERE s.F_TORIKESI = 0
  AND s.D_KENSIN >= '2025/08/01'
  AND LTRIM(RTRIM(s.UKE_NO_KENSA)) LIKE '4%'
  AND (k.KEKKA IS NULL OR LTRIM(RTRIM(k.KEKKA)) = '')
ORDER BY s.D_KENSIN DESC, s.UKE_NO_KENSA
"@

"=== 心電図の実施状況 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
"" | Out-File $out -Append -Encoding Default
"--- 1. 受診日ごと: 心電図の枠がある人と、所見が入っている人 ---" | Out-File $out -Append -Encoding Default
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows 80 *>&1 | Out-File $out -Append -Encoding Default
"" | Out-File $out -Append -Encoding Default
"--- 2. 枠はあるが所見が空の人 (この人たちが「撮っていない」のか「入力漏れ」なのか) ---" | Out-File $out -Append -Encoding Default
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql2 -MaxRows 120 *>&1 | Out-File $out -Append -Encoding Default
"" | Out-File $out -Append -Encoding Default
"※ 1で「枠がある人数」と「所見が入っている人数」が毎回同じなら、全員撮っていることになります。" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
