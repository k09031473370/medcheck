<#
  社員番号の置き場所を決める・その2 (check9.ps1)
  check9.bat をダブルクリックすると実行され、結果 r_check9.txt がメモ帳で開きます。
  DBは読むだけ、帳票テンプレートも読むだけで、一切変更しません。

  check8 で分かったこと:
    空いている欄  … 個人1.TECHO_NO / 個人2.SYUSSEKI_NO / 個人2.DIVISION /
                     受診.FREE_CODE / 受診.BIKO
    使われている欄 … 個人1.KOJIN_NO(81件) / 受診.BUSYO_CD(158件) / 受診.OCR_CODE(4293件)

  ここで知りたいこと:
    1) 個人1.KOJIN_NO は「社員番号を入れる欄」として健診ナビが用意したものか
       (総人数に対して81件なら、正しい欄が一部でだけ使われている可能性が高い)
    2) 事業所名・部署名はもうDBに入っているか
    3) 結果票のテンプレートにどんな欄があるか
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check9.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 社員番号の置き場所 調査2 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default

if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 個人1.KOJIN_NO は本来の社員番号欄か (総人数と、入っている人数) ---' @"
SELECT COUNT(*) AS 個人マスタの総人数,
       SUM(CASE WHEN LTRIM(RTRIM(ISNULL(KOJIN_NO,''))) <> '' THEN 1 ELSE 0 END) AS KOJIN_NOが入っている人数
FROM T_KOJIN1
"@ 10

Q '--- 2. KOJIN_NO が入っている人の中身 (社員番号らしい値か) ---' @"
SELECT TOP 25 j.KOJIN_NO AS 個人番号, j.KANJI_SIMEI AS 氏名,
       d.MEISYO1 AS 事業所名
FROM T_KOJIN1 j
LEFT JOIN (SELECT DISTINCT KOJIN_ID, DANTAI_CD1 FROM T_KENSIN) s ON s.KOJIN_ID = j.KOJIN_ID
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE LTRIM(RTRIM(ISNULL(j.KOJIN_NO,''))) <> ''
ORDER BY j.KOJIN_NO
"@ 40

Q '--- 3. 7/12 リアンの人に今なにが入っているか ---' @"
SELECT TOP 12 s.UKE_NO_KENSA AS 受付番号, j.KANJI_SIMEI AS 氏名,
       s.DANTAI_CD1 AS 事業所CD, d.MEISYO1 AS 事業所名,
       s.BUSYO_CD AS 部署CD, b.BUSYO_MEI AS 部署名,
       j.KOJIN_NO AS 個人番号, j.KARUTE_NO AS カルテNo, j.TECHO_NO AS 手帳No,
       j2.SYUSSEKI_NO AS 出席番号, s.FREE_CODE AS フリーコード
FROM T_KENSIN s
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_KOJIN2 j2 ON j2.KOJIN_ID = s.KOJIN_ID
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_BUSYO b ON b.DANTAI_CD1 = s.DANTAI_CD1 AND b.BUSYO_CD = s.BUSYO_CD
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN = '2026/07/12'
ORDER BY s.UKE_NO_KENSA
"@ 20

Q '--- 4. 部署マスタ T_BUSYO の中身 (リアンの事業所の部署が登録されているか) ---' @"
SELECT TOP 30 b.DANTAI_CD1 AS 事業所CD, d.MEISYO1 AS 事業所名,
       b.BUSYO_CD AS 部署CD, b.BUSYO_MEI AS 部署名
FROM T_BUSYO b LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = b.DANTAI_CD1
ORDER BY b.DANTAI_CD1, b.BUSYO_CD
"@ 50

Q '--- 5. リアンの事業所が団体マスタにどう登録されているか ---' @"
SELECT DANTAI_CD1 AS 事業所CD, MEISYO1 AS 事業所名, RYAKUSYO1 AS 略称
FROM M_DANTAI
WHERE MEISYO1 LIKE N'%木下%' OR RYAKUSYO1 LIKE N'%木下%'
"@ 20

# ---- 帳票テンプレートに何の欄があるか (読むだけ) ----
W ''
W '--- 6. 結果票テンプレートの中の文字 (どの欄が印字できるか) ---'
$reader = Join-Path $dir 'xlsx_read.ps1'
if (-not (Test-Path $reader)) {
    W 'xlsx_read.ps1 が同じフォルダにありません'
} else {
    . $reader
    $roots = @('\\KNSV\KenshinNavi', 'C:\KenshinNavi', 'D:\KenshinNavi')
    $tpl = $null
    foreach ($r in $roots) {
        if (-not (Test-Path $r)) { continue }
        $hit = Get-ChildItem -Path $r -Filter '*個人結果票*.xls*' -Recurse -ErrorAction SilentlyContinue |
               Select-Object -First 1
        if ($hit) { $tpl = $hit.FullName; break }
    }
    if (-not $tpl) {
        W '個人結果票のテンプレートが見つかりませんでした。'
        W '見つかる場所が分かれば、そのパスを教えてください。'
    } else {
        W "テンプレート: $tpl"
        W ''
        try {
            $rows = Read-Xlsx $tpl
            $n = 0
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $line = @()
                for ($c = 0; $c -lt $rows[$i].Count; $c++) {
                    $v = [string]$rows[$i][$c]
                    if ($v -ne $null -and $v.Trim() -ne '') { $line += ('{0}:{1}' -f ($c + 1), $v.Trim()) }
                }
                if ($line.Count -gt 0) {
                    W ('{0,4}行  {1}' -f ($i + 1), ($line -join ' | '))
                    $n++
                    if ($n -ge 120) { W '(以下省略)'; break }
                }
            }
        } catch {
            W ("読めませんでした: {0}" -f $_.Exception.Message)
        }
    }
}

W ''
W '=== 完了。この内容をチャットに貼り付けてください ==='
notepad $out
