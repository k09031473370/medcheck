<#
  帳票が使う名前の一覧 (T_HOKOKU など) を見る (check53.ps1)
  check53.bat をダブルクリックすると実行され、結果 r_check53.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  帳票303の **[空腹時血糖] を **[血糖] に、**[中性脂肪] を **[中性脂肪未判別] に書き換えたら
  健診ナビのプレビューが 0x800A03EC で落ちた。LAP のときは同じやり方で通った。
  健診ナビが帳票に使える名前を T_HOKOKU のような表で持っているなら、そこに無い名前で落ちる。
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check53.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 帳票が使う名前の一覧 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_HOKOKU の列 ---' @"
SELECT c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型, c.CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = 'T_HOKOKU' ORDER BY c.ORDINAL_POSITION
"@ 10

Q '--- 2. T_HOKOKU の行数と先頭30行 ---' @"
SELECT TOP 30 * FROM T_HOKOKU
"@ 35
Q '--- 2b. 行数 ---' 'SELECT COUNT(*) AS 行数 FROM T_HOKOKU' 5

Q '--- 3. ★VR_MEISYO1 が 血糖 / 中性脂肪未判別 / LAP / 中性脂肪 / 空腹時血糖 の項目 (重複していないか) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, '[' + ISNULL(VR_MEISYO1,'') + ']' AS VR_MEISYO1, MEISYO1,
       '[' + ISNULL(TAN_I,'') + ']' AS 単位, SYOUSUU AS 小数, S_1, S_2, S_3
FROM T_KOMOKU
WHERE LTRIM(RTRIM(ISNULL(VR_MEISYO1,''))) IN (N'血糖', N'中性脂肪未判別', N'LAP', N'中性脂肪', N'空腹時血糖', N'随時血糖', N'随時中性脂肪')
ORDER BY VR_MEISYO1, KOMOKU_CD
"@ 20

Q '--- 4. ★基準値マスタ: 血糖(068682) / 中性脂肪(R06-0002) / LAP(068651) / 空腹時血糖(H30-0011) ---' @"
SELECT KOMOKU_CD, KIJUN_CD, RENBAN, SEIBETU, HANTEI_KIGO, TEISEI, HL, HYOJI_YO, KIJUN_MEISYO, HANTEI_CD
FROM T_KIJUN2
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068682','R06-0002','068651','H30-0011','068702')
ORDER BY KOMOKU_CD, KIJUN_CD, RENBAN
"@ 80

Q '--- 5. 帳票の設定らしき表 (列に HOKOKU / CHOHYO / EXCEL / FILE を含む表) ---' @"
SELECT c.TABLE_NAME AS 表, c.COLUMN_NAME AS 列
FROM INFORMATION_SCHEMA.COLUMNS c JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME
WHERE t.TABLE_TYPE = 'BASE TABLE'
  AND (c.COLUMN_NAME LIKE '%HOKOKU%' OR c.COLUMN_NAME LIKE '%CHOHYO%' OR c.COLUMN_NAME LIKE '%EXCEL%' OR c.COLUMN_NAME LIKE '%XLS%')
ORDER BY 1, 2
"@ 40

W ''
W '=== 読み方 ==='
W '  2 に 帳票の名前 (GOT / 血色素量 / LAP …) の一覧が出ていれば、健診ナビはその表で名前を解決しています。'
W '     血糖・中性脂肪未判別 がそこに無ければ、それが落ちた理由です。'
W '  3 で同じ名前の項目が2つ以上あれば、それも落ちる理由になります。'
W '  4 は 基単 (基準値) を作るときに読む表です。血糖・中性脂肪の行の形が LAP と違わないかを見ます。'
notepad $out
