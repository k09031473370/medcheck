<#
  帳票が項目を探すときの「名前」がどの列かを確かめる (check52.ps1)
  check52.bat をダブルクリックすると実行され、結果 r_check52.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  帳票303の差込は **[GOT]結果値_1 / **[HbA1cNGSP]結果値_1 / **[血色素量]結果値_1 のように書いてある。
  項目名(MEISYO1)は AST(GOT) / HbA1c(NGSP) / ﾍﾓｸﾞﾛﾋﾞﾝ なので一致しない。
  T_KOMOKU の右のほうの列 (VR_MEISYO1 / MEI_HOKOKU / MEI_SEIKYU) のどれかが帳票用の名前のはず。
  これが分かれば、血糖(未判別)・中性脂肪(未判別) を帳票303に書き足すときの正確な文字が決まる。
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check52.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 帳票の項目名の列を確かめる $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★帳票に出ている項目 (差込の名前が分かっているもの) の名前の列たち ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD,
       '[' + ISNULL(MEISYO1,'') + ']'    AS MEISYO1,
       '[' + ISNULL(MEISYO2,'') + ']'    AS MEISYO2,
       '[' + ISNULL(VR_MEISYO1,'') + ']' AS VR_MEISYO1,
       '[' + ISNULL(MEI_HOKOKU,'') + ']' AS MEI_HOKOKU,
       '[' + ISNULL(MEI_SEIKYU,'') + ']' AS MEI_SEIKYU
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN
  ('068636','068637','068689','069005','069006','068705','068708','068709','068651','068645','068674','001211','067012','067018')
ORDER BY KOMOKU_CD
"@ 20
W '   帳票の差込: 068636=[GOT] 068637=[GPT] 068689=[HbA1cNGSP] 069005=[血色素量] 069006=[ﾍﾏﾄｸﾘｯﾄ] 068705=[総ｺﾚｽﾃﾛｰﾙ]'
W '              068708=[HDLｺﾚｽﾃﾛｰﾙ] 068709=[LDLｺﾚｽﾃﾛｰﾙ] 068651=[LAP] 068645=[ALP] 068674=[ｸﾚｱﾁﾆﾝ] 001211=[身長] 067012=[視力右] 067018=[矯正右]'

Q '--- 2. ★血糖・中性脂肪まわり。同じ列たち ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD,
       '[' + ISNULL(MEISYO1,'') + ']'    AS MEISYO1,
       '[' + ISNULL(VR_MEISYO1,'') + ']' AS VR_MEISYO1,
       '[' + ISNULL(MEI_HOKOKU,'') + ']' AS MEI_HOKOKU,
       '[' + ISNULL(MEI_SEIKYU,'') + ']' AS MEI_SEIKYU
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068681','068682','H30-0011','H30-0012','H30-0013','068700','068702','R06-0001','R06-0002','R06-0003')
ORDER BY KOMOKU_CD
"@ 20

Q '--- 3. ★「(未判別)」のカッコが全角か半角か (文字コード。40=半角( 65288=全角（) ---' @"
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS 項目CD, MEISYO1,
       UNICODE(SUBSTRING(MEISYO1, CHARINDEX(N'未', MEISYO1) - 1, 1)) AS 開きカッコ,
       UNICODE(SUBSTRING(MEISYO1, CHARINDEX(N'別', MEISYO1) + 1, 1)) AS 閉じカッコ,
       LEN(MEISYO1) AS 文字数,
       UNICODE(SUBSTRING(ISNULL(MEI_HOKOKU,''), CHARINDEX(N'未', ISNULL(MEI_HOKOKU,'')) - 1, 1)) AS 報告名の開きカッコ
FROM T_KOMOKU
WHERE LTRIM(RTRIM(KOMOKU_CD)) IN ('068682','R06-0002')
"@ 10

Q '--- 4. 帳票の定義を持つ表があるか (CHOHYO / HOKOKU / REPORT / SASHIKOMI) ---' @"
SELECT t.TABLE_NAME AS 表, COUNT(*) AS 列数
FROM INFORMATION_SCHEMA.TABLES t JOIN INFORMATION_SCHEMA.COLUMNS c ON c.TABLE_NAME = t.TABLE_NAME
WHERE t.TABLE_TYPE = 'BASE TABLE'
  AND (t.TABLE_NAME LIKE '%CHOHYO%' OR t.TABLE_NAME LIKE '%CHOUHYOU%' OR t.TABLE_NAME LIKE '%HOKOKU%'
       OR t.TABLE_NAME LIKE '%REPORT%' OR t.TABLE_NAME LIKE '%SASHI%' OR t.TABLE_NAME LIKE '%EXCEL%')
GROUP BY t.TABLE_NAME ORDER BY 1
"@ 30

W ''
W '=== 読み方 ==='
W '  1 で、[GOT] [HbA1cNGSP] [血色素量] と同じ文字が入っている列が「帳票用の名前」です。'
W '  2 の同じ列に、血糖(未判別)・中性脂肪(未判別) がどう書かれているかを見ます。'
W '     その文字をそのまま帳票303に書けば出るようになります。'
W '  3 はカッコの全角/半角です。40 なら半角 ( 、65288 なら全角 （ です。'
notepad $out
