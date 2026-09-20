<#
  コースを新しく作るための下調べ その2 (check62.ps1)
  check61 で列名を間違えたので、実際の列に合わせて取り直す。
  check62.bat をダブルクリックすると実行され、結果 r_check62.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  分かっている構造 (check61 より)
    T_COURSE1 : COURSE_CD, DANTAI_CD1, KIJUN_CD, MEISYO, TYOHYO, IRAI_ID, KENSIN_SYUBETU,
                JIKO_FUTAN_BUN, F_SOTOZEI, F_SYOONIN, F_FAX, F_RIYOUKEN, F_MUKOU,
                MEISYO_RYAKUSYO, KOKYAKU_CD, F_KYOKAI, F_TENKIYOUSI, F_TOKUTEI, F_TOSINKYO,
                F_TOYAKU_SYUBETU, LABEL_MAISUU
    T_COURSE2 : COURSE_CD, DANTAI_CD1, KOMOKU_CD, SET_CD, NARABI     ← コースに入る検査項目
    T_COURSE3 : COURSE_CD, DANTAI_CD1, 健保/団体/個人/代行/その他 の 料金と上限
    T_COURSE4 : COURSE_CD, DANTAI_CD1, 年齢詳細・偶奇数・性別・続柄・部署・雇用形態・有効期間 など
#>
[CmdletBinding()]
param([string]$Name = 'マイクロン')
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check62.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== コースを新しく作るための下調べ2 ($Name) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. T_DANTAI1 の列一覧 ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_DANTAI1') ORDER BY c.column_id
"@ 120

# 列名が分からないので、T_DANTAI1 の文字列の列を全部さがす
Q "--- 2. ★「$Name」を T_DANTAI1 の全テキスト列からさがす ---" @"
DECLARE @w nvarchar(max) = ''
SELECT @w = @w + ' OR CONVERT(nvarchar(4000), [' + c.name + ']) LIKE N''%$Name%'''
FROM sys.columns c JOIN sys.types t ON t.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_DANTAI1') AND t.name IN ('varchar','nvarchar','char','nchar')
EXEC('SELECT TOP 20 * FROM T_DANTAI1 WHERE 1=0' + @w)
"@ 30

Q '--- 3. コースの多い団体 上位20 (どこに雛形があるか) ---' @"
SELECT TOP 20 LTRIM(RTRIM(DANTAI_CD1)) AS 団体CD, COUNT(*) AS コース数
FROM T_COURSE1 GROUP BY DANTAI_CD1 ORDER BY COUNT(*) DESC
"@ 30

Q '--- 4. ★コース名に「ドック」が付くコース (雛形さがし) ---' @"
SELECT TOP 40 LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(c.COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       LTRIM(RTRIM(ISNULL(c.MEISYO_RYAKUSYO,''))) AS 略称,
       LTRIM(RTRIM(ISNULL(c.KIJUN_CD,''))) AS 基準値CD,
       LTRIM(RTRIM(ISNULL(c.TYOHYO,''))) AS 帳票,
       LTRIM(RTRIM(ISNULL(c.F_TOSINKYO,''))) AS 東振協F,
       LTRIM(RTRIM(ISNULL(c.F_MUKOU,''))) AS 無効F,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c
WHERE c.MEISYO LIKE N'%ドック%'
ORDER BY 項目数 DESC
"@ 50

Q '--- 5. ★コース名に「生活習慣」または「簡易」が付くコース (A2/Bの雛形) ---' @"
SELECT TOP 40 LTRIM(RTRIM(c.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(c.COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       LTRIM(RTRIM(ISNULL(c.MEISYO_RYAKUSYO,''))) AS 略称,
       LTRIM(RTRIM(ISNULL(c.KIJUN_CD,''))) AS 基準値CD,
       LTRIM(RTRIM(ISNULL(c.TYOHYO,''))) AS 帳票,
       LTRIM(RTRIM(ISNULL(c.F_TOSINKYO,''))) AS 東振協F,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_COURSE1 c
WHERE c.MEISYO LIKE N'%生活習慣%' OR c.MEISYO LIKE N'%簡易%' OR c.COURSE_CD IN ('A2','B','YA','YB')
ORDER BY c.DANTAI_CD1, c.COURSE_CD
"@ 60

Q '--- 6. 2026/08/21 城西学園で使ったコース (実績のある東振協コース) ---' @"
SELECT DISTINCT LTRIM(RTRIM(s.DANTAI_CD1)) AS 団体CD, LTRIM(RTRIM(s.COURSE_CD)) AS コースCD,
       LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名,
       (SELECT COUNT(*) FROM T_COURSE2 x WHERE x.DANTAI_CD1 = c.DANTAI_CD1 AND x.COURSE_CD = c.COURSE_CD) AS 項目数
FROM T_KENSIN s
LEFT JOIN T_COURSE1 c ON c.DANTAI_CD1 = s.DANTAI_CD1 AND c.COURSE_CD = s.COURSE_CD
WHERE s.D_KENSIN = '2026/08/21' AND s.F_TORIKESI = 0
"@ 30

Q '--- 7. T_KOMOKU の列一覧 (項目名の列を確かめる) ---' @"
SELECT c.column_id AS 順, c.name AS 列, ty.name AS 型, c.max_length AS 長さ
FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID('T_KOMOKU') ORDER BY c.column_id
"@ 120

Q '--- 8. T_COURSE2 の SET_CD の使われ方 (セットでまとめて入れているか) ---' @"
SELECT TOP 30 LTRIM(RTRIM(ISNULL(SET_CD,'(空)'))) AS セットCD, COUNT(*) AS 件数
FROM T_COURSE2 GROUP BY SET_CD ORDER BY COUNT(*) DESC
"@ 40

W ''
W '=== 読み方 ==='
W '  2 でマイクロンが団体マスタに登録済みかが分かる。0件なら団体から作ることになる。'
W '  4・5 が本命。ドックコース / 東振協コースの雛形が既にあるかを見る。'
W '  あった団体CDとコースCDを教えてもらえれば、その中身 (T_COURSE2の項目) を次に出す。'
W '  ※ 読むだけです。コースは作っていません。'
notepad $out
