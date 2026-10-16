<#
  石屋: SRL に便潜血の結果があるのに、健診ナビに便潜血の枠が無い人 (check92.ps1)
  check92.bat をダブルクリックすると実行され、結果 r_check92.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。
  SRL ファイルで便潜血 (金ｺﾛｲﾄﾞ1/2) に値がある受付番号 (103 人) のうち、069245 の枠が無い人を出す。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check92.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 便潜血の結果はあるが枠が無い人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. ★SRL に便潜血の値があるのに、健診ナビに便潜血1 (069245) の枠が無い人 ---' @"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 漢字,
       LTRIM(RTRIM(ISNULL(s.COURSE_CD,''))) AS コース, LTRIM(RTRIM(ISNULL(c.MEISYO,''))) AS コース名, LTRIM(RTRIM(ISNULL(d.MEISYO1,''))) AS 事業所
FROM T_KENSIN s
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
LEFT JOIN T_COURSE1 c ON c.COURSE_CD = s.COURSE_CD AND c.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_DANTAI1 d ON d.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10')
  AND s.UKE_NO_KENSA IN (4631,4632,4633,4634,4635,4640,4643,4645,4647,4648,4653,4654,4655,4658,4660,4662,4663,4664,4666,4668,4669,4671,4672,4673,4674,4675,4677,4678,4679,4681,4682,4683,4686,4687,4688,4693,4698,4699,4700,4702,4703,4704,4705,4706,4707,4708,4709,4710,4717,4718,4719,4720,4721,4723,4726,4727,4730,4732,4733,4735,4736,4737,4738,4739,4740,4749,4750,4751,4752,4753,4757,4758,4763,4769,4772,4773,4779,4781,4782,4784,4788,4792,4793,4794,4795,4799,4801,4802,4803,4804,4805,4806,4811,4812,4814,4815,4817,4818,4823,4824,4828,4830,4831)
  AND NOT EXISTS (SELECT 1 FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '069245')
ORDER BY 1, 2
"@ 30

W ''
W '=== 読み方 ==='
W '  ここに出た人は、便を提出して SRL の結果もあるのに、コースに便潜血が無いので健診ナビに入っていない。'
W '  健診ナビで【便潜血2日法】(OPJ010) を足してから、SRL 53 列版で受付番号を指定して 3→4 すれば入る。0 行なら問題なし。'
notepad $out
