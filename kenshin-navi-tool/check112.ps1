<#
  石屋: くくり判定が D 以上 (D/E/F/G) なのに、医師指示事項 (総合判定N(編集) 11001〜11020) に同じ種類の文が無い人 (check112.ps1)
  文の種類は言葉で見る (例 糖代謝 = 糖代謝|糖尿病|血糖|ＨｂＡ１ｃ)。J (治療中) は見ない。
  check112.bat → r_check112.txt。DBは読むだけ。
#>
[CmdletBinding()]
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check112.txt'
function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''; W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}
"=== 石屋: くくり判定 D 以上なのに医師指示事項に文が無い人 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }
$ISHIYA = "s.F_TORIKESI = 0 AND s.D_KENSIN IN ('2026/09/03','2026/09/09','2026/09/10') AND s.UKE_NO_KENSA BETWEEN 4631 AND 4831"
$SLOTS = "'11001','11002','11003','11004','11005','11006','11007','11008','11009','11010','11011','11012','11013','11014','11015','11016','11017','11018','11019','11020'"
$KINDS = @(
    @{ Cd = '600020'; Name = '身体計測'; Words = @('身体計測','肥満','体重') },
    @{ Cd = '600040'; Name = '視力';     Words = @('視力') },
    @{ Cd = '600060'; Name = '眼底';     Words = @('眼底') },
    @{ Cd = '600070'; Name = '聴力';     Words = @('聴力') },
    @{ Cd = '600090'; Name = '胸部X線';  Words = @('胸部') },
    @{ Cd = '600130'; Name = '心電図';   Words = @('心電図') },
    @{ Cd = '600150'; Name = '血圧';     Words = @('血圧') },
    @{ Cd = '600160'; Name = '消化器';   Words = @('胃') },
    @{ Cd = '600200'; Name = '便潜血';   Words = @('便潜血','大腸') },
    @{ Cd = '600220'; Name = '腎機能';   Words = @('腎機能','ｅＧＦＲ','eGFR','クレアチニン','尿蛋白') },
    @{ Cd = '600310'; Name = '血液一般'; Words = @('血液','貧血','ヘモグロビン','赤血球','白血球','血小板','ヘマトクリット') },
    @{ Cd = '600320'; Name = '脂質代謝'; Words = @('脂質','中性脂肪','コレステロール','ＬＤＬ','ＨＤＬ','LDL','HDL') },
    @{ Cd = '600331'; Name = '肝機能';   Words = @('肝機能','ＡＳＴ','ＡＬＴ','ＧＴＰ','ＡＬＰ','γ','ビリルビン','アルブミン') },
    @{ Cd = '600340'; Name = '糖代謝';   Words = @('糖代謝','糖尿病','血糖','ＨｂＡ１ｃ','HbA1c','尿糖') }   # 健診ナビの糖代謝のくくりには尿糖も入る (4644・4734 は尿糖 精検で糖代謝 F),
    @{ Cd = '600460'; Name = '尿酸';     Words = @('尿酸') }
)
$parts = foreach ($k in $KINDS) {
    $likes = ($k.Words | ForEach-Object { "c.KEKKA LIKE N'%" + $_ + "%'" }) -join ' OR '
@"
SELECT CONVERT(varchar(10), s.D_KENSIN, 111) AS 受診日, s.UKE_NO_KENSA AS 受付番号, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS 氏名,
       N'$($k.Name)' AS くくり, LTRIM(RTRIM(k.HANTEI_KIGO)) AS 判定
FROM T_KENSIN s JOIN T_KENSA k ON k.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(k.KOMOKU_CD)) = '$($k.Cd)'
LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID
WHERE $ISHIYA AND LTRIM(RTRIM(ISNULL(k.HANTEI_KIGO,''))) IN ('D','E','F','G')
  AND NOT EXISTS (SELECT 1 FROM T_KENSA c WHERE c.PK_SEQ = s.PK_SEQ AND LTRIM(RTRIM(c.KOMOKU_CD)) IN ($SLOTS) AND ($likes))
"@
}
Q '--- 1. くくり判定 D 以上で、同じ種類の医師指示事項が無いもの ---' (($parts -join "UNION ALL`r`n") + "ORDER BY 1, 2, 4") 400

W ''
W '=== 読み方 ==='
W '  ここに出た人は、結果票の医師指示事項にその検査の説明が無い。画面で文を足すか、先生に確認する。'
W '  文の種類は言葉で見ているので、言い回しが違うと「無い」と出ることがある (画面で見て、入っていればよい)。'
notepad $out
