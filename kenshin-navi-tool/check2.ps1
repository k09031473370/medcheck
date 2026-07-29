<#
  SRL血液の直接取込に向けた調査 (check2.ps1)
  check2.bat をダブルクリックすると実行され、結果 r_check2.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。個人情報は出力しません(列名とコードのみ)。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check2.txt'

if (-not (Test-Path $tool)) {
    "db_tool.ps1 が同じフォルダにありません: $dir" | Out-File $out -Encoding Default
    notepad $out
    return
}

# 血液Excel_to_DAT.py の COL_TO_CODE (4桁) → DATの8桁 "10"+code+"00"
$codes = @(
    @{ C='1001'; N='総蛋白(TP)' },      @{ C='1002'; N='アルブミン' },
    @{ C='1029'; N='尿素窒素' },        @{ C='1030'; N='クレアチニン' },
    @{ C='1028'; N='尿酸' },            @{ C='1021'; N='総コレステロール' },
    @{ C='1418'; N='LDL' },             @{ C='1023'; N='HDL' },
    @{ C='1024'; N='中性脂肪' },        @{ C='1009'; N='AST' },
    @{ C='1010'; N='ALT' },             @{ C='1159'; N='ALP' },
    @{ C='1013'; N='γ-GTP' },           @{ C='1050'; N='血糖' },
    @{ C='1068'; N='HbA1c' },           @{ C='2001'; N='白血球' },
    @{ C='2002'; N='赤血球' },          @{ C='2003'; N='血色素' },
    @{ C='2004'; N='ヘマトクリット' },  @{ C='2008'; N='血小板' },
    @{ C='1032'; N='Na' },              @{ C='1033'; N='K' },
    @{ C='1034'; N='Cl' },              @{ C='1059'; N='eGFR' },
    @{ C='2005'; N='MCV' },             @{ C='2006'; N='MCH' },
    @{ C='2007'; N='MCHC' },            @{ C='4353'; N='便潜血1' },
    @{ C='4365'; N='便潜血2' }
)
$outCds = ($codes | ForEach-Object { "'10" + $_.C + "00'" }) -join ','

$queries = @(
    @{ Title = '1. 外注コード → 健診ナビの項目コード (江東微研の変換表より)'
       Sql = "SELECT h.OUT_CD, h.KOMOKU_CD, k.MEISYO1 FROM M_HENKAN_IRAI h LEFT JOIN T_KOMOKU k ON k.KOMOKU_CD = h.KOMOKU_CD WHERE h.CENTER_CD = 1 AND h.OUT_CD IN ($outCds) ORDER BY h.OUT_CD"
       Max = 40 }
    @{ Title = '2. 個人マスタ T_KOJIN1 の列一覧 (カルテNoの所在を確認)'
       Sql = "SELECT c.name AS COL, ty.name AS TYPE FROM sys.columns c JOIN sys.types ty ON ty.user_type_id = c.user_type_id WHERE c.object_id = OBJECT_ID('T_KOJIN1') ORDER BY c.column_id"
       Max = 80 }
    @{ Title = '3. カルテNoらしい列を持つテーブル'
       Sql = "SELECT t.name AS TBL, c.name AS COL FROM sys.tables t JOIN sys.columns c ON c.object_id = t.object_id WHERE c.name LIKE '%KARUTE%' OR c.name LIKE '%KARTE%' OR c.name LIKE '%CARTE%' ORDER BY t.name"
       Max = 40 }
)

"=== SRL血液 取込調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
foreach ($q in $queries) {
    "" | Out-File $out -Append -Encoding Default
    "--- $($q.Title) ---" | Out-File $out -Append -Encoding Default
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $q.Sql -MaxRows $q.Max *>&1 |
        Out-File $out -Append -Encoding Default
}

"" | Out-File $out -Append -Encoding Default
"=== 完了。この内容をチャットに貼り付けてください ===" | Out-File $out -Append -Encoding Default
notepad $out
