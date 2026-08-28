<#
  結果票のクリニック名がどこから来ているか調べる (check15.ps1)
  check15.bat をダブルクリックすると実行され、結果 r_check15.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  やりたいこと: 別のクリニック名で結果票を出したい。
  そのためには「いまの名前がどこに書いてあるか」を先に特定する必要がある。

  考えられる置き場所は3つ:
    A. DBの施設マスタ  … 1か所直せば全部変わる (ただし全部変わってしまう)
    B. 帳票テンプレート … Excelに直接書いてある。テンプレートを複製すれば名前違いを作れる
    C. 団体・コースごと … 契約先ごとに切り替える仕組みが元からある (これが一番良い)

  どれなのかで、やり方も手間も変わる。
#>
$ErrorActionPreference = 'Continue'
$dir  = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out  = Join-Path $dir 'r_check15.txt'

# 探す名前。違うクリニックのときはここを書き換える
$KEY = '芝浦'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== 結果票のクリニック名の出どころ 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W "探す文字: $KEY"
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. 施設・医療機関らしいテーブル ---' @"
SELECT TABLE_NAME AS テーブル,
       (SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS c WHERE c.TABLE_NAME = t.TABLE_NAME) AS 列数
FROM INFORMATION_SCHEMA.TABLES t
WHERE TABLE_NAME LIKE '%SISETU%' OR TABLE_NAME LIKE '%SISET%'
   OR TABLE_NAME LIKE '%KIKAN%'  OR TABLE_NAME LIKE '%IIN%'
   OR TABLE_NAME LIKE '%CLINIC%' OR TABLE_NAME LIKE '%HOSP%'
   OR TABLE_NAME LIKE '%JIGYOSYA%' OR TABLE_NAME LIKE '%KENSIN_KIKAN%'
   OR TABLE_NAME LIKE 'M_SYS%'  OR TABLE_NAME LIKE '%KANRI%'
   OR TABLE_NAME LIKE '%SETTEI%' OR TABLE_NAME LIKE '%CONFIG%'
ORDER BY TABLE_NAME
"@ 60

Q '--- 2. 「名称」「名前」らしい列を持つマスタ (M_ / T_ で始まる小さい表) ---' @"
SELECT c.TABLE_NAME AS テーブル, c.COLUMN_NAME AS 列, c.DATA_TYPE AS 型,
       c.CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS c
JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_NAME = c.TABLE_NAME AND t.TABLE_TYPE = 'BASE TABLE'
WHERE (c.COLUMN_NAME LIKE '%MEISYO%' OR c.COLUMN_NAME LIKE '%NAME%' OR c.COLUMN_NAME LIKE '%SIMEI%')
  AND (c.TABLE_NAME LIKE '%SISETU%' OR c.TABLE_NAME LIKE '%KIKAN%'
       OR c.TABLE_NAME LIKE 'M_SYS%' OR c.TABLE_NAME LIKE '%KANRI%' OR c.TABLE_NAME LIKE '%SETTEI%')
ORDER BY c.TABLE_NAME, c.ORDINAL_POSITION
"@ 80

# 3. 全テーブルの文字列列から KEY を探す。
#    どの表に自院名が入っているかが分かれば、そこが差し替え先の候補になる。
W ''
W "--- 3. 「$KEY」を含む値があるテーブル・列を総当たりで探す ---"
W '    (時間がかかります。見つかった場所が名前の置き場所の候補です)'
& powershell -NoProfile -ExecutionPolicy Bypass -File $tool -FindText $KEY *>&1 | Out-File $out -Append -Encoding Default

Q '--- 4. 団体ごとに出力先や帳票が切り替わる仕組みがあるか (団体マスタの列一覧) ---' @"
SELECT COLUMN_NAME AS 列, DATA_TYPE AS 型, CHARACTER_MAXIMUM_LENGTH AS 長さ
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'M_DANTAI'
ORDER BY ORDINAL_POSITION
"@ 200

W ''
W '--- 5. 帳票テンプレートの置き場所と、クリニック名が書いてあるファイル ---'
$roots = @('\\KNSV\KenshinNavi', 'C:\KenshinNavi')
$found = $false
foreach ($r in $roots) {
    if (-not (Test-Path $r)) { W "  (無い) $r"; continue }
    $found = $true
    W "  調べる場所: $r"
    $tpl = Get-ChildItem -Path $r -Include '*.xls','*.xlsx','*.xlsm' -Recurse -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -match '結果票|報告書|通知書' } |
           Select-Object -First 40
    if (-not $tpl) { W '  結果票らしいテンプレートが見つかりませんでした。'; continue }
    foreach ($f in $tpl) {
        W ("    {0}  ({1:N0} バイト, {2})" -f $f.FullName, $f.Length, $f.LastWriteTime.ToString('yyyy/MM/dd'))
    }

    # xlsx は zip なので、中の文字列を直接のぞける (Excelを開かずに済む)
    W ''
    W "    ↓ この中で「$KEY」という文字が入っているテンプレート"
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    foreach ($f in $tpl) {
        if ($f.Extension -eq '.xls') { continue }   # 旧形式は中身を覗けない
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($f.FullName)
            $hit = $false
            foreach ($e in $zip.Entries) {
                if ($e.FullName -notmatch '\.xml$') { continue }
                $sr = New-Object System.IO.StreamReader($e.Open())
                $txt = $sr.ReadToEnd(); $sr.Close()
                if ($txt -like "*$KEY*") { $hit = $true; break }
            }
            $zip.Dispose()
            if ($hit) { W ("      ★ {0}" -f $f.FullName) }
        }
        catch { W ("      (読めません) {0} : {1}" -f $f.Name, $_.Exception.Message) }
    }
}
if (-not $found) {
    W '  健診ナビのフォルダが見えません。'
    W '  自宅PCなど、サーバーに繋がっていない場所で実行した場合はこれで正常です。'
    W '  この部分だけは、クリニックのPCで実行してください。'
}

W ''
W '=== 読み方 ==='
W '  3 で M_SISETU / M_SYSTEM のようなマスタに名前が見つかった'
W '      → DBを直せば変えられる。ただし全部の帳票が変わるので、契約先ごとに出し分けるなら不可。'
W '  5 で ★ が付いたテンプレートがある'
W '      → 名前はExcelに直接書いてある。テンプレートを複製して名前だけ変えるのが安全。'
W '  4 の団体マスタに施設コードらしい列がある'
W '      → 元から契約先ごとに切り替えられる可能性がある。これが一番きれい。'

notepad $out
