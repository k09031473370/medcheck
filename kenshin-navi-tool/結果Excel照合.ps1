<#
.SYNOPSIS
  結果報告書のExcel と 東振協CSV を1人ずつ突き合わせる (結果Excel照合.ps1)

.DESCRIPTION
  健診ナビから出した結果報告書 (帳票303・1人1ファイルの .xlsx) の中身と、
  取り込みの元になった東振協のCSVを、人ごと・項目ごとに比べる。
  DBには接続しない。Excelも起動しない (xlsxを直接読む)。

  比べ方
    ・帳票303のひな形で「どのセルにどの項目が入るか」を form\帳票303_セル対応.csv に持っている
    ・CSVの列 ↔ 帳票の項目 の対応は、このファイルの $RULES に書いてある
    ・数値は数として比べる (5.80 と 5.8 は同じ)
    ・尿・便潜血は (-)(+-)(+)… にそろえてから比べる
    ・聴力は 所見なし/所見あり にそろえてから比べる
    ・胸部/胃部/心電図の判定は 東振協 C→D, D→F に読み替えてから比べる
    ・所見 (胸部X線・胃部・心電図・診察) はCSVがコードなので、
      合否は付けず「目視」として CSVのコードと帳票の文字を並べて出す
    ・自覚症状・既往歴は、CSVの印の付いた見出しを前から4つ取って比べる

  出るもの
    ・照合結果_<受診日>.txt … 人ごとの 一致/不一致/目視 の件数と、不一致の一覧 (メモ帳で開く)
    ・照合結果_<受診日>.csv … 比べた全項目 (Excelで開いて絞り込める)

.EXAMPLE
  結果Excel照合.bat に、結果報告書の入ったフォルダをドラッグ＆ドロップ → CSVを聞かれるので入れる

  powershell -ExecutionPolicy Bypass -File 結果Excel照合.ps1 -Folder C:\...\結果 -Csv C:\...\20260821.csv
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Folder,            # 結果報告書(.xlsx)の入ったフォルダ (1ファイルでも可)
    [Parameter(Position = 1)]
    [string]$Csv,               # 東振協のCSV
    [string]$CellMap            # 省略時は form\帳票303_セル対応.csv
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression.FileSystem
try { [void][System.Text.Encoding]::GetEncoding(932) }
catch { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) }

if (-not $CellMap) { $CellMap = Join-Path $PSScriptRoot 'form\帳票303_セル対応.csv' }

# ============================================================================
# CSVの列 ↔ 帳票の項目 (略称)
#   Col  … CSVの列番号 (1始まり)
#   Ryaku… 帳票のひな形に書いてある **[略称]結果値_1 の略称。複数書くと、値の入っている方を使う
#   Kind … NUM=数値 / NYOU=尿・便 / CHO=聴力 / HAN=判定(東振協→健診ナビ) / TXT=文字 / CODE=所見コード(目視)
# ============================================================================
$RULES = @(
    @{ Name='身長';            Col=11;  Ryaku=@('身長');                     Kind='NUM' }
    @{ Name='体重';            Col=12;  Ryaku=@('体重');                     Kind='NUM' }
    @{ Name='腹囲';            Col=14;  Ryaku=@('腹囲');                     Kind='NUM' }
    @{ Name='右視力 裸眼';     Col=16;  Ryaku=@('視力右');                   Kind='NUM' }
    @{ Name='左視力 裸眼';     Col=17;  Ryaku=@('視力左');                   Kind='NUM' }
    @{ Name='右視力 矯正';     Col=18;  Ryaku=@('矯正右');                   Kind='NUM' }
    @{ Name='左視力 矯正';     Col=19;  Ryaku=@('矯正左');                   Kind='NUM' }
    # 帳票の「最高血圧その他」は、2回測った人は 1回目と2回目の低い方 (最高・最低それぞれ) が出る (健診ナビの作り)。
    @{ Name='最高血圧';        Col=22;  Col2=24; Ryaku=@('最高血圧その他');  Kind='NUM' }
    @{ Name='最低血圧';        Col=23;  Col2=25; Ryaku=@('最低血圧その他');  Kind='NUM' }
    @{ Name='尿糖';            Col=27;  Ryaku=@('尿糖');                     Kind='NYOU' }
    @{ Name='尿蛋白';          Col=28;  Ryaku=@('尿蛋白');                   Kind='NYOU' }
    @{ Name='尿潜血';          Col=29;  Ryaku=@('尿潜血');                   Kind='NYOU' }
    @{ Name='聴力右1000';      Col=35;  Ryaku=@('聴力右1000');               Kind='CHO' }
    @{ Name='聴力左1000';      Col=36;  Ryaku=@('聴力左1000');               Kind='CHO' }
    @{ Name='聴力右4000';      Col=37;  Ryaku=@('聴力右4000');               Kind='CHO' }
    @{ Name='聴力左4000';      Col=38;  Ryaku=@('聴力左4000');               Kind='CHO' }
    @{ Name='胸部所見1';       Col=44;  Ryaku=@('胸部X線1');                 Kind='CODE' }
    @{ Name='胸部所見2';       Col=48;  Ryaku=@('胸部X線2');                 Kind='CODE' }
    @{ Name='胸部判定';        Col=54;  Ryaku=@('判定_胸部X線');             Kind='HAN' }
    @{ Name='胃部所見1';       Col=62;  Ryaku=@('胃検査所見1');              Kind='CODE' }
    @{ Name='胃部所見2';       Col=66;  Ryaku=@('胃検査所見2');              Kind='CODE' }
    @{ Name='胃部判定';        Col=68;  Ryaku=@('判定_消化器');              Kind='HAN' }
    @{ Name='心電図所見1';     Col=71;  Ryaku=@('心電図1');                  Kind='CODE' }
    @{ Name='心電図所見2';     Col=72;  Ryaku=@('心電図2');                  Kind='CODE' }
    @{ Name='心電図所見3';     Col=73;  Ryaku=@();                           Kind='CODE' }   # 帳票に枠が無い
    @{ Name='心電図判定';      Col=74;  Ryaku=@('判定_心電図');              Kind='HAN' }
    @{ Name='便潜血1';         Col=77;  Ryaku=@('便潜血1');                  Kind='NYOU' }
    @{ Name='便潜血2';         Col=78;  Ryaku=@('便潜血2');                  Kind='NYOU' }
    @{ Name='他覚所見1';       Col=80;  Ryaku=@('内科診察1');                Kind='CODE' }
    @{ Name='他覚所見2';       Col=81;  Ryaku=@('内科診察2');                Kind='CODE' }
    @{ Name='他覚所見3';       Col=82;  Ryaku=@('内科診察3');                Kind='CODE' }
    @{ Name='白血球数';        Col=86;  Ryaku=@('白血球数');                 Kind='NUM' }
    @{ Name='赤血球数';        Col=88;  Ryaku=@('赤血球数');                 Kind='NUM' }
    @{ Name='ヘモグロビン';    Col=90;  Ryaku=@('血色素量');                 Kind='NUM' }
    @{ Name='ヘマトクリット';  Col=92;  Ryaku=@('ﾍﾏﾄｸﾘｯﾄ');                 Kind='NUM' }
    @{ Name='血小板数';        Col=95;  Ryaku=@('血小板数');                 Kind='NUM' }
    @{ Name='総コレステロール';Col=97;  Ryaku=@('総ｺﾚｽﾃﾛｰﾙ');               Kind='NUM' }
    @{ Name='HDL';             Col=99;  Ryaku=@('HDLｺﾚｽﾃﾛｰﾙ');              Kind='NUM' }
    @{ Name='LDL';             Col=101; Ryaku=@('LDLｺﾚｽﾃﾛｰﾙ');              Kind='NUM' }
    @{ Name='中性脂肪';        Col=103; Ryaku=@('中性脂肪未判別','中性脂肪','随時中性脂肪'); Kind='NUM' }   # 帳票303(直し後)は 中性脂肪未判別
    @{ Name='AST(GOT)';        Col=106; Ryaku=@('GOT');                      Kind='NUM' }
    @{ Name='ALT(GPT)';        Col=108; Ryaku=@('GPT');                      Kind='NUM' }
    @{ Name='γ-GTP';           Col=110; Ryaku=@('γ-GTP');                    Kind='NUM' }
    @{ Name='尿素窒素';        Col=113; Ryaku=@('尿素窒素');                 Kind='NUM' }
    @{ Name='クレアチニン';    Col=115; Ryaku=@('ｸﾚｱﾁﾆﾝ');                  Kind='NUM' }
    @{ Name='尿酸';            Col=117; Ryaku=@('尿酸');                     Kind='NUM' }
    @{ Name='血糖';            Col=120; Ryaku=@('血糖','空腹時血糖','随時血糖'); Kind='NUM' }             # 帳票303(直し後)は 血糖
    @{ Name='HbA1c';           Col=122; Ryaku=@('HbA1cNGSP');                Kind='NUM' }
    @{ Name='総蛋白';          Col=125; Ryaku=@('総蛋白');                   Kind='NUM' }
    @{ Name='総ビリルビン';    Col=127; Ryaku=@('総ﾋﾞﾘﾙﾋﾞﾝ');                Kind='NUM' }
    @{ Name='アルブミン';      Col=142; Ryaku=@('ｱﾙﾌﾞﾐﾝ');                  Kind='NUM' }
    @{ Name='ALP';             Col=143; Ryaku=@('ALP');                      Kind='NUM' }
    @{ Name='LAP';             Col=144; Ryaku=@('LAP');                      Kind='NUM' }
    @{ Name='喫煙';            Col=210; Ryaku=@('特定健診_質問8');           Kind='TXT' }
    @{ Name='血圧の薬';        Col=212; Ryaku=@('薬剤治療中1');              Kind='TXT' }
    @{ Name='血糖の薬';        Col=213; Ryaku=@('薬剤治療中2');              Kind='TXT' }
    @{ Name='脂質の薬';        Col=214; Ryaku=@('薬剤治療中3');              Kind='TXT' }
)
$COLS_JIKAKU = 195..207              # 自覚症状 (見出しが症状名)
$COLS_KIOU   = @(153..172) + @(173..192)   # 既往歴(1)(2)
$COL_YMD = 1; $COL_KANJI = 4; $COL_KANA = 5

# 東振協 → 健診ナビ の判定 (code_map.csv の TOS_HANTEI と同じ)
$HANTEI_MAP = @{ 'A'='A'; 'B'='B'; 'C'='D'; 'D'='F' }
$cmPath = Join-Path $PSScriptRoot 'form\code_map.csv'
if (Test-Path $cmPath) {
    foreach ($r in (Import-Csv -Path $cmPath -Encoding UTF8)) {
        if ($r.MapName -eq 'TOS_HANTEI' -and $r.FromCode) { $HANTEI_MAP[$r.FromCode.Trim().ToUpper()] = $r.ToCode.Trim().ToUpper() }
    }
}

# ============================================================================
# 文字のそろえ方
# ============================================================================
function Narrow([string]$s) {
    if ($null -eq $s) { return '' }
    $t = [Microsoft.VisualBasic.Strings]::StrConv($s, [Microsoft.VisualBasic.VbStrConv]::Narrow, 1041)
    return ($t -replace '[\s\u3000]+', ' ').Trim()
}
function NoSpace([string]$s) { return ((Narrow $s) -replace ' ', '') }
function Norm-Nyou([string]$s) {
    $t = (NoSpace $s) -replace '[()（）]', ''
    switch -Regex ($t) {
        '^(-|ｰ|−|―|陰性)$'        { return '(-)' }
        '^(\+-|±|\+−|擬陽性|偽陽性)$' { return '(+-)' }
        '^(\+|1\+|陽性)$'          { return '(+)' }
        '^(\+\+|2\+)$'             { return '(2+)' }
        '^(\+\+\+|3\+)$'           { return '(3+)' }
        '^(\+\+\+\+|4\+)$'         { return '(4+)' }
        '^$'                       { return '' }
        default                    { return $t }
    }
}
function Norm-Cho([string]$s) {
    $t = NoSpace $s
    switch -Regex ($t) {
        '^(1|異常なし|なし|正常|所見なし|所見無し)$' { return '所見なし' }
        '^(2|異常あり|あり|所見あり|所見有り)$'      { return '所見あり' }
        '^$'                                        { return '' }
        default                                     { return $t }
    }
}
function Norm-Num([string]$s) {
    $t = (NoSpace $s) -replace '[^0-9.\-]', ''
    $d = 0.0
    if ($t -ne '' -and [double]::TryParse($t, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        return $d.ToString('0.####', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return (NoSpace $s)
}
function Norm-Han([string]$s, [bool]$fromCsv) {
    $t = (NoSpace $s).ToUpper()
    if ($fromCsv -and $HANTEI_MAP.ContainsKey($t)) { return $HANTEI_MAP[$t] }
    return $t
}

# ============================================================================
# CSV (cp932, ダブルクォート対応)
# ============================================================================
function Read-CsvRows([string]$path) {
    $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($path, [System.Text.Encoding]::GetEncoding(932))
    $p.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters(',')
    $p.HasFieldsEnclosedInQuotes = $true
    $p.TrimWhiteSpace = $false
    $rows = @()
    while (-not $p.EndOfData) { $rows += ,$p.ReadFields() }
    $p.Close()
    return ,$rows
}
function F($fields, [int]$col) {
    if ($col -le 0 -or $col -gt $fields.Count) { return '' }
    return [string]$fields[$col - 1]
}

# ============================================================================
# xlsx を直接読む → "sheet1.xml!AI52" → 値
# ============================================================================
function Read-XlsxCells([string]$path) {
    $NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $getXml = {
            param($name)
            foreach ($e in $zip.Entries) {
                if ($e.FullName -ne $name) { continue }
                $st = $e.Open(); $sr = New-Object System.IO.StreamReader($st, [System.Text.Encoding]::UTF8)
                $txt = $sr.ReadToEnd(); $sr.Close(); $st.Close()
                $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml($txt); return ,$doc
            }
            return $null
        }
        $shared = @()
        $xs = & $getXml 'xl/sharedStrings.xml'
        if ($xs) {
            $m = New-Object System.Xml.XmlNamespaceManager($xs.NameTable); $m.AddNamespace('m', $NS)
            foreach ($si in $xs.SelectNodes('/m:sst/m:si', $m)) {
                $sb = New-Object System.Text.StringBuilder
                foreach ($t in $si.SelectNodes('m:t | m:r/m:t', $m)) { [void]$sb.Append($t.InnerText) }
                $shared += $sb.ToString()
            }
        }
        $cells = @{}
        foreach ($e in @($zip.Entries)) {
            if ($e.FullName -notmatch '^xl/worksheets/(sheet\d+\.xml)$') { continue }
            $sheet = $Matches[1]
            $doc = & $getXml $e.FullName
            $m = New-Object System.Xml.XmlNamespaceManager($doc.NameTable); $m.AddNamespace('m', $NS)
            foreach ($c in $doc.SelectNodes('/m:worksheet/m:sheetData/m:row/m:c', $m)) {
                $t = $c.GetAttribute('t'); $v = $c.SelectSingleNode('m:v', $m); $val = ''
                if ($t -eq 'inlineStr') {
                    foreach ($tn in $c.SelectNodes('m:is//m:t', $m)) { $val += $tn.InnerText }
                }
                elseif ($t -eq 's') {
                    if ($v) { $ix = -1; if ([int]::TryParse($v.InnerText, [ref]$ix) -and $ix -ge 0 -and $ix -lt $shared.Count) { $val = $shared[$ix] } }
                }
                elseif ($v) {
                    $raw = $v.InnerText; $d = 0.0
                    if ($t -ne 'str' -and $t -ne 'e' -and [double]::TryParse($raw, [System.Globalization.NumberStyles]::Float,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
                        $val = $d.ToString('0.##########', [System.Globalization.CultureInfo]::InvariantCulture)
                    } else { $val = $raw }
                }
                if ($val -ne '') { $cells[$sheet + '!' + $c.GetAttribute('r')] = $val }
            }
        }
        return $cells
    }
    finally { $zip.Dispose() }
}

# ============================================================================
# 入力
# ============================================================================
if (-not $Folder) { $Folder = Read-Host '結果報告書(.xlsx)の入ったフォルダをドラッグ＆ドロップして Enter' }
$Folder = ($Folder -replace '^"|"$', '').Trim()
if (-not (Test-Path $Folder)) { throw "見つかりません: $Folder" }
if (-not $Csv) { $Csv = Read-Host '東振協のCSVをドラッグ＆ドロップして Enter' }
$Csv = ($Csv -replace '^"|"$', '').Trim()
if (-not (Test-Path $Csv)) { throw "見つかりません: $Csv" }
if (-not (Test-Path $CellMap)) { throw "帳票のセル対応表がありません: $CellMap" }

$xlsxFiles = @()
if ((Get-Item $Folder).PSIsContainer) {
    $xlsxFiles = @(Get-ChildItem -LiteralPath $Folder -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)
} else { $xlsxFiles = @(Get-Item $Folder) }
if ($xlsxFiles.Count -eq 0) { throw "フォルダに .xlsx がありません: $Folder" }

# ---- 帳票のセル対応 (略称 → セル) ----
$cellOf = @{}    # 略称 → 'sheet2.xml!AI52'  (結果値_1)
$hanOf  = @{}    # 略称 → セル (判定_1)
$misc   = @{}    # 漢字氏名 / ｶﾅ氏名 / 既往歴_状況N_1 / 自覚症状N
foreach ($r in (Import-Csv -Path $CellMap -Encoding UTF8)) {
    $key = ($r.Sheet -replace '^Sheet', 'sheet') + '.xml!' + $r.Cell
    if ($r.Kind -eq '結果値_1' -and $r.Ryaku -ne '') { $cellOf[$r.Ryaku] = $key }
    elseif ($r.Kind -eq '判定_1' -and $r.Ryaku -ne '') { $hanOf[$r.Ryaku] = $key }
    elseif ($r.Ryaku -eq '') { $misc[$r.Kind] = $key }
}
foreach ($need in @('身長','GOT','視力右','尿蛋白','自覚症状1')) {
    if (-not $cellOf.ContainsKey($need)) { throw "セル対応表に [$need] がありません。帳票303のひな形から作り直してください。" }
}
if (-not $misc.ContainsKey('漢字氏名')) { throw 'セル対応表に **漢字氏名 がありません。' }

# ---- CSV ----
$rows = Read-CsvRows $Csv
if ($rows.Count -lt 2) { throw "CSVにデータがありません: $Csv" }
$head = $rows[0]; $data = $rows[1..($rows.Count - 1)]
if ($head.Count -lt 214) { throw ("列が足りません ({0}列)。東振協の250列のCSVを渡してください。" -f $head.Count) }
$ymd = (F $data[0] $COL_YMD) -replace '[^0-9]', ''
$csvBy = @{}      # 氏名(空白なし) → fields
$csvKana = @{}
foreach ($f in $data) {
    $k = NoSpace (F $f $COL_KANJI); $a = NoSpace (F $f $COL_KANA)
    if ($k -ne '') { $csvBy[$k] = $f }
    if ($a -ne '') { $csvKana[$a] = $f }
}
$jikakuName = @{}; foreach ($c in $COLS_JIKAKU) { $jikakuName[$c] = Narrow $head[$c - 1] }
$kiouName   = @{}; foreach ($c in $COLS_KIOU)   { $kiouName[$c]   = Narrow $head[$c - 1] }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 結果報告書(Excel) と 東振協CSV の突き合わせ' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  Excel: {0} ファイル ({1})" -f $xlsxFiles.Count, $Folder)
Write-Host ("  CSV  : {0} 人 ({1})" -f $data.Count, [System.IO.Path]::GetFileName($Csv))
Write-Host ''

# ============================================================================
# 突き合わせ
# ============================================================================
$all    = @()      # 全項目 (CSV出力用)
$people = @()      # 人ごとの集計
$usedCsv = @{}
$n = 0
foreach ($fx in $xlsxFiles) {
    $n++
    Write-Host ("  [{0}/{1}] {2}" -f $n, $xlsxFiles.Count, $fx.Name) -ForegroundColor DarkGray
    try { $cells = Read-XlsxCells $fx.FullName }
    catch {
        $people += New-Object PSObject -Property @{ 氏名 = $fx.Name; 一致 = 0; 不一致 = 0; 目視 = 0; 状態 = ("読めません: " + $_.Exception.Message) }
        continue
    }
    $xName = ''; if ($cells.ContainsKey($misc['漢字氏名'])) { $xName = NoSpace $cells[$misc['漢字氏名']] }
    $xKana = ''; if ($misc.ContainsKey('ｶﾅ氏名') -and $cells.ContainsKey($misc['ｶﾅ氏名'])) { $xKana = NoSpace $cells[$misc['ｶﾅ氏名']] }
    $label = $(if ($xName -ne '') { $xName } else { $fx.BaseName })

    $f = $null
    if ($xName -ne '' -and $csvBy.ContainsKey($xName)) { $f = $csvBy[$xName] }
    elseif ($xKana -ne '' -and $csvKana.ContainsKey($xKana)) { $f = $csvKana[$xKana] }
    if ($null -eq $f) {
        # 東振協のCSVは受診した人しか載らない。帳票がほぼ空なら未受診。
        $filled = 0
        foreach ($ry in @('身長','体重','GOT','白血球数','視力右')) {
            if ($cellOf.ContainsKey($ry) -and $cells.ContainsKey($cellOf[$ry])) { $filled++ }
        }
        $st = $(if ($filled -eq 0) { '未受診 (CSVに無く帳票も空)。印刷しない' } else { 'CSVにこの人がいません (帳票には値あり)' })
        $people += New-Object PSObject -Property @{ 氏名 = $label; 一致 = 0; 不一致 = 0; 目視 = 0; 状態 = $st }
        continue
    }
    $usedCsv[(NoSpace (F $f $COL_KANJI))] = 1

    $ok = 0; $ng = 0; $eye = 0
    $add = {
        param($item, $csvV, $xlV, $res, $note)
        $script:all += New-Object PSObject -Property @{
            氏名 = $label; 項目 = $item; CSV = $csvV; 帳票 = $xlV; 結果 = $res; 備考 = $note }
    }

    foreach ($ru in $RULES) {
        $csvRaw = Narrow (F $f $ru.Col)
        if ($ru.ContainsKey('Col2')) {
            # 1回目と2回目の両方があれば低い方。片方だけならある方。
            $v2 = Narrow (F $f $ru.Col2)
            $d1 = 0.0; $d2 = 0.0
            $ok1 = [double]::TryParse((Norm-Num $csvRaw), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d1)
            $ok2 = [double]::TryParse((Norm-Num $v2),     [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d2)
            if ($ok1 -and $ok2) { if ($d2 -lt $d1) { $csvRaw = $v2 } }
            elseif ($ok2)       { $csvRaw = $v2 }
        }
        # 帳票側: 候補の略称のうち、値が入っているセルを使う
        $xlRaw = ''; $slot = ''
        foreach ($ry in $ru.Ryaku) {
            $key = $null
            if ($ru.Kind -eq 'HAN') { if ($hanOf.ContainsKey($ry)) { $key = $hanOf[$ry] } }
            else                    { if ($cellOf.ContainsKey($ry)) { $key = $cellOf[$ry] } }
            if ($key -and $cells.ContainsKey($key)) { $v = Narrow $cells[$key]; if ($v -ne '') { $xlRaw = $v; $slot = $ry; break } }
        }
        if ($csvRaw -eq '' -and $xlRaw -eq '') { continue }     # 両方空 = 検査していない
        if ($ru.Kind -eq 'CODE') {
            if ($ru.Ryaku.Count -eq 0) {
                if ($csvRaw -ne '') { $eye++; & $add $ru.Name $csvRaw '' '目視' '帳票に枠が無い項目' }
                continue
            }
            $eye++; & $add $ru.Name $csvRaw $xlRaw '目視' 'コード→文字なので目で確認'
            continue
        }
        $a = ''; $b = ''
        switch ($ru.Kind) {
            'NUM'  { $a = Norm-Num $csvRaw;      $b = Norm-Num $xlRaw }
            'NYOU' { $a = Norm-Nyou $csvRaw;     $b = Norm-Nyou $xlRaw }
            'CHO'  { $a = Norm-Cho $csvRaw;      $b = Norm-Cho $xlRaw }
            'HAN'  { $a = Norm-Han $csvRaw $true; $b = Norm-Han $xlRaw $false }
            default{ $a = NoSpace $csvRaw;       $b = NoSpace $xlRaw }
        }
        if ($a -eq $b) { $ok++; & $add $ru.Name $csvRaw $xlRaw '一致' $slot; continue }
        if ($ru.Kind -eq 'TXT' -and $a -ne '' -and $b -ne '' -and ($b.StartsWith($a) -or $a.StartsWith($b))) {
            $ok++; & $add $ru.Name $csvRaw $xlRaw '一致' '表記が少し違うだけ'; continue
        }
        $ng++
        $why = ''
        if ($csvRaw -ne '' -and $xlRaw -eq '')     { $why = '★CSVにあるのに帳票に出ていない' }
        elseif ($csvRaw -eq '' -and $xlRaw -ne '') { $why = '★CSVに無いのに帳票に値がある' }
        else                                       { $why = '★値が違う' }
        & $add $ru.Name $csvRaw $xlRaw '不一致' $why
    }

    # ---- 自覚症状 (前から4つ) ----
    $hits = @(); foreach ($c in $COLS_JIKAKU) { if ((Narrow (F $f $c)) -ne '') { $hits += $jikakuName[$c] } }
    $hits = @($hits | Select-Object -First 4)
    for ($i = 1; $i -le 4; $i++) {
        $exp = $(if ($i -le $hits.Count) { $hits[$i - 1] } else { '' })
        $key = $null; if ($cellOf.ContainsKey("自覚症状$i")) { $key = $cellOf["自覚症状$i"] }
        $got = ''; if ($key -and $cells.ContainsKey($key)) { $got = Narrow $cells[$key] }
        if ($exp -eq '' -and $got -eq '') { continue }
        if ((NoSpace $exp) -eq (NoSpace $got)) { $ok++; & $add "自覚症状$i" $exp $got '一致' '' }
        else { $ng++; & $add "自覚症状$i" $exp $got '不一致' '★自覚症状が違う' }
    }

    # ---- 既往歴 (前から4つ。(1)を先に、(2)にしか無いものを後ろに) ----
    $kh = @(); $seen = @{}
    foreach ($c in $COLS_KIOU) {
        if ((Narrow (F $f $c)) -eq '') { continue }
        $nm = $kiouName[$c]
        if (-not $seen.ContainsKey($nm)) { $seen[$nm] = 1; $kh += $nm }
    }
    $kh = @($kh | Select-Object -First 4)
    for ($i = 1; $i -le 4; $i++) {
        $exp = $(if ($i -le $kh.Count) { $kh[$i - 1] } else { '' })
        $key = $null; if ($misc.ContainsKey("既往歴_状況${i}_1")) { $key = $misc["既往歴_状況${i}_1"] }
        $got = ''; if ($key -and $cells.ContainsKey($key)) { $got = Narrow $cells[$key] }
        if ($kh.Count -eq 0 -and $i -eq 1 -and (NoSpace $got) -eq '特記事項なし') { $ok++; & $add '既往歴1' '(なし)' $got '一致' '既往歴なし'; continue }
        if ($exp -eq '' -and $got -eq '') { continue }
        if ((NoSpace $exp) -eq (NoSpace $got)) { $ok++; & $add "既往歴$i" $exp $got '一致' '' }
        else { $ng++; & $add "既往歴$i" $exp $got '不一致' '★既往歴が違う' }
    }

    # ---- 総合判定: 帳票 (健診ナビの自動判定) と CSV10列目 (提供元の総合判定) を並べる (参考) ----
    $sg = ''
    if ($hanOf.ContainsKey('総合判定(編集)') -and $cells.ContainsKey($hanOf['総合判定(編集)'])) { $sg = Narrow $cells[$hanOf['総合判定(編集)']] }
    $csvSgRaw = (NoSpace (F $f 10)).ToUpper()
    $csvSg = Norm-Han $csvSgRaw $true          # 東振協 A/B/C/D → 健診ナビ A/B/D/F に読み替え

    $people += New-Object PSObject -Property @{
        氏名 = $label; 一致 = $ok; 不一致 = $ng; 目視 = $eye; 総合 = $sg; CSV総合 = $csvSgRaw; CSV総合換算 = $csvSg
        状態 = $(if ($ng -eq 0) { 'OK' } else { "不一致 $ng 件" }) }
}

# CSVにいてExcelが無い人
$noXlsx = @()
foreach ($f in $data) {
    $k = NoSpace (F $f $COL_KANJI)
    if ($k -ne '' -and -not $usedCsv.ContainsKey($k)) { $noXlsx += $k }
}

# ============================================================================
# 出す
# ============================================================================
$outTxt = Join-Path $PSScriptRoot ("照合結果_{0}.txt" -f $ymd)
$outCsv = Join-Path $PSScriptRoot ("照合結果_{0}.csv" -f $ymd)
$L = New-Object System.Collections.Generic.List[string]
$L.Add(("=== 結果報告書(Excel) と 東振協CSV の突き合わせ  {0} ===" -f (Get-Date -Format 'yyyy/MM/dd HH:mm')))
$L.Add(("Excel: {0} ファイル / CSV: {1} 人" -f $xlsxFiles.Count, $data.Count))
$L.Add('')
$L.Add('--- 1. 人ごとの結果 ---')
$L.Add(($people | Sort-Object 氏名 | Select-Object 氏名, 総合, 一致, 不一致, 目視, 状態 | Format-Table -AutoSize | Out-String -Width 160).TrimEnd())
$L.Add('')
$ngRows = @($all | Where-Object { $_.結果 -eq '不一致' })
$L.Add(("--- 2. ★不一致の一覧 ({0} 件) ---" -f $ngRows.Count))
if ($ngRows.Count -eq 0) { $L.Add('  ありません。比べられる項目は全員ぶん一致しました。') }
else { $L.Add(($ngRows | Select-Object 氏名, 項目, CSV, 帳票, 備考 | Format-Table -AutoSize -Wrap | Out-String -Width 160).TrimEnd()) }
$L.Add('')
$eyeRows = @($all | Where-Object { $_.結果 -eq '目視' })
$L.Add(("--- 3. 目視で確かめる項目 (所見はCSVがコードなので合否を付けていません) ({0} 件) ---" -f $eyeRows.Count))
$L.Add(($eyeRows | Select-Object 氏名, 項目, CSV, 帳票 | Format-Table -AutoSize -Wrap | Out-String -Width 160).TrimEnd())
$L.Add('')
# 総合判定の比較 (参考)。基準が違うので一致しなくても誤りではないが、大きく違う人は見ておく
$RANK = @{ 'A'=1; 'B'=2; 'C'=3; 'D'=4; 'E'=5; 'F'=6; 'G'=7; 'J'=7 }
$sgRows = @()
foreach ($pp in ($people | Where-Object { $_.状態 -eq 'OK' -or $_.状態 -like '不一致*' })) {
    $x = [string]$pp.総合; $y = [string]$pp.CSV総合換算
    if ($x -eq '' -and $y -eq '') { continue }
    $rx = 0; $ry = 0
    if ($RANK.ContainsKey($x)) { $rx = $RANK[$x] }
    if ($RANK.ContainsKey($y)) { $ry = $RANK[$y] }
    $rel = ''
    if ($x -eq $y)      { $rel = '同じ' }
    elseif ($rx -gt $ry) { $rel = '帳票の方が悪い' }
    elseif ($rx -lt $ry) { $rel = '★帳票の方が軽い' }
    else                 { $rel = '?' }
    $sgRows += New-Object PSObject -Property @{ 氏名 = $pp.氏名; 帳票 = $x; CSV = $pp.CSV総合; CSV換算 = $y; 関係 = $rel }
}
$L.Add(("--- 3b. 参考: 総合判定 帳票(健診ナビ) vs CSV10列目(提供元) ({0} 人) ---" -f $sgRows.Count))
$L.Add('  判定の基準が違うので一致しなくても誤りではありません。「★帳票の方が軽い」の人だけ理由を確認してください。')
$L.Add(('  同じ: {0} 人 / 帳票の方が悪い: {1} 人 / ★帳票の方が軽い: {2} 人' -f
    @($sgRows | Where-Object { $_.関係 -eq '同じ' }).Count,
    @($sgRows | Where-Object { $_.関係 -eq '帳票の方が悪い' }).Count,
    @($sgRows | Where-Object { $_.関係 -like '★*' }).Count))
$L.Add(($sgRows | Sort-Object 関係, 氏名 | Select-Object 氏名, 帳票, CSV, CSV換算, 関係 | Format-Table -AutoSize | Out-String -Width 120).TrimEnd())
$L.Add('')
$L.Add(("--- 4. CSVにいるのに結果報告書のExcelが無い人 ({0} 人) ---" -f $noXlsx.Count))
$L.Add('  ' + ($noXlsx -join ' / '))
$L.Add('  (未受診の人はここに出ます)')
$L.Add('')
$okP = @($people | Where-Object { $_.状態 -eq 'OK' }).Count
$L.Add('=== まとめ ===')
$L.Add(("  比べた人: {0} 人 / 全項目一致: {1} 人 / 不一致あり: {2} 人 / CSVにいない・読めない: {3} 人" -f
    $people.Count, $okP, @($people | Where-Object { $_.状態 -like '不一致*' }).Count,
    @($people | Where-Object { $_.状態 -ne 'OK' -and $_.状態 -notlike '不一致*' }).Count))
$L.Add(("  比べた項目: {0} 件 / 一致 {1} 件 / 不一致 {2} 件 / 目視 {3} 件" -f
    $all.Count, @($all | Where-Object { $_.結果 -eq '一致' }).Count, $ngRows.Count, $eyeRows.Count))
$L.Add('')
$L.Add('  ※ 比べていないもの: BMI・eGFR・判定(血圧など)・医師指示事項 は健診ナビが計算するのでCSVにありません。')
$L.Add('  ※ 「目視」は 3 の一覧で、CSVのコードと帳票の文字が合っているかを見てください。')

$enc = [System.Text.Encoding]::GetEncoding(932)
[System.IO.File]::WriteAllLines($outTxt, $L, $enc)
$all | Select-Object 氏名, 項目, CSV, 帳票, 結果, 備考 | Export-Csv -Path $outCsv -NoTypeInformation -Encoding Default

Write-Host ''
foreach ($ln in $L) { Write-Host $ln }
Write-Host ''
Write-Host ("[出力] {0}" -f $outTxt) -ForegroundColor Green
Write-Host ("[出力] {0}  (Excelで開いて絞り込めます)" -f $outCsv) -ForegroundColor Green
notepad $outTxt
