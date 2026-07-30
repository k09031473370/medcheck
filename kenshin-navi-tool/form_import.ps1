<#
.SYNOPSIS
  健診ナビ11 検査結果一括取込ツール (form_import.ps1)

.DESCRIPTION
  自院Excelフォーム(1行目ヘッダ・1人1行)をCSV保存したものを読み込み、
  健診ナビのDB (SQL Server / T_KENSA) へ UPDATE で結果を書き込む。

  - 人の特定 : T_KANJA_G の KEN_YMD('YYYY/MM/DD') + KEN_NO(受付番号) → PK_SEQ
  - 結果書込 : UPDATE T_KENSA SET KEKKA=値 WHERE PK_SEQ=@p AND KOMOKU_CD=@cd
               (枠=T_KENSA行は予約時に作成済みの前提。INSERTは行わない)
  - 所見項目 : KEKKA=所見文 / KEKKA_CD=結果CD / HANTEI_KIGO=判定
               所見マスタ T_SYOKEN2(SYOKEN_CD,KEKKA_CD,SYOKEN,HANTEI_KIGO) と突合
  - BMI・総合判定などは書込後に健診ナビの「自動判定」で計算する

.EXAMPLE
  # プレビュー(書込なし)
  powershell -ExecutionPolicy Bypass -File C:\Users\User\Documents\excel_tools\form_import.ps1 -Csv C:\...\form.csv -Only 4001 -KenYmd 2026/07/02

  # 書込
  powershell -ExecutionPolicy Bypass -File C:\Users\User\Documents\excel_tools\form_import.ps1 -Csv C:\...\form.csv -Only 4001 -KenYmd 2026/07/02 -Commit

  # CSVの列番号・ヘッダ・値の確認(DB接続なし)
  powershell -ExecutionPolicy Bypass -File C:\Users\User\Documents\excel_tools\form_import.ps1 -Csv C:\...\form.csv -Only 4001 -Inspect

  # 対象者の T_KENSA 行(KOMOKU_CDと現在値)を一覧表示 + CSVダンプ
  powershell -ExecutionPolicy Bypass -File C:\Users\User\Documents\excel_tools\form_import.ps1 -Only 4001 -KenYmd 2026/07/02 -DumpItems

  # 所見マスタの一覧 (例: 胸部X線の所見側)
  powershell -ExecutionPolicy Bypass -File C:\Users\User\Documents\excel_tools\form_import.ps1 -DumpSyoken ZK021
#>
[CmdletBinding()]
param(
    [string]$Csv,                 # フォームCSVのパス
    [string]$Only,                # 受付番号で絞り込み (例: 4001)
    [switch]$Commit,              # 付けると書込。付けなければプレビューのみ
    [switch]$Force,               # エラー行があってもOK行のみ書込
    [switch]$Inspect,             # CSVの列番号/ヘッダ/値/変換結果を表示(DB接続なし)
    [switch]$NoHeader,            # 1行目からデータの場合に指定(ヘッダ行なし)
    [switch]$SetUkeNo,            # 受付番号を健診ナビへ設定 (氏名で照合)
    [switch]$Overwrite,           # -SetUkeNo で既存の受付番号も上書きする
    [switch]$DumpItems,           # 対象者のT_KENSA行を一覧表示
    [switch]$ShowYmd,             # ファイルの日付と指定日を表示するだけ (GUIの確認用・DB接続なし)
    [switch]$IgnoreName,          # 氏名が一致しなくても取り込む (テストデータ等)
    [string]$DumpSyoken,          # 指定SYOKEN_CDのT_SYOKEN2一覧を表示 (SHIN/GANTEI/ZK011/ZK020/ZK021/ZK030/ZK031/ZK041)
    [string]$KenYmd,              # 受診日 'YYYY/MM/DD'。CSVに日付列が無い場合に指定
    [string]$Roster,              # 名簿ファイル(Excel/CSV)。ここに載っている人だけを取り込む
    [string]$MapDir,              # 対応表フォルダ (既定: スクリプトと同じ場所の form\)
    [string]$Mapping,             # 使用する対応表ファイル名 (既定: mapping.csv)
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString,    # 接続文字列を直接指定する場合
    [ValidateSet('SJIS','UTF8')]
    [string]$CsvEncoding = 'SJIS' # ExcelからのCSV保存は通常SJIS
)

$ErrorActionPreference = 'Stop'
if (-not $MapDir) { $MapDir = Join-Path $PSScriptRoot 'form' }
$BackupDir = Join-Path $PSScriptRoot 'backup'

# --- SJIS(cp932) が使えるようにする (PowerShell 7 対策。5.1では最初から使える) ---
try { [void][System.Text.Encoding]::GetEncoding(932) }
catch { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) }

# ============================================================================
# 汎用関数
# ============================================================================

function Normalize-Text([string]$s) {
    if ($null -eq $s) { return '' }
    return $s.Trim()
}

function Normalize-KenNo([string]$s) {
    $v = Normalize-Text $s
    # Excel数値の "4001.0" 対策
    $v = $v -replace '\.0+$', ''
    return $v
}

function Normalize-Ymd([string]$s) {
    $v = Normalize-Text $s
    if ($v -eq '') { return $null }
    # Excelシリアル値 (例: 46204)
    if ($v -match '^\d{4,6}(\.0+)?$' -and [double]$v -gt 20000 -and [double]$v -lt 80000) {
        return ([datetime]::FromOADate([double]$v)).ToString('yyyy/MM/dd')
    }
    $dt = [datetime]::MinValue
    #  yyMMdd (260712) も受け付ける。Excelシリアル値(20000〜80000)は上で処理済みなので衝突しない
    $formats = @('yyyy/MM/dd','yyyy/M/d','yyyy-MM-dd','yyyy-M-d','yyyyMMdd','yyyy年M月d日','yyyy.M.d','yyMMdd','yy/MM/dd')
    foreach ($f in $formats) {
        if ([datetime]::TryParseExact($v, $f, [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.ToString('yyyy/MM/dd')
        }
    }
    if ([datetime]::TryParse($v, [ref]$dt)) { return $dt.ToString('yyyy/MM/dd') }
    return $null
}

# CSVテキストをパース (ダブルクォート・カンマ・改行入りフィールド対応)
function Parse-CsvText([string]$text) {
    $rows = New-Object System.Collections.ArrayList
    $cur  = New-Object System.Collections.ArrayList
    $sb   = New-Object System.Text.StringBuilder
    $inQ  = $false
    $n = $text.Length
    for ($i = 0; $i -lt $n; $i++) {
        $c = $text[$i]
        if ($inQ) {
            if ($c -eq '"') {
                if (($i + 1) -lt $n -and $text[$i + 1] -eq '"') { [void]$sb.Append('"'); $i++ }
                else { $inQ = $false }
            } else { [void]$sb.Append($c) }
        }
        else {
            if ($c -eq '"') { $inQ = $true }
            elseif ($c -eq ',') { [void]$cur.Add($sb.ToString()); [void]$sb.Clear() }
            elseif ($c -eq "`r") { }
            elseif ($c -eq "`n") {
                [void]$cur.Add($sb.ToString()); [void]$sb.Clear()
                [void]$rows.Add($cur.ToArray()); $cur.Clear()
            }
            else { [void]$sb.Append($c) }
        }
    }
    if ($sb.Length -gt 0 -or $cur.Count -gt 0) {
        [void]$cur.Add($sb.ToString())
        [void]$rows.Add($cur.ToArray())
    }
    return ,$rows.ToArray()
}

# .xlsx / .xlsm の読み込み (Excelを使わない共通部品)
$XlsxLib = Join-Path $PSScriptRoot 'xlsx_read.ps1'
if (-not (Test-Path $XlsxLib)) { throw "xlsx_read.ps1 が見つかりません: $XlsxLib" }
. $XlsxLib

function Read-FormCsv([string]$path, [string]$encName) {
    if (-not (Test-Path $path)) { throw "ファイルが見つかりません: $path" }
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    if ($ext -eq '.xlsx' -or $ext -eq '.xlsm') {
        $rows = Read-Xlsx $path
    }
    else {
        $enc = if ($encName -eq 'UTF8') { New-Object System.Text.UTF8Encoding($false) }
               else { [System.Text.Encoding]::GetEncoding(932) }
        $text = [System.IO.File]::ReadAllText($path, $enc)
        # UTF-8 BOM除去
        if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
        $rows = Parse-CsvText $text
    }
    # 完全空行を除去
    $out = @()
    foreach ($r in $rows) {
        $isEmpty = $true
        foreach ($f in $r) { if ((Normalize-Text $f) -ne '') { $isEmpty = $false; break } }
        if (-not $isEmpty) { $out += ,$r }
    }
    if ($out.Count -lt 1) { throw "CSVが空です: $path" }
    return ,$out
}

# ヘッダ行とデータ行に分離
#  -NoHeader 指定時は必ず見出し無し。未指定なら 1行目の受付番号列を見て自動判定する
#  (受付番号列が数字ならデータ行、そうでなければ見出し行)
function Split-HeaderData($rows, $idCols) {
    $noHdr = [bool]$NoHeader
    if (-not $noHdr -and $idCols -and $idCols.KenNo -gt 0) {
        $v = Normalize-KenNo (Get-Field $rows[0] $idCols.KenNo)
        if ($v -match '^\d+$') {
            $noHdr = $true
            Write-Host '[自動判別] 見出し行なし (1行目からデータ)' -ForegroundColor DarkGray
        }
    }
    if ($noHdr) {
        $h = @()
        for ($i = 1; $i -le $rows[0].Count; $i++) { $h += "列$i" }
        return @{ Header = $h; Data = $rows }
    }
    if ($rows.Count -lt 2) { throw '見出し行だけでデータがありません。2行目以降に内容を入力してください。' }
    return @{ Header = $rows[0]; Data = @($rows[1..($rows.Count - 1)]) }
}

function Get-Field($fields, [int]$col) {
    if ($col -le 0 -or $col -gt $fields.Count) { return '' }
    return [string]$fields[$col - 1]
}

# ============================================================================
# 対応表の読み込み
# ============================================================================

# CSVの1行目を見て、どの対応表(レイアウト)かを自動で判別する
#   対応表に書く指示: # DETECT=HEADER:受付NO,Q1   / # DETECT=COLS:40-50
function Detect-MappingPath([string]$path) {
    if (-not $path) { $path = $Csv }
    if (-not $path -or -not (Test-Path $path)) { return $null }
    $first = ''
    $colCount = 0
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    if ($ext -eq '.xlsx' -or $ext -eq '.xlsm') {
        try {
            $r1 = (Read-Xlsx $path)[0]
            $first = ($r1 -join ',')
            $colCount = $r1.Count
        } catch { return $null }
    }
    else {
        $enc = if ($CsvEncoding -eq 'UTF8') { New-Object System.Text.UTF8Encoding($false) } else { [System.Text.Encoding]::GetEncoding(932) }
        try {
            $sr = New-Object System.IO.StreamReader($path, $enc)
            $first = $sr.ReadLine()
            $sr.Close()
        } catch { return $null }
        if ($null -eq $first) { return $null }
        if ($first.Length -gt 0 -and $first[0] -eq [char]0xFEFF) { $first = $first.Substring(1) }
        $colCount = ($first -split ',').Count
    }

    $byHeader = $null; $byCols = $null
    foreach ($f in (Get-ChildItem -Path $MapDir -Filter 'mapping*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        foreach ($ln in (Get-Content $f.FullName -TotalCount 8 -Encoding UTF8)) {
            if ($ln -notmatch '^#\s*DETECT\s*=\s*(.+)$') { continue }
            $spec = $Matches[1].Trim()
            if ($spec -match '^(?i)HEADER\s*:\s*(.+)$') {
                $kws = @($Matches[1] -split '[,|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
                $all = $true
                foreach ($k in $kws) { if ($first -notlike "*$k*") { $all = $false; break } }
                if ($all -and -not $byHeader) { $byHeader = $f.FullName }
            }
            elseif ($spec -match '^(?i)COLS\s*:\s*(\d+)\s*-\s*(\d+)$') {
                if ($colCount -ge [int]$Matches[1] -and $colCount -le [int]$Matches[2] -and -not $byCols) { $byCols = $f.FullName }
            }
        }
    }
    if ($byHeader) { return $byHeader }
    return $byCols
}

function Load-Mapping {
    $p = $null
    if ($Mapping -and $Mapping -ne 'auto') {
        $p = if ([System.IO.Path]::IsPathRooted($Mapping)) { $Mapping } else { Join-Path $MapDir $Mapping }
    }
    else {
        $p = Detect-MappingPath
        if ($p) { Write-Host "[自動判別] $([System.IO.Path]::GetFileName($p))" -ForegroundColor DarkGray }
        else {
            $d = Join-Path $MapDir 'mapping.csv'
            if (Test-Path $d) { $p = $d }
            else {
                $cand = @(Get-ChildItem -Path $MapDir -Filter 'mapping*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name)
                if ($cand.Count -eq 0) { throw "対応表(mapping*.csv)が $MapDir にありません。" }
                # ファイルを使わないモード(枠一覧・所見マスタ)では対応表の中身を使わないので、判別できなくてよい
                if (-not $Csv) { $p = $cand[0].FullName }
                else {
                    throw ("レイアウトを自動判別できませんでした。GUIの「レイアウト」で選んでください。候補: " + (($cand | ForEach-Object { $_.Name }) -join ' / '))
                }
            }
        }
    }
    if (-not (Test-Path $p)) { throw "対応表が見つかりません: $p" }
    Write-Host "[対応表] $([System.IO.Path]::GetFileName($p))" -ForegroundColor DarkGray
    $rows = Import-Csv -Path $p -Encoding UTF8
    $valid = @('KENNO','KENYMD','NAMEKANJI','NAMEKANA','VALUE','NYOU','CHORYOKU','MONSHIN','VISION','SHOKEN','SHOKEN2','SHOKENCD','SHOKENCD2','IGNORE')
    foreach ($r in $rows) {
        $k = (Normalize-Text $r.Kind).ToUpper()
        if ($k -ne '' -and $valid -notcontains $k) {
            throw "mapping.csv の Kind が不正です: '$($r.Kind)' (行: $($r.Label))"
        }
    }
    return $rows
}

# コード変換表 (フォームのコード → 健診ナビの結果CD)
# form\code_map.csv: MapName,FromCode,ToCode,Note
function Load-CodeMap {
    $p = Join-Path $MapDir 'code_map.csv'
    $map = @{}
    if (-not (Test-Path $p)) { return $map }
    foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
        $name = (Normalize-Text $r.MapName).ToUpper()
        $from = Normalize-Text $r.FromCode
        if ($name -eq '' -or $from -eq '') { continue }
        if (-not $map.ContainsKey($name)) { $map[$name] = @{} }
        $map[$name][$from] = Normalize-Text $r.ToCode
    }
    return $map
}

# mapping.csv の CodeMap 列に従ってコードを変換
# 戻り値: @{ Code=変換後; Status='OK'|'NOMAP'|'DROP' }
#   DROP = 変換表に「空欄」と定義済み(意図的に取り込まない値)
function Convert-Code($m, [string]$raw) {
    $mapName = (Normalize-Text $m.CodeMap).ToUpper()
    if ($mapName -eq '') { return @{ Code = $raw; Status = 'OK' } }
    $cm = $script:CodeMap
    if (-not $cm.ContainsKey($mapName)) { return @{ Code = $raw; Status = 'NOMAP' } }
    if (-not $cm[$mapName].ContainsKey($raw)) { return @{ Code = $raw; Status = 'NOMAP' } }
    $to = $cm[$mapName][$raw]
    if ($to -eq '') { return @{ Code = ''; Status = 'DROP' } }
    return @{ Code = $to; Status = 'OK' }
}

function Load-ValueMap {
    $p = Join-Path $MapDir 'value_map.csv'
    $map = @{}
    if (-not (Test-Path $p)) { return $map }
    foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
        $kind = (Normalize-Text $r.Kind).ToUpper()
        if ($kind -eq '') { continue }
        if (-not $map.ContainsKey($kind)) { $map[$kind] = @{} }
        $map[$kind][(Normalize-Text $r.From)] = (Normalize-Text $r.To)
    }
    return $map
}

$AllowedSets = @{
    NYOU     = @('(-)','(+-)','(+)','(2+)','(3+)')
    CHORYOKU = @('所見なし','所見あり')
}

# 問診の選択肢 (KOMOKU_CD + コード → 表示ラベル)。KEKKA=ラベル / KEKKA_CD=コード で書込
function Load-MonshinMap {
    $p = Join-Path $MapDir 'monshin_map.csv'
    $map = @{}
    if (-not (Test-Path $p)) { return $map }
    foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
        $cd = Normalize-Text $r.KOMOKU_CD
        $code = Normalize-Text $r.CD
        if ($cd -eq '' -or $code -eq '') { continue }
        if (-not $map.ContainsKey($cd)) { $map[$cd] = @{} }
        $map[$cd][$code] = Normalize-Text $r.LABEL
    }
    return $map
}

# 戻り値: @{ Status='OK'|'EMPTY'|'CONVERR'; Value=変換後 }
function Convert-FormValue([hashtable]$valueMap, [string]$kind, [string]$raw) {
    $v = Normalize-Text $raw
    if ($v -eq '') { return @{ Status = 'EMPTY' } }
    if ($valueMap.ContainsKey($kind) -and $valueMap[$kind].ContainsKey($v)) {
        $v = $valueMap[$kind][$v]
    }
    if ($AllowedSets.ContainsKey($kind) -and ($AllowedSets[$kind] -notcontains $v)) {
        return @{ Status = 'CONVERR'; Value = $v }
    }
    return @{ Status = 'OK'; Value = $v }
}

# ============================================================================
# DB
# ============================================================================

function Resolve-ConnectionString {
    if ($ConnectionString) { return $ConnectionString }
    if (Test-Path $ConnFile) {
        $txt = Get-Content -Path $ConnFile -Raw
        $lines = @($txt -split "`r?`n" | Where-Object { (Normalize-Text $_) -ne '' })
        # 1) 接続文字列がそのまま書かれている場合
        foreach ($ln in $lines) {
            if ($ln -match '(?i)(data source|server)\s*=') {
                $cs = Normalize-Text $ln
                if ($cs -notmatch '(?i)(initial catalog|database)\s*=') {
                    $cs = $cs.TrimEnd(';') + ';Initial Catalog=K166_SIBAURAUSER'
                }
                return $cs
            }
        }
        # 2) key=value 形式の場合
        $kv = @{}
        foreach ($ln in $lines) {
            if ($ln -match '^\s*([^=:]+?)\s*[=:]\s*(.*?)\s*$') {
                $kv[$Matches[1].ToUpper()] = $Matches[2]
            }
        }
        $srv = $null; $db = $null; $uid = $null; $pwd = $null
        foreach ($k in $kv.Keys) {
            if ($k -match 'SERVER|SRV|DATASOURCE') { $srv = $kv[$k] }
            elseif ($k -match 'DATABASE|CATALOG|DB') { $db = $kv[$k] }
            elseif ($k -match '^(UID|USER|USERID|USER ID)$') { $uid = $kv[$k] }
            elseif ($k -match '^(PWD|PASS|PASSWORD)$') { $pwd = $kv[$k] }
        }
        if ($srv) {
            if (-not $db) { $db = 'K166_SIBAURAUSER' }
            if ($uid) { return "Data Source=$srv;Initial Catalog=$db;User ID=$uid;Password=$pwd" }
            return "Data Source=$srv;Initial Catalog=$db;Integrated Security=True"
        }
        Write-Warning "接続ファイルを解釈できませんでした: $ConnFile"
        Write-Warning "内容(先頭200文字): $($txt.Substring(0, [Math]::Min(200, $txt.Length)))"
        Write-Warning "-ConnectionString で接続文字列を直接指定してください。既定値で続行します。"
    }
    else {
        Write-Warning "接続ファイルが見つかりません: $ConnFile → 既定値(Windows認証)で続行します。"
    }
    return 'Data Source=KNSV\SQLEXPRESS;Initial Catalog=K166_SIBAURAUSER;Integrated Security=True'
}

function Open-Db {
    $cs = Resolve-ConnectionString
    $masked = $cs -replace '(?i)(password|pwd)\s*=\s*[^;]*', '$1=***'
    Write-Host "[DB] 接続先: $masked" -ForegroundColor DarkGray
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    return $conn
}

function Invoke-DbQuery($conn, [string]$sql, [hashtable]$params) {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    if ($params) {
        foreach ($k in $params.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $params[$k]) }
    }
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    $cmd.Dispose()
    return ,$dt
}

function Invoke-DbExec($conn, $tran, [string]$sql, [hashtable]$params) {
    $cmd = $conn.CreateCommand()
    if ($tran) { $cmd.Transaction = $tran }
    $cmd.CommandText = $sql
    if ($params) {
        foreach ($k in $params.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $params[$k]) }
    }
    $n = $cmd.ExecuteNonQuery()
    $cmd.Dispose()
    return $n
}

function Resolve-PkSeq($conn, [string]$ymd, [string]$kenNo) {
    # 1) 受付済み: T_KANJA_G (KEN_YMD + KEN_NO)
    $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KANJA_G WHERE KEN_YMD = @ymd AND KEN_NO = @no' @{ ymd = $ymd; no = $kenNo }
    if ($dt.Rows.Count -eq 0 -and $kenNo -match '^\d+$') {
        $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KANJA_G WHERE KEN_YMD = @ymd AND KEN_NO = @no' @{ ymd = $ymd; no = [int]$kenNo }
    }
    if ($dt.Rows.Count -eq 1) { return $dt.Rows[0].PK_SEQ }
    if ($dt.Rows.Count -gt 1) { throw "受診者が複数見つかりました (KEN_YMD=$ymd, KEN_NO=$kenNo)。中止します。" }
    # 2) 未受付: T_KENSIN (D_KENSIN + UKE_NO_KENSA、予約取消は除外)
    $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KENSIN WHERE D_KENSIN = @ymd AND UKE_NO_KENSA = @no AND F_TORIKESI = 0' @{ ymd = $ymd; no = $kenNo }
    if ($dt.Rows.Count -eq 0 -and $kenNo -match '^\d+$') {
        $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KENSIN WHERE D_KENSIN = @ymd AND UKE_NO_KENSA = @no AND F_TORIKESI = 0' @{ ymd = $ymd; no = [int]$kenNo }
    }
    if ($dt.Rows.Count -eq 0) { return $null }
    if ($dt.Rows.Count -gt 1) { throw "受診者が複数見つかりました (D_KENSIN=$ymd, 受付No=$kenNo)。中止します。" }
    Write-Host "[情報] 未受付のため T_KENSIN から特定しました (受付No=$kenNo)" -ForegroundColor DarkYellow
    return $dt.Rows[0].PK_SEQ
}

# 対象者の現在の T_KENSA 行 (KOMOKU_CD → 行) を取得
# PK_SEQ から健診ナビ側の氏名を引く (取り違え防止の照合に使う)
function Get-NaviName($conn, $pkSeq) {
    $dt = Invoke-DbQuery $conn @'
SELECT TOP 1 k.KANJI_SIMEI, k.KANA_SIMEI
FROM T_KENSIN s LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
WHERE s.PK_SEQ = @p
'@ @{ p = $pkSeq }
    if ($dt.Rows.Count -eq 0) { return $null }
    return @{
        Kanji = Normalize-Text ([string]$dt.Rows[0].KANJI_SIMEI)
        Kana  = Normalize-Text ([string]$dt.Rows[0].KANA_SIMEI)
    }
}

function Get-CurrentKensa($conn, $pkSeq) {
    $dt = Invoke-DbQuery $conn 'SELECT KOMOKU_CD, KEKKA, KEKKA_CD, HANTEI_KIGO FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pkSeq }
    $h = @{}
    foreach ($r in $dt.Rows) { $h[(Normalize-Text ([string]$r.KOMOKU_CD))] = $r }
    return $h
}

$script:SyokenCache = @{}
function Get-SyokenTable($conn, [string]$cd) {
    if (-not $script:SyokenCache.ContainsKey($cd)) {
        $script:SyokenCache[$cd] = Invoke-DbQuery $conn 'SELECT KEKKA_CD, SYOKEN, HANTEI_KIGO FROM T_SYOKEN2 WHERE SYOKEN_CD = @cd' @{ cd = $cd }
    }
    return $script:SyokenCache[$cd]
}

# 結果CD(KEKKA_CD)でマスタを引く (SHOKENCD/SHOKENCD2/MONSHIN用)
# DBへ直接問い合わせる (前後空白・ゼロ埋めの差を吸収)
function Find-SyokenCd($conn, [string]$cd, [string]$code) {
    $c = Normalize-Text $code
    if ($cd -eq '' -or $c -eq '') { return $null }
    # 1) 文字列として一致 (前後空白は無視)
    $dt = Invoke-DbQuery $conn @'
SELECT TOP 1 KEKKA_CD, SYOKEN, HANTEI_KIGO FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) = @cd AND LTRIM(RTRIM(KEKKA_CD)) = @kc
'@ @{ cd = $cd; kc = $c }
    if ($dt.Rows.Count -gt 0) { return $dt.Rows[0] }
    # 2) 数値として一致 ('1' と '01' / '040' と '40' の差を吸収)
    if ($c -match '^\d+$') {
        $dt = Invoke-DbQuery $conn @'
SELECT TOP 1 KEKKA_CD, SYOKEN, HANTEI_KIGO FROM T_SYOKEN2
WHERE LTRIM(RTRIM(SYOKEN_CD)) = @cd AND ISNUMERIC(KEKKA_CD) = 1
  AND CAST(LTRIM(RTRIM(KEKKA_CD)) AS int) = @n
'@ @{ cd = $cd; n = [int]$c }
        if ($dt.Rows.Count -gt 0) { return $dt.Rows[0] }
    }
    return $null
}

# 項目マスタ(T_KOMOKU)から選択肢リスト名(SYOKEN_CD)を取得
$script:ItemSyokenCache = @{}
function Get-ItemSyokenCd($conn, [string]$komoku) {
    if (-not $script:ItemSyokenCache.ContainsKey($komoku)) {
        $dt = Invoke-DbQuery $conn 'SELECT SYOKEN_CD FROM T_KOMOKU WHERE KOMOKU_CD = @cd' @{ cd = $komoku }
        $v = ''
        if ($dt.Rows.Count -gt 0) { $v = Normalize-Text ([string]$dt.Rows[0].SYOKEN_CD) }
        $script:ItemSyokenCache[$komoku] = $v
    }
    return $script:ItemSyokenCache[$komoku]
}

# 所見文をマスタと突合 (完全一致 → 空白無視一致)
function Find-Syoken($conn, [string]$cd, [string]$text) {
    $t = Normalize-Text $text
    if ($cd -eq '' -or $t -eq '') { return $null }
    $dt = Invoke-DbQuery $conn 'SELECT KEKKA_CD, SYOKEN, HANTEI_KIGO FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) = @cd' @{ cd = $cd }
    foreach ($r in $dt.Rows) {
        if ((Normalize-Text ([string]$r.SYOKEN)) -eq $t) { return $r }
    }
    $t2 = $t -replace '[\s　]', ''
    foreach ($r in $dt.Rows) {
        if ((([string]$r.SYOKEN) -replace '[\s　]', '') -eq $t2) { return $r }
    }
    return $null
}

# ============================================================================
# CSV行の選択 (受付番号・受診日の解決)
# ============================================================================

function Get-IdColumns($mapRows) {
    $kenNoCol = 0; $ymdCol = 0; $kanjiCol = 0; $kanaCol = 0
    foreach ($m in $mapRows) {
        $kind = (Normalize-Text $m.Kind).ToUpper()
        $c = 0
        [void][int]::TryParse((Normalize-Text $m.Col), [ref]$c)
        if ($kind -eq 'KENNO')     { $kenNoCol = $c }
        if ($kind -eq 'KENYMD')    { $ymdCol = $c }
        if ($kind -eq 'NAMEKANJI') { $kanjiCol = $c }
        if ($kind -eq 'NAMEKANA')  { $kanaCol = $c }
    }
    return @{ KenNo = $kenNoCol; Ymd = $ymdCol; Kanji = $kanjiCol; Kana = $kanaCol }
}

# 氏名の照合用に正規化 (空白・記号を除去)
function Normalize-Name([string]$s) {
    $v = Normalize-Text $s
    $v = $v -replace '[\s　・･,、]', ''
    return $v
}

function Select-TargetRows($dataRows, $idCols) {
    if (-not $Only) { return ,$dataRows }
    if ($idCols.KenNo -le 0) {
        throw "mapping.csv の KENNO 行に列番号(Col)が設定されていません。まず -Inspect で受付番号の列番号を確認し、mapping.csv に記入してください。"
    }
    # カンマ区切りで複数指定可 (例: -Only 4005,4009)
    $wants = @($Only -split '[,、]' | ForEach-Object { Normalize-KenNo $_ } | Where-Object { $_ -ne '' })
    $sel = @()
    foreach ($r in $dataRows) {
        if ($wants -contains (Normalize-KenNo (Get-Field $r $idCols.KenNo))) { $sel += ,$r }
    }
    if ($sel.Count -eq 0) { throw "受付番号 $Only の行がCSVに見つかりません。" }
    $found = @($sel | ForEach-Object { Normalize-KenNo (Get-Field $_ $idCols.KenNo) })
    $missing = @($wants | Where-Object { $found -notcontains $_ })
    if ($missing.Count -gt 0) { Write-Warning ("CSVに見つからない受付番号: " + ($missing -join ', ')) }
    return ,$sel
}

$script:YmdNoticeShown = $false

function Resolve-RowYmd($fields, $idCols) {
    # 受診日を明示指定したときは、ファイルの日付列よりそちらを優先する
    # (テストで別日の枠に入れたい場合など。通常は空欄にしてファイルの日付を使う)
    if ($KenYmd) {
        $y = Normalize-Ymd $KenYmd
        if (-not $y) { throw ("受診日「{0}」を解釈できません。2026/07/12 のように入力してください。" -f $KenYmd) }
        if (-not $script:YmdNoticeShown) {
            $script:YmdNoticeShown = $true
            if ($idCols.Ymd -gt 0) {
                $fileYmd = Normalize-Ymd (Get-Field $fields $idCols.Ymd)
                if ($fileYmd -and $fileYmd -ne $y) {
                    Write-Host ("[受診日] 指定された {0} を使います (ファイルの日付 {1} は使いません)" -f $y, $fileYmd) -ForegroundColor Yellow
                }
            }
        }
        return $y
    }
    if ($idCols.Ymd -gt 0) {
        $y = Normalize-Ymd (Get-Field $fields $idCols.Ymd)
        if ($y) { return $y }
    }
    throw "受診日が特定できません。ファイルに日付の列が無い場合は「受診日」に 2026/07/12 のように入力してください。"
}

# ============================================================================
# 名簿によるしぼり込み (-Roster)
# ============================================================================
# 血液などの外部データには名簿外の人が混ざってくるため、
# 名簿(リアン等)に載っている人だけを取り込めるようにする。
# 人の同定は最終的に PK_SEQ で行うので、名簿と結果ファイルのキーが違っても照合できる。


# 名簿ファイルから「取込を許可する人」の PK_SEQ 一覧を作る
function Load-RosterPkSeq($conn, [string]$path) {
    if (-not (Test-Path $path)) { throw "名簿ファイルが見つかりません: $path" }
    $p = $path
    $mapPath = Detect-MappingPath $p
    if (-not $mapPath) { throw "名簿のレイアウトを判別できませんでした: $path" }
    Write-Host "[名簿] $([System.IO.Path]::GetFileName($path)) / 対応表 $([System.IO.Path]::GetFileName($mapPath))" -ForegroundColor DarkGray
    $rmap = Import-Csv -Path $mapPath -Encoding UTF8
    $rid  = Get-IdColumns $rmap
    if ($rid.KenNo -le 0) { throw "名簿の対応表に受付番号(KENNO)の列がありません: $mapPath" }

    $rrows = Read-FormCsv $p $CsvEncoding
    $rhd   = Split-HeaderData $rrows $rid

    $set = @{}
    $miss = @()
    foreach ($f in $rhd.Data) {
        $no = Normalize-KenNo (Get-Field $f $rid.KenNo)
        if ($no -eq '') { continue }
        $ymd = $null
        if ($rid.Ymd -gt 0) { $ymd = Normalize-Ymd (Get-Field $f $rid.Ymd) }
        if (-not $ymd -and $KenYmd) { $ymd = Normalize-Ymd $KenYmd }
        if (-not $ymd) { throw "名簿に受診日の列がありません。-KenYmd で受診日を指定してください: $path" }
        $pk = Resolve-PkSeq $conn $ymd $no
        if ($null -eq $pk) { $miss += $no; continue }
        $set[[string]$pk] = $no
    }
    if ($miss.Count -gt 0) {
        Write-Warning ("名簿にあるが健診ナビで見つからない受付番号 ({0}件): {1}" -f $miss.Count, ($miss -join ', '))
    }
    Write-Host ("[名簿] 取込対象 {0} 人" -f $set.Count) -ForegroundColor DarkGray
    if ($set.Count -eq 0) { throw "名簿から取込対象を1人も特定できませんでした: $path" }
    return $set
}

# ============================================================================
# 取込プラン作成 (1人分)
# ============================================================================
# 戻り値: レポート行の配列。書込対象は .Update に UPDATE 内容を持つ。

function Build-Plan($conn, $mapRows, $valueMap, $fields, $current) {
    $plan = @()
    foreach ($m in $mapRows) {
        $kind = (Normalize-Text $m.Kind).ToUpper()
        if ($kind -eq '' -or $kind -eq 'KENNO' -or $kind -eq 'KENYMD' -or $kind -eq 'IGNORE') { continue }

        $col = 0
        [void][int]::TryParse((Normalize-Text $m.Col), [ref]$col)
        $komoku = Normalize-Text $m.KOMOKU_CD
        $label  = Normalize-Text $m.Label

        $rep = New-Object PSObject -Property @{
            Col = $(if ($col -gt 0) { $col } else { '-' })
            Label = $label; KomokuCd = $komoku
            Now = ''; New = ''; KekkaCd = ''; Hantei = ''
            Status = ''; Update = $null
        }

        if ($col -le 0) { $rep.Status = '列未設定'; $plan += $rep; continue }
        $raw = Get-Field $fields $col

        # 所見欄が空でも、検査を実施していれば既定コード(異常なし)を書く
        # IfEmpty = 書き込む結果CD / ReqCol = 実施を示す列(この列が空なら何もしない)
        $ifEmpty = Normalize-Text $m.IfEmpty
        if ($ifEmpty -ne '' -and (Normalize-Text $raw) -eq '') {
            $reqCol = 0
            [void][int]::TryParse((Normalize-Text $m.ReqCol), [ref]$reqCol)
            $done = $true
            if ($reqCol -gt 0) { $done = (Normalize-Text (Get-Field $fields $reqCol)) -ne '' }
            if (-not $done) { continue }   # 検査未実施 → 何も書かない
            $raw = $ifEmpty
            $rep.Label = $label + '(所見なし)'
        }

        if ($kind -eq 'SHOKENCD' -or $kind -eq 'SHOKENCD2') {
            # ---- 所見を結果CD(コード)で直接指定する形式 ----
            $shoCode = (Normalize-Text $raw).ToUpper()
            if ($shoCode -ne '') {
                $cv = Convert-Code $m $shoCode
                if ($cv.Status -eq 'DROP') { continue }
                if ($cv.Status -eq 'NOMAP') {
                    $rep.Status = "変換表に無いコード($shoCode)"; $rep.New = $shoCode; $plan += $rep; continue
                }
                $shoCode = $cv.Code
            }
            $buiCode = ''
            if ($kind -eq 'SHOKENCD2') {
                $col2 = 0
                [void][int]::TryParse((Normalize-Text $m.Col2), [ref]$col2)
                if ($col2 -gt 0) { $buiCode = (Normalize-Text (Get-Field $fields $col2)).ToUpper() }
            }
            if ($shoCode -eq '' -and $buiCode -eq '') { continue }
            if ($komoku -eq '') { $rep.Status = '項目CD未設定'; $rep.New = ($buiCode + ' ' + $shoCode).Trim(); $plan += $rep; continue }

            if ($kind -eq 'SHOKENCD') {
                $cd = Normalize-Text $m.SYOKEN_CD
                $hit = Find-SyokenCd $conn $cd $shoCode
                if (-not $hit) { $rep.Status = "所見CD未登録(${cd}:${shoCode})"; $rep.New = $shoCode; $plan += $rep; continue }
                $rep.New = [string]$hit.SYOKEN
                $rep.KekkaCd = Normalize-Text ([string]$hit.KEKKA_CD)
                $rep.Hantei  = Normalize-Text ([string]$hit.HANTEI_KIGO)
            }
            else {
                # SYOKEN_CD=部位マスタ / SYOKEN_CD2=所見マスタ (どちらもコード指定)
                $cdBui = Normalize-Text $m.SYOKEN_CD
                $cdSho = Normalize-Text $m.SYOKEN_CD2
                $hitB = $null
                if ($buiCode -ne '') {
                    $hitB = Find-SyokenCd $conn $cdBui $buiCode
                    if (-not $hitB) { $rep.Status = "所見CD未登録(${cdBui}:部位${buiCode})"; $rep.New = $buiCode; $plan += $rep; continue }
                }
                if ($shoCode -eq '') { $rep.Status = '所見CDが空(部位のみ)'; $rep.New = $buiCode; $plan += $rep; continue }
                $hitS = Find-SyokenCd $conn $cdSho $shoCode
                if (-not $hitS) { $rep.Status = "所見CD未登録(${cdSho}:${shoCode})"; $rep.New = $shoCode; $plan += $rep; continue }
                $buiStr = ''
                if ($hitB) { $buiStr = [string]$hitB.SYOKEN }
                $rep.New = (($buiStr + ' ' + [string]$hitS.SYOKEN).Trim())
                $rep.KekkaCd = Normalize-Text ([string]$hitS.KEKKA_CD)
                $rep.Hantei  = Normalize-Text ([string]$hitS.HANTEI_KIGO)
            }

            if (-not $current.ContainsKey($komoku)) { $rep.Status = '枠なし'; $plan += $rep; continue }
            $rep.Now = Normalize-Text ([string]$current[$komoku].KEKKA)
            $rep.Status = 'OK'
            $rep.Update = @{
                Sql = 'UPDATE T_KENSA SET KEKKA = @k, KEKKA_CD = @kc, HANTEI_KIGO = @h WHERE PK_SEQ = @p AND KOMOKU_CD = @cd'
                P   = @{ k = $rep.New; kc = $rep.KekkaCd; h = $rep.Hantei; cd = $komoku }
            }
            $plan += $rep
        }
        elseif ($kind -eq 'SHOKEN' -or $kind -eq 'SHOKEN2') {
            # ---- 所見系 ----
            $shoText = Normalize-Text $raw
            $buiText = ''
            if ($kind -eq 'SHOKEN2') {
                $col2 = 0
                [void][int]::TryParse((Normalize-Text $m.Col2), [ref]$col2)
                if ($col2 -gt 0) { $buiText = Normalize-Text (Get-Field $fields $col2) }
            }
            if ($shoText -eq '' -and $buiText -eq '') { continue }  # 空欄はスキップ(表示もしない)

            if ($komoku -eq '') { $rep.Status = '項目CD未設定'; $rep.New = ($buiText + ' ' + $shoText).Trim(); $plan += $rep; continue }

            if ($kind -eq 'SHOKEN') {
                $cd = Normalize-Text $m.SYOKEN_CD
                $hit = Find-Syoken $conn $cd $shoText
                if (-not $hit) { $rep.Status = "所見未登録($cd)"; $rep.New = $shoText; $plan += $rep; continue }
                $rep.New = [string]$hit.SYOKEN
                $rep.KekkaCd = Normalize-Text ([string]$hit.KEKKA_CD)
                $rep.Hantei  = Normalize-Text ([string]$hit.HANTEI_KIGO)
            }
            else {
                # 2段階: SYOKEN_CD=部位マスタ / SYOKEN_CD2=所見マスタ
                $cdBui = Normalize-Text $m.SYOKEN_CD
                $cdSho = Normalize-Text $m.SYOKEN_CD2
                $hitB = $null
                if ($buiText -ne '') {
                    $hitB = Find-Syoken $conn $cdBui $buiText
                    if (-not $hitB) { $rep.Status = "所見未登録(${cdBui}:部位)"; $rep.New = $buiText; $plan += $rep; continue }
                }
                $hitS = Find-Syoken $conn $cdSho $shoText
                if (-not $hitS) { $rep.Status = "所見未登録($cdSho)"; $rep.New = $shoText; $plan += $rep; continue }
                $buiStr = ''
                if ($hitB) { $buiStr = [string]$hitB.SYOKEN }
                $rep.New = (($buiStr + ' ' + [string]$hitS.SYOKEN).Trim())
                $rep.KekkaCd = Normalize-Text ([string]$hitS.KEKKA_CD)
                $rep.Hantei  = Normalize-Text ([string]$hitS.HANTEI_KIGO)
            }

            if (-not $current.ContainsKey($komoku)) { $rep.Status = '枠なし'; $plan += $rep; continue }
            $rep.Now = Normalize-Text ([string]$current[$komoku].KEKKA)
            $rep.Status = 'OK'
            $rep.Update = @{
                Sql = 'UPDATE T_KENSA SET KEKKA = @k, KEKKA_CD = @kc, HANTEI_KIGO = @h WHERE PK_SEQ = @p AND KOMOKU_CD = @cd'
                P   = @{ k = $rep.New; kc = $rep.KekkaCd; h = $rep.Hantei; cd = $komoku }
            }
            $plan += $rep
        }
        elseif ($kind -eq 'VISION') {
            # ---- 視力 (種別列で 裸眼/矯正 を振り分け) ----
            # Col=値の列 / Col2=種別の列 / KOMOKU_CD=裸眼の項目 / SYOKEN_CD=矯正の項目
            $val = Normalize-Text $raw
            if ($val -eq '') { continue }
            $col2 = 0
            [void][int]::TryParse((Normalize-Text $m.Col2), [ref]$col2)
            $shubetsu = if ($col2 -gt 0) { Normalize-KenNo (Get-Field $fields $col2) } else { '' }
            $target = '裸眼'
            if ($shubetsu -ne '') {
                if ($script:VisionMap.ContainsKey($shubetsu)) { $target = $script:VisionMap[$shubetsu] }
                else { $rep.Status = "視力の種別コード($shubetsu)が未設定"; $rep.New = $val; $plan += $rep; continue }
            }
            $komoku = if ($target -eq '矯正') { Normalize-Text $m.SYOKEN_CD } else { Normalize-Text $m.KOMOKU_CD }
            $rep.Label = $label + "($target)"
            $rep.KomokuCd = $komoku
            if ($komoku -eq '') { $rep.Status = '項目CD未設定'; $rep.New = $val; $plan += $rep; continue }
            $fmt = Normalize-Text $m.Format
            if ($fmt -ne '') { $d = 0.0; if ([double]::TryParse($val, [ref]$d)) { $val = $d.ToString($fmt) } }
            $rep.New = $val
            if (-not $current.ContainsKey($komoku)) { $rep.Status = '枠なし'; $plan += $rep; continue }
            $rep.Now = Normalize-Text ([string]$current[$komoku].KEKKA)
            $rep.Status = 'OK'
            $rep.Update = @{
                Sql = 'UPDATE T_KENSA SET KEKKA = @k WHERE PK_SEQ = @p AND KOMOKU_CD = @cd'
                P   = @{ k = $val; cd = $komoku }
            }
            $plan += $rep
        }
        elseif ($kind -eq 'MONSHIN') {
            # ---- 選択式項目 (問診・尿・聴力等): コード → ラベル+判定 ----
            # T_KOMOKU.SYOKEN_CD → T_SYOKEN2 の選択肢リストで解決 (無ければ monshin_map.csv)
            $rawCode = Normalize-KenNo (Get-Field $fields $col)
            if ($rawCode -eq '') { continue }
            if ($komoku -eq '') { $rep.Status = '項目CD未設定'; $rep.New = $rawCode; $plan += $rep; continue }
            $cv = Convert-Code $m $rawCode
            if ($cv.Status -eq 'DROP') { continue }   # 変換表で「取り込まない」と定義済み
            if ($cv.Status -eq 'NOMAP') {
                $rep.Status = "変換表に無いコード($rawCode)"; $rep.New = $rawCode; $plan += $rep; continue
            }
            $code = $cv.Code
            $label = $null
            $hanteiK = ''
            $scd = Get-ItemSyokenCd $conn $komoku
            if ($scd -ne '') {
                $hit = Find-SyokenCd $conn $scd $code
                if ($hit) {
                    $label = [string]$hit.SYOKEN
                    $hanteiK = Normalize-Text ([string]$hit.HANTEI_KIGO)
                }
            }
            if ($null -eq $label) {
                $mm = $script:MonshinMap
                if ($mm.ContainsKey($komoku) -and $mm[$komoku].ContainsKey($code)) { $label = $mm[$komoku][$code] }
            }
            if ($null -eq $label) {
                $rep.Status = "選択肢未登録(コード$code" + $(if ($scd) { "/リスト$scd" } else { '' }) + ')'
                $rep.New = $code; $plan += $rep; continue
            }
            if (-not $current.ContainsKey($komoku)) { $rep.Status = '枠なし'; $rep.New = $label; $plan += $rep; continue }
            $rep.Now = Normalize-Text ([string]$current[$komoku].KEKKA)
            $rep.New = $label
            $rep.KekkaCd = $code
            $rep.Hantei = $hanteiK
            $rep.Status = 'OK'
            if ($hanteiK -ne '') {
                $rep.Update = @{
                    Sql = 'UPDATE T_KENSA SET KEKKA = @k, KEKKA_CD = @kc, HANTEI_KIGO = @h WHERE PK_SEQ = @p AND KOMOKU_CD = @cd'
                    P   = @{ k = $label; kc = $code; h = $hanteiK; cd = $komoku }
                }
            } else {
                $rep.Update = @{
                    Sql = 'UPDATE T_KENSA SET KEKKA = @k, KEKKA_CD = @kc WHERE PK_SEQ = @p AND KOMOKU_CD = @cd'
                    P   = @{ k = $label; kc = $code; cd = $komoku }
                }
            }
            $plan += $rep
        }
        else {
            # ---- 値系 (VALUE / NYOU / CHORYOKU) ----
            $conv = Convert-FormValue $valueMap $kind $raw
            if ($conv.Status -eq 'EMPTY') { continue }  # 空欄はスキップ
            if ($komoku -eq '') { $rep.Status = '項目CD未設定'; $rep.New = (Normalize-Text $raw); $plan += $rep; continue }
            if ($conv.Status -eq 'CONVERR') { $rep.Status = "変換不可($kind)"; $rep.New = $conv.Value; $plan += $rep; continue }
            $rep.New = $conv.Value
            # 表記を健診ナビに合わせる (Format=0.0 なら 147 → 147.0)
            $fmt = Normalize-Text $m.Format
            if ($fmt -ne '') {
                $d = 0.0
                if ([double]::TryParse($rep.New, [ref]$d)) { $rep.New = $d.ToString($fmt) }
            }
            if (-not $current.ContainsKey($komoku)) { $rep.Status = '枠なし'; $plan += $rep; continue }
            $rep.Now = Normalize-Text ([string]$current[$komoku].KEKKA)
            $rep.Status = 'OK'
            $rep.Update = @{
                Sql = 'UPDATE T_KENSA SET KEKKA = @k WHERE PK_SEQ = @p AND KOMOKU_CD = @cd'
                P   = @{ k = $rep.New; cd = $komoku }
            }
            $plan += $rep
        }
    }
    return ,$plan
}

function Show-Plan($plan, [string]$who) {
    Write-Host ''
    Write-Host "=== 取込プレビュー: $who ===" -ForegroundColor Cyan
    if ($plan.Count -eq 0) { Write-Host '(取込対象データなし)'; return }
    $plan |
        Select-Object @{n='列';e={$_.Col}}, @{n='項目';e={$_.Label}}, @{n='KOMOKU_CD';e={$_.KomokuCd}},
                      @{n='現在値';e={$_.Now}}, @{n='新しい値';e={$_.New}},
                      @{n='KEKKA_CD';e={$_.KekkaCd}}, @{n='判定';e={$_.Hantei}}, @{n='状態';e={$_.Status}} |
        Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host
    $ok   = @($plan | Where-Object { $_.Status -eq 'OK' })
    $warn = @($plan | Where-Object { $_.Status -eq '列未設定' -or $_.Status -eq '項目CD未設定' })
    $err  = @($plan | Where-Object { $_.Status -ne 'OK' -and $_.Status -ne '列未設定' -and $_.Status -ne '項目CD未設定' })
    Write-Host ("書込可能: {0} 件 / 対応表未設定(スキップ): {1} 件 / エラー: {2} 件" -f $ok.Count, $warn.Count, $err.Count) `
        -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })
}

# 結果入力画面で編集中(ロック中)かどうか。ロック中の書込は画面側の登録で上書きされる危険がある
function Test-Locked($conn, $pkSeq) {
    try {
        $dt = Invoke-DbQuery $conn 'SELECT PC_NAME, USER_ID, LOCKED_DATE FROM T_MULTI WHERE PK_SEQ = @p' @{ p = $pkSeq }
        if ($dt.Rows.Count -gt 0) {
            $r = $dt.Rows[0]
            return "$($r.PC_NAME) / $($r.USER_ID) / $($r.LOCKED_DATE)"
        }
    } catch { }
    return $null
}

function Backup-Kensa($conn, $pkSeq) {
    if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
    $dt = Invoke-DbQuery $conn 'SELECT * FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pkSeq }
    $file = Join-Path $BackupDir ("T_KENSA_{0}_{1}.csv" -f $pkSeq, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $dt | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Host "[バックアップ] 書込前の T_KENSA を保存: $file" -ForegroundColor DarkGray
    return $file
}

function Commit-Plan($conn, $pkSeq, $plan) {
    $targets = @($plan | Where-Object { $_.Status -eq 'OK' -and $_.Update })
    if ($targets.Count -eq 0) { Write-Host '書込対象がありません。'; return }
    $lock = Test-Locked $conn $pkSeq
    if ($lock) {
        throw "この受診者は健診ナビの結果入力画面で編集中です ($lock)。画面を閉じてから再実行してください。"
    }
    [void](Backup-Kensa $conn $pkSeq)
    $tran = $conn.BeginTransaction()
    try {
        $done = 0
        foreach ($t in $targets) {
            $p = @{}
            foreach ($k in $t.Update.P.Keys) { $p[$k] = $t.Update.P[$k] }
            $p['p'] = $pkSeq
            $n = Invoke-DbExec $conn $tran $t.Update.Sql $p
            if ($n -ne 1) {
                throw ("UPDATE影響行数が {0} でした (KOMOKU_CD={1})。ロールバックします。" -f $n, $t.KomokuCd)
            }
            $done++
        }
        $tran.Commit()
        Write-Host ("[書込完了] {0} 項目を更新しました (PK_SEQ={1})" -f $done, $pkSeq) -ForegroundColor Green
        Write-Host '※ 健診ナビで対象者を開き「自動判定」を実行してください (BMI・判定の計算)。'
    }
    catch {
        $tran.Rollback()
        throw
    }
}

# ============================================================================
# メイン
# ============================================================================

# ---- モード: 所見マスタ一覧 ----
if ($DumpSyoken) {
    $conn = Open-Db
    try {
        $dt = Invoke-DbQuery $conn 'SELECT SYOKEN_CD, KEKKA_CD, SYOKEN, HANTEI_KIGO FROM T_SYOKEN2 WHERE SYOKEN_CD = @cd ORDER BY KEKKA_CD' @{ cd = $DumpSyoken }
        Write-Host "=== T_SYOKEN2 (SYOKEN_CD=$DumpSyoken) : $($dt.Rows.Count) 件 ===" -ForegroundColor Cyan
        $dt | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    }
    finally { $conn.Close() }
    return
}

$mapRows  = Load-Mapping
$valueMap = Load-ValueMap
$script:MonshinMap = Load-MonshinMap
$script:CodeMap = Load-CodeMap
$script:VisionMap = @{}
$vp = Join-Path $MapDir 'vision_map.csv'
if (Test-Path $vp) {
    foreach ($r in (Import-Csv -Path $vp -Encoding UTF8)) {
        $c = Normalize-Text $r.Code
        if ($c -ne '') { $script:VisionMap[$c] = Normalize-Text $r.Target }
    }
}
$idCols   = Get-IdColumns $mapRows

# ---- モード: 対象者の T_KENSA 一覧 ----
if ($DumpItems) {
    # 受付番号・受診日が未指定なら、選んだファイルの先頭の人で確認する
    $dumpNo = Normalize-KenNo $Only
    $ymd = $null
    if ($Csv) {
        $rows = Read-FormCsv $Csv $CsvEncoding
        $hd = Split-HeaderData $rows $idCols
        $sel = $hd.Data
        if ($dumpNo -ne '') {
            $sel = Select-TargetRows $hd.Data $idCols
        }
        elseif ($idCols.KenNo -gt 0) {
            $dumpNo = Normalize-KenNo (Get-Field $sel[0] $idCols.KenNo)
            if ($dumpNo -ne '') {
                Write-Host ("[受付番号] 未入力のため、ファイルの先頭 {0} で確認します" -f $dumpNo) -ForegroundColor DarkGray
            }
        }
        $ymd = Resolve-RowYmd $sel[0] $idCols
    }
    else {
        if (-not $KenYmd) { throw '受診日を入力するか、ファイルを選んでください。' }
        $ymd = Normalize-Ymd $KenYmd
        if (-not $ymd) { throw ("受診日「{0}」を解釈できません。2026/07/05 のように入力してください。" -f $KenYmd) }
    }
    if ($dumpNo -eq '') { throw '受付番号を入力してください。' }
    $Only = $dumpNo
    $conn = Open-Db
    try {
        $pk = Resolve-PkSeq $conn $ymd (Normalize-KenNo $Only)
        if ($null -eq $pk) { throw "受診者が見つかりません (KEN_YMD=$ymd, KEN_NO=$Only)" }
        Write-Host "PK_SEQ = $pk (KEN_YMD=$ymd, KEN_NO=$Only)" -ForegroundColor Cyan
        $dt = Invoke-DbQuery $conn 'SELECT KOMOKU_CD, KEKKA, KEKKA_CD, HANTEI_KIGO FROM T_KENSA WHERE PK_SEQ = @p ORDER BY KOMOKU_CD' @{ p = $pk }
        $dt | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
        if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
        $full = Invoke-DbQuery $conn 'SELECT * FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pk }
        $file = Join-Path $BackupDir ("dump_T_KENSA_{0}_{1}.csv" -f $pk, (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $full | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
        Write-Host "全列ダンプ: $file" -ForegroundColor DarkGray
    }
    finally { $conn.Close() }
    return
}

if (-not $Csv) { throw '使い方: form_import.ps1 -Csv <form.csv> [-Only 4001] [-KenYmd 2026/07/02] [-Commit] / -Inspect / -DumpItems / -DumpSyoken <CD>' }

$rows = Read-FormCsv $Csv $CsvEncoding
$hd = Split-HeaderData $rows $idCols
$header = $hd.Header
$data = $hd.Data

# ---- モード: 受診日の確認だけ (DB接続なし・GUIが上書き確認に使う) ----
if ($ShowYmd) {
    $fileYmd = ''
    if ($idCols.Ymd -gt 0) {
        foreach ($f in $data) {
            $y = Normalize-Ymd (Get-Field $f $idCols.Ymd)
            if ($y) { $fileYmd = $y; break }
        }
    }
    $inYmd = ''
    if ($KenYmd) {
        $n = Normalize-Ymd $KenYmd
        if ($n) { $inYmd = $n }
    }
    Write-Output ("FILEYMD=" + $fileYmd)
    Write-Output ("INPUTYMD=" + $inYmd)
    return
}

# ---- モード: CSV列の確認 (DB接続なし) ----
if ($Inspect) {
    $target = $data
    if ($Only -and $idCols.KenNo -gt 0) { $target = Select-TargetRows $data $idCols }
    elseif ($Only) { Write-Warning 'KENNO列が未設定のため -Only は無視し、先頭行を表示します。' }
    $fields = $target[0]
    Write-Host "=== CSV列一覧 (全 $($header.Count) 列 / データ $($data.Count) 行) ===" -ForegroundColor Cyan
    $items = @()
    for ($i = 0; $i -lt $header.Count; $i++) {
        $col = $i + 1
        $mapped = ''
        foreach ($m in $mapRows) {
            $c = 0;  [void][int]::TryParse((Normalize-Text $m.Col),  [ref]$c)
            $c2 = 0; [void][int]::TryParse((Normalize-Text $m.Col2), [ref]$c2)
            if ($c -eq $col -or $c2 -eq $col) { $mapped = "$($m.Label) [$($m.Kind)]"; break }
        }
        $conv = ''
        if ($mapped -ne '') {
            foreach ($m in $mapRows) {
                $c = 0; [void][int]::TryParse((Normalize-Text $m.Col), [ref]$c)
                if ($c -ne $col) { continue }
                $kind = (Normalize-Text $m.Kind).ToUpper()
                if ($kind -eq 'NYOU' -or $kind -eq 'CHORYOKU' -or $kind -eq 'VALUE') {
                    $r = Convert-FormValue $valueMap $kind (Get-Field $fields $col)
                    if ($r.Status -eq 'OK') { $conv = $r.Value }
                    elseif ($r.Status -eq 'CONVERR') { $conv = "!変換不可: $($r.Value)" }
                }
                break
            }
        }
        $items += New-Object PSObject -Property @{
            '列' = $col; 'ヘッダ' = $header[$i]; '値' = (Get-Field $fields $col); '変換後' = $conv; 'マッピング' = $mapped
        }
    }
    $items | Select-Object 列, ヘッダ, 値, 変換後, マッピング | Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host
    return
}

# ---- モード: 受付番号を健診ナビへ設定 (氏名で照合) ----
if ($SetUkeNo) {
    if ($idCols.Kanji -le 0 -and $idCols.Kana -le 0) {
        throw 'mapping.csv に NAMEKANJI または NAMEKANA の列(氏名)を設定してください。'
    }
    if ($idCols.KenNo -le 0) { throw 'mapping.csv の KENNO 行に列番号を設定してください。' }
    $rowsSel = if ($Only) { Select-TargetRows $data $idCols } else { $data }
    $conn = Open-Db
    try {
        $cache = @{}
        $plan = @()
        foreach ($fields in $rowsSel) {
            $kenNo = Normalize-KenNo (Get-Field $fields $idCols.KenNo)
            if ($kenNo -eq '') { continue }
            $ymd = Resolve-RowYmd $fields $idCols
            if (-not $cache.ContainsKey($ymd)) {
                $cache[$ymd] = Invoke-DbQuery $conn @'
SELECT s.PK_SEQ, s.UKE_NO_KENSA, k.KANJI_SIMEI, k.KANA_SIMEI
FROM T_KENSIN s LEFT JOIN T_KOJIN1 k ON k.KOJIN_ID = s.KOJIN_ID
WHERE s.D_KENSIN = @ymd AND s.F_TORIKESI = 0
'@ @{ ymd = $ymd }
            }
            $navi = $cache[$ymd].Rows
            $kanji = Normalize-Name (Get-Field $fields $idCols.Kanji)
            $kana  = Normalize-Name (Get-Field $fields $idCols.Kana)

            $hits = @()
            if ($kana -ne '')  { $hits = @($navi | Where-Object { (Normalize-Name ([string]$_.KANA_SIMEI)) -eq $kana }) }
            if ($hits.Count -eq 0 -and $kanji -ne '') {
                $hits = @($navi | Where-Object { (Normalize-Name ([string]$_.KANJI_SIMEI)) -eq $kanji })
            }
            elseif ($hits.Count -gt 1 -and $kanji -ne '') {
                $narrow = @($hits | Where-Object { (Normalize-Name ([string]$_.KANJI_SIMEI)) -eq $kanji })
                if ($narrow.Count -eq 1) { $hits = $narrow }
            }

            $rep = New-Object PSObject -Property @{
                受付番号 = $kenNo; 受診日 = $ymd
                氏名 = (Normalize-Text (Get-Field $fields $idCols.Kanji))
                現在の番号 = ''; PkSeq = $null; 状態 = ''
            }
            if ($hits.Count -eq 0)     { $rep.状態 = '該当者なし'; $plan += $rep; continue }
            if ($hits.Count -gt 1)     { $rep.状態 = "同名が{0}人いて特定できません" -f $hits.Count; $plan += $rep; continue }

            $hit = $hits[0]
            $cur = Normalize-KenNo ([string]$hit.UKE_NO_KENSA)
            $rep.現在の番号 = $cur
            $rep.PkSeq = $hit.PK_SEQ
            # その日の他の人に同じ番号が付いていないか
            $dup = @($navi | Where-Object {
                (Normalize-KenNo ([string]$_.UKE_NO_KENSA)) -eq $kenNo -and $_.PK_SEQ -ne $hit.PK_SEQ })
            if ($dup.Count -gt 0) { $rep.状態 = '同じ受付番号が別の人に設定済み'; $plan += $rep; continue }

            if ($cur -eq $kenNo)   { $rep.状態 = '設定済み(変更なし)' }
            elseif ($cur -eq '')   { $rep.状態 = 'OK' }
            elseif ($Overwrite)    { $rep.状態 = 'OK(上書き)' }
            else                   { $rep.状態 = "別の番号が設定済み($cur)" }
            $plan += $rep
        }

        Write-Host ''
        Write-Host '=== 受付番号の設定プレビュー ===' -ForegroundColor Cyan
        $plan | Select-Object 受付番号, 受診日, 氏名, 現在の番号, 状態 |
            Format-Table -AutoSize -Wrap | Out-String -Width 200 | Write-Host
        $ok  = @($plan | Where-Object { $_.状態 -like 'OK*' })
        $skip = @($plan | Where-Object { $_.状態 -eq '設定済み(変更なし)' })
        $err = @($plan | Where-Object { $_.状態 -notlike 'OK*' -and $_.状態 -ne '設定済み(変更なし)' })
        Write-Host ("設定: {0} 件 / 変更なし: {1} 件 / 要確認: {2} 件" -f $ok.Count, $skip.Count, $err.Count) `
            -ForegroundColor $(if ($err.Count -gt 0) { 'Yellow' } else { 'Green' })

        if (-not $Commit) {
            Write-Host ''
            Write-Host '※ プレビューのみ。設定するには「受付番号を設定(書込)」を実行してください。' -ForegroundColor Yellow
            return
        }
        if ($ok.Count -eq 0) { Write-Host '設定する対象がありません。'; return }
        if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
        $bfile = Join-Path $BackupDir ("UKE_NO_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $plan | Select-Object 受付番号, 受診日, 氏名, 現在の番号, 状態, PkSeq |
            Export-Csv -Path $bfile -NoTypeInformation -Encoding UTF8
        Write-Host "[バックアップ] 設定前の状態: $bfile" -ForegroundColor DarkGray

        $tran = $conn.BeginTransaction()
        try {
            $done = 0
            foreach ($t in $ok) {
                $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSIN SET UKE_NO_KENSA = @no WHERE PK_SEQ = @p' `
                    @{ no = $t.受付番号; p = $t.PkSeq }
                if ($n -ne 1) { throw ("UPDATE影響行数が {0} でした (受付番号={1})。ロールバックします。" -f $n, $t.受付番号) }
                $done++
            }
            $tran.Commit()
            Write-Host ("[設定完了] {0} 人の受付番号を設定しました。" -f $done) -ForegroundColor Green
        }
        catch { $tran.Rollback(); throw }
    }
    finally { $conn.Close() }
    return
}

# ---- 通常モード: プレビュー / 書込 ----
$targets = Select-TargetRows $data $idCols
if ($idCols.KenNo -le 0) {
    throw "mapping.csv の KENNO 行に列番号(Col)が設定されていません。まず -Inspect で確認してください。"
}

$conn = Open-Db
try {
    $rosterSet = $null
    if ($Roster) { $rosterSet = Load-RosterPkSeq $conn $Roster }

    $hadError = $false
    $skipped = 0
    # 全体集計 (26人ぶんを1件ずつ目で追わなくて済むように)
    $sumPeople = 0; $sumOk = 0; $sumErr = 0
    $errNos = @(); $notFoundNos = @(); $wroteNos = @(); $nameNgNos = @()
    foreach ($fields in $targets) {
        $kenNo = Normalize-KenNo (Get-Field $fields $idCols.KenNo)
        $ymd = Resolve-RowYmd $fields $idCols
        $pk = Resolve-PkSeq $conn $ymd $kenNo
        if ($null -eq $pk) {
            Write-Warning "受診者が見つかりません (KEN_YMD=$ymd, KEN_NO=$kenNo) → スキップ"
            $hadError = $true
            $notFoundNos += $kenNo
            continue
        }
        if ($rosterSet -and -not $rosterSet.ContainsKey([string]$pk)) {
            $skipped++
            continue   # 名簿外の人 (エラーではない)
        }
        # ---- 氏名の突き合わせ (取り違え防止) ----
        $navi = Get-NaviName $conn $pk
        $naviLabel = if ($navi) { (@($navi.Kanji, $navi.Kana) | Where-Object { $_ -ne '' }) -join ' / ' } else { '(氏名不明)' }
        $nameNg = $false
        $fileKanji = if ($idCols.Kanji -gt 0) { Normalize-Text (Get-Field $fields $idCols.Kanji) } else { '' }
        $fileKana  = if ($idCols.Kana  -gt 0) { Normalize-Text (Get-Field $fields $idCols.Kana)  } else { '' }
        if ($navi -and ($fileKanji -ne '' -or $fileKana -ne '')) {
            # カナを優先して比べる (漢字は旧字体で違うことがあるため)
            if ($fileKana -ne '' -and $navi.Kana -ne '') {
                $nameNg = (Normalize-Name $fileKana) -ne (Normalize-Name $navi.Kana)
            }
            elseif ($fileKanji -ne '' -and $navi.Kanji -ne '') {
                $nameNg = (Normalize-Name $fileKanji) -ne (Normalize-Name $navi.Kanji)
            }
        }

        $current = Get-CurrentKensa $conn $pk
        Write-Host ''
        Write-Host ("--- 受付番号 {0} / 受診日 {1} / {2} / PK_SEQ {3} / 既存T_KENSA枠 {4} 行 ---" -f $kenNo, $ymd, $naviLabel, $pk, $current.Count) -ForegroundColor Cyan
        if ($nameNg) {
            $fileLabel = (@($fileKanji, $fileKana) | Where-Object { $_ -ne '' }) -join ' / '
            Write-Host ("  [氏名が一致しません]  ファイル: {0}   健診ナビ: {1}" -f $fileLabel, $naviLabel) -ForegroundColor Red
            if ($IgnoreName) {
                Write-Host '  → 「氏名の違いを無視」が入っているため、このまま続行します。' -ForegroundColor Yellow
            } else {
                Write-Host '  → 別人に書き込む恐れがあるため、この人はスキップします。受付番号を確認してください。' -ForegroundColor Red
                Write-Host '     (テストデータなど、違っていて当然の場合は「氏名の違いを無視」にチェックを入れてください)' -ForegroundColor DarkGray
                $hadError = $true
                $nameNgNos += $kenNo
                continue
            }
        }
        $plan = Build-Plan $conn $mapRows $valueMap $fields $current
        Show-Plan $plan ("受付番号 " + $kenNo)

        $err = @($plan | Where-Object { $_.Status -ne 'OK' -and $_.Status -ne '列未設定' -and $_.Status -ne '項目CD未設定' })
        $sumPeople++
        $sumOk += @($plan | Where-Object { $_.Status -eq 'OK' }).Count
        $sumErr += $err.Count
        if ($err.Count -gt 0) { $errNos += $kenNo }
        if ($Commit) {
            if ($err.Count -gt 0 -and -not $Force) {
                Write-Host "エラー行があるため書込を中止しました (受付番号 $kenNo)。内容を修正するか、OK行のみ書込む場合は -Force を付けてください。" -ForegroundColor Red
                $hadError = $true
                continue
            }
            Commit-Plan $conn $pk $plan
            $wroteNos += $kenNo
        }
    }
    if ($skipped -gt 0) {
        Write-Host ''
        Write-Host ("[名簿しぼり込み] 名簿に無い {0} 人は取り込みませんでした。" -f $skipped) -ForegroundColor Yellow
    }
    # ---- 全体の集計 ----
    Write-Host ''
    Write-Host ('=' * 60) -ForegroundColor Cyan
    if ($Commit) {
        Write-Host ("【全体】{0} 人を処理 / 書込 {1} 人 / 書込項目 {2} 件 / エラー {3} 件" -f $sumPeople, $wroteNos.Count, $sumOk, $sumErr) -ForegroundColor Cyan
    } else {
        Write-Host ("【全体】{0} 人を確認 / 書込可能 {1} 件 / エラー {2} 件" -f $sumPeople, $sumOk, $sumErr) -ForegroundColor Cyan
    }
    if ($notFoundNos.Count -gt 0) {
        Write-Host ("  受診者が見つからない受付番号 ({0}人): {1}" -f $notFoundNos.Count, ($notFoundNos -join ', ')) -ForegroundColor Red
    }
    if ($nameNgNos.Count -gt 0) {
        Write-Host ("  氏名が一致せずスキップした受付番号 ({0}人): {1}" -f $nameNgNos.Count, ($nameNgNos -join ', ')) -ForegroundColor Red
    }
    if ($errNos.Count -gt 0) {
        Write-Host ("  エラーのある受付番号 ({0}人): {1}" -f $errNos.Count, (($errNos | Select-Object -Unique) -join ', ')) -ForegroundColor Red
        Write-Host '  → 上にスクロールしてその人の「状態」列を確認してください。' -ForegroundColor Red
    }
    if ($notFoundNos.Count -eq 0 -and $errNos.Count -eq 0 -and $nameNgNos.Count -eq 0) {
        Write-Host '  エラーはありません。' -ForegroundColor Green
    }
    Write-Host ('=' * 60) -ForegroundColor Cyan
    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ プレビューのみ実行しました。書込むには「4. 書込実行」を押してください。' -ForegroundColor Yellow
    }
    if ($hadError) { exit 1 }
}
finally { $conn.Close() }
