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
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -Csv C:\...\form.csv -Only 4001 -KenYmd 2026/07/02

  # 書込
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -Csv C:\...\form.csv -Only 4001 -KenYmd 2026/07/02 -Commit

  # CSVの列番号・ヘッダ・値の確認(DB接続なし)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -Csv C:\...\form.csv -Only 4001 -Inspect

  # 対象者の T_KENSA 行(KOMOKU_CDと現在値)を一覧表示 + CSVダンプ
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -Only 4001 -KenYmd 2026/07/02 -DumpItems

  # 所見マスタの一覧 (例: 胸部X線の所見側)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\form_import.ps1 -DumpSyoken ZK021
#>
[CmdletBinding()]
param(
    [string]$Csv,                 # フォームCSVのパス
    [string]$Only,                # 受付番号で絞り込み (例: 4001)
    [switch]$Commit,              # 付けると書込。付けなければプレビューのみ
    [switch]$Force,               # エラー行があってもOK行のみ書込
    [switch]$Inspect,             # CSVの列番号/ヘッダ/値/変換結果を表示(DB接続なし)
    [switch]$NoHeader,            # 1行目からデータの場合に指定(ヘッダ行なし)
    [switch]$DumpItems,           # 対象者のT_KENSA行を一覧表示
    [string]$DumpSyoken,          # 指定SYOKEN_CDのT_SYOKEN2一覧を表示 (SHIN/GANTEI/ZK011/ZK020/ZK021/ZK030/ZK031/ZK041)
    [string]$KenYmd,              # 受診日 'YYYY/MM/DD'。CSVに日付列が無い場合に指定
    [string]$MapDir,              # 対応表フォルダ (既定: スクリプトと同じ場所の form\)
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
    $formats = @('yyyy/MM/dd','yyyy/M/d','yyyy-MM-dd','yyyy-M-d','yyyyMMdd','yyyy年M月d日','yyyy.M.d')
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

function Read-FormCsv([string]$path, [string]$encName) {
    if (-not (Test-Path $path)) { throw "CSVファイルが見つかりません: $path" }
    $enc = if ($encName -eq 'UTF8') { New-Object System.Text.UTF8Encoding($false) }
           else { [System.Text.Encoding]::GetEncoding(932) }
    $text = [System.IO.File]::ReadAllText($path, $enc)
    # UTF-8 BOM除去
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    $rows = Parse-CsvText $text
    # 完全空行を除去
    $out = @()
    foreach ($r in $rows) {
        $isEmpty = $true
        foreach ($f in $r) { if ((Normalize-Text $f) -ne '') { $isEmpty = $false; break } }
        if (-not $isEmpty) { $out += ,$r }
    }
    $minRows = if ($NoHeader) { 1 } else { 2 }
    if ($out.Count -lt $minRows) { throw "CSVにデータ行がありません: $path (ヘッダ行が無いフォームの場合は -NoHeader を指定)" }
    return ,$out
}

# ヘッダ行とデータ行に分離 (-NoHeader 時はダミーヘッダを生成)
function Split-HeaderData($rows) {
    if ($NoHeader) {
        $h = @()
        for ($i = 1; $i -le $rows[0].Count; $i++) { $h += "列$i" }
        return @{ Header = $h; Data = $rows }
    }
    return @{ Header = $rows[0]; Data = @($rows[1..($rows.Count - 1)]) }
}

function Get-Field($fields, [int]$col) {
    if ($col -le 0 -or $col -gt $fields.Count) { return '' }
    return [string]$fields[$col - 1]
}

# ============================================================================
# 対応表の読み込み
# ============================================================================

function Load-Mapping {
    $p = Join-Path $MapDir 'mapping.csv'
    if (-not (Test-Path $p)) { throw "対応表が見つかりません: $p" }
    $rows = Import-Csv -Path $p -Encoding UTF8
    $valid = @('KENNO','KENYMD','VALUE','NYOU','CHORYOKU','SHOKEN','SHOKEN2','SHOKENCD','SHOKENCD2','IGNORE')
    foreach ($r in $rows) {
        $k = (Normalize-Text $r.Kind).ToUpper()
        if ($k -ne '' -and $valid -notcontains $k) {
            throw "mapping.csv の Kind が不正です: '$($r.Kind)' (行: $($r.Label))"
        }
    }
    return $rows
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
    $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KANJA_G WHERE KEN_YMD = @ymd AND KEN_NO = @no' @{ ymd = $ymd; no = $kenNo }
    if ($dt.Rows.Count -eq 0 -and $kenNo -match '^\d+$') {
        # KEN_NO が数値列の場合を考慮して int でも試す
        $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KANJA_G WHERE KEN_YMD = @ymd AND KEN_NO = @no' @{ ymd = $ymd; no = [int]$kenNo }
    }
    if ($dt.Rows.Count -eq 0) { return $null }
    if ($dt.Rows.Count -gt 1) { throw "受診者が複数見つかりました (KEN_YMD=$ymd, KEN_NO=$kenNo)。中止します。" }
    return $dt.Rows[0].PK_SEQ
}

# 対象者の現在の T_KENSA 行 (KOMOKU_CD → 行) を取得
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

# 結果CD(KEKKA_CD)でマスタを引く (SHOKENCD/SHOKENCD2用)
function Find-SyokenCd($conn, [string]$cd, [string]$code) {
    $c = (Normalize-Text $code).ToUpper()
    $dt = Get-SyokenTable $conn $cd
    foreach ($r in $dt.Rows) {
        if ((Normalize-Text ([string]$r.KEKKA_CD)).ToUpper() -eq $c) { return $r }
    }
    return $null
}

# 所見文をマスタと突合 (完全一致 → 空白無視一致)
function Find-Syoken($conn, [string]$cd, [string]$text) {
    $t = Normalize-Text $text
    $dt = Get-SyokenTable $conn $cd
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
    $kenNoCol = 0; $ymdCol = 0
    foreach ($m in $mapRows) {
        $kind = (Normalize-Text $m.Kind).ToUpper()
        $c = 0
        [void][int]::TryParse((Normalize-Text $m.Col), [ref]$c)
        if ($kind -eq 'KENNO')  { $kenNoCol = $c }
        if ($kind -eq 'KENYMD') { $ymdCol = $c }
    }
    return @{ KenNo = $kenNoCol; Ymd = $ymdCol }
}

function Select-TargetRows($dataRows, $idCols) {
    if (-not $Only) { return ,$dataRows }
    if ($idCols.KenNo -le 0) {
        throw "mapping.csv の KENNO 行に列番号(Col)が設定されていません。まず -Inspect で受付番号の列番号を確認し、mapping.csv に記入してください。"
    }
    $want = Normalize-KenNo $Only
    $sel = @()
    foreach ($r in $dataRows) {
        if ((Normalize-KenNo (Get-Field $r $idCols.KenNo)) -eq $want) { $sel += ,$r }
    }
    if ($sel.Count -eq 0) { throw "受付番号 $Only の行がCSVに見つかりません。" }
    return ,$sel
}

function Resolve-RowYmd($fields, $idCols) {
    if ($idCols.Ymd -gt 0) {
        $y = Normalize-Ymd (Get-Field $fields $idCols.Ymd)
        if ($y) { return $y }
    }
    if ($KenYmd) {
        $y = Normalize-Ymd $KenYmd
        if (-not $y) { throw "-KenYmd の日付を解釈できません: $KenYmd" }
        return $y
    }
    throw "受診日が特定できません。mapping.csv の KENYMD 行に列番号を設定するか、-KenYmd 2026/07/02 のように指定してください。"
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

        if ($kind -eq 'SHOKENCD' -or $kind -eq 'SHOKENCD2') {
            # ---- 所見を結果CD(コード)で直接指定する形式 ----
            $shoCode = (Normalize-Text $raw).ToUpper()
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
        else {
            # ---- 値系 (VALUE / NYOU / CHORYOKU) ----
            $conv = Convert-FormValue $valueMap $kind $raw
            if ($conv.Status -eq 'EMPTY') { continue }  # 空欄はスキップ
            if ($komoku -eq '') { $rep.Status = '項目CD未設定'; $rep.New = (Normalize-Text $raw); $plan += $rep; continue }
            if ($conv.Status -eq 'CONVERR') { $rep.Status = "変換不可($kind)"; $rep.New = $conv.Value; $plan += $rep; continue }
            $rep.New = $conv.Value
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
$idCols   = Get-IdColumns $mapRows

# ---- モード: 対象者の T_KENSA 一覧 ----
if ($DumpItems) {
    if (-not $Only) { throw '-DumpItems には -Only <受付番号> が必要です。' }
    $ymd = $null
    if ($Csv) {
        $rows = Read-FormCsv $Csv $CsvEncoding
        $hd = Split-HeaderData $rows
        $sel = Select-TargetRows $hd.Data $idCols
        $ymd = Resolve-RowYmd $sel[0] $idCols
    } else {
        if (-not $KenYmd) { throw '-DumpItems には -KenYmd 2026/07/02 のように受診日も指定してください。' }
        $ymd = Normalize-Ymd $KenYmd
    }
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
$hd = Split-HeaderData $rows
$header = $hd.Header
$data = $hd.Data

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

# ---- 通常モード: プレビュー / 書込 ----
$targets = Select-TargetRows $data $idCols
if ($idCols.KenNo -le 0) {
    throw "mapping.csv の KENNO 行に列番号(Col)が設定されていません。まず -Inspect で確認してください。"
}

$conn = Open-Db
try {
    $hadError = $false
    foreach ($fields in $targets) {
        $kenNo = Normalize-KenNo (Get-Field $fields $idCols.KenNo)
        $ymd = Resolve-RowYmd $fields $idCols
        $pk = Resolve-PkSeq $conn $ymd $kenNo
        if ($null -eq $pk) {
            Write-Warning "受診者が見つかりません (KEN_YMD=$ymd, KEN_NO=$kenNo) → スキップ"
            $hadError = $true
            continue
        }
        $current = Get-CurrentKensa $conn $pk
        Write-Host ''
        Write-Host ("--- 受付番号 {0} / 受診日 {1} / PK_SEQ {2} / 既存T_KENSA枠 {3} 行 ---" -f $kenNo, $ymd, $pk, $current.Count) -ForegroundColor Cyan
        $plan = Build-Plan $conn $mapRows $valueMap $fields $current
        Show-Plan $plan ("受付番号 " + $kenNo)

        $err = @($plan | Where-Object { $_.Status -ne 'OK' -and $_.Status -ne '列未設定' -and $_.Status -ne '項目CD未設定' })
        if ($Commit) {
            if ($err.Count -gt 0 -and -not $Force) {
                Write-Host "エラー行があるため書込を中止しました (受付番号 $kenNo)。内容を修正するか、OK行のみ書込む場合は -Force を付けてください。" -ForegroundColor Red
                $hadError = $true
                continue
            }
            Commit-Plan $conn $pk $plan
        }
    }
    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ プレビューのみ実行しました。書込むには -Commit を付けて再実行してください。' -ForegroundColor Yellow
    }
    if ($hadError) { exit 1 }
}
finally { $conn.Close() }
