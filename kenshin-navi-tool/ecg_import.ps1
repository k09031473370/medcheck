<#
.SYNOPSIS
  健診ナビ11 心電図結果取込 (ecg_import.ps1)

.DESCRIPTION
  心電計の結果CSV (検査日,ID,判定記号,判定,所見コード1,所見1,...,所見コード5,所見5,コメント,担当医師)
  を読み込み、T_KENSA の心電図所見1〜5へ UPDATE で書き込む。
  - 人の特定: T_KANJA_G の KEN_YMD(=検査日列) + KEN_NO(=ID列) → PK_SEQ
  - 所見: 装置コード(B01等) → form\ecg_code_map.csv → T_SYOKEN2(ZK011) の KEKKA_CD
          で突合し、KEKKA=マスタ所見文 / KEKKA_CD / HANTEI_KIGO を書込
  - 所見が無く判定A(異常なし)の人は「NONE」行の定義で「異常なし」を書込
  - 書込先の KOMOKU_CD は form\ecg_items.csv で指定 (所見1〜5)
  - 総合判定は書き込まない (健診ナビの「自動判定」で計算)

.EXAMPLE
  # 装置コードの一覧と変換表の設定状況を確認 (DB接続なし)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\ecg_import.ps1 -Csv <ecg.csv> -ListCodes

  # プレビュー (ID=13 のみ)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\ecg_import.ps1 -Csv <ecg.csv> -Only 13

  # 全員書込
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\ecg_import.ps1 -Csv <ecg.csv> -Commit
#>
[CmdletBinding()]
param(
    [string]$Csv,
    [string]$Only,                # ID(受付番号)で絞り込み
    [switch]$Commit,
    [switch]$Force,
    [switch]$ListCodes,           # CSV内の装置コード一覧と変換表の充足を表示 (DB不要)
    [string]$KenYmd,              # 受診日を強制指定 (通常はCSVの検査日列を使う)
    [string]$MapDir,
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString,
    [ValidateSet('AUTO','SJIS','UTF8')]
    [string]$CsvEncoding = 'AUTO'
)

$ErrorActionPreference = 'Stop'
if (-not $MapDir) { $MapDir = Join-Path $PSScriptRoot 'form' }
$BackupDir = Join-Path $PSScriptRoot 'backup'
$SyokenCd = 'ZK011'   # 心電図の所見マスタ

try { [void][System.Text.Encoding]::GetEncoding(932) }
catch { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) }

# ============================================================================
# 汎用
# ============================================================================

function Normalize-Text([string]$s) {
    if ($null -eq $s) { return '' }
    return $s.Trim()
}

function Normalize-KenNo([string]$s) {
    $v = Normalize-Text $s
    $v = $v -replace '\.0+$', ''
    return $v
}

function Normalize-Ymd([string]$s) {
    $v = Normalize-Text $s
    if ($v -eq '') { return $null }
    if ($v -match '^\d{4,6}(\.0+)?$' -and [double]$v -gt 20000 -and [double]$v -lt 80000) {
        return ([datetime]::FromOADate([double]$v)).ToString('yyyy/MM/dd')
    }
    $dt = [datetime]::MinValue
    foreach ($f in @('yyyy/MM/dd','yyyy/M/d','yyyy-MM-dd','yyyy-M-d','yyyyMMdd','yyyy年M月d日','yyyy.M.d')) {
        if ([datetime]::TryParseExact($v, $f, [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.ToString('yyyy/MM/dd') }
    }
    if ([datetime]::TryParse($v, [ref]$dt)) { return $dt.ToString('yyyy/MM/dd') }
    return $null
}

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
    if ($sb.Length -gt 0 -or $cur.Count -gt 0) { [void]$cur.Add($sb.ToString()); [void]$rows.Add($cur.ToArray()) }
    return ,$rows.ToArray()
}

# エンコーディング自動判定つき読み込み (ヘッダに「検査日」が含まれる方を採用)
function Read-EcgCsv([string]$path) {
    if (-not (Test-Path $path)) { throw "CSVファイルが見つかりません: $path" }
    $encList = switch ($CsvEncoding) {
        'SJIS' { @([System.Text.Encoding]::GetEncoding(932)) }
        'UTF8' { @((New-Object System.Text.UTF8Encoding($false))) }
        default { @([System.Text.Encoding]::GetEncoding(932), (New-Object System.Text.UTF8Encoding($false))) }
    }
    foreach ($enc in $encList) {
        $text = [System.IO.File]::ReadAllText($path, $enc)
        if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
        $rows = Parse-CsvText $text
        $out = @()
        foreach ($r in $rows) {
            $isEmpty = $true
            foreach ($f in $r) { if ((Normalize-Text $f) -ne '') { $isEmpty = $false; break } }
            if (-not $isEmpty) { $out += ,$r }
        }
        if ($out.Count -ge 1 -and (($out[0] -join ',') -match '検査日')) { return ,$out }
    }
    throw "CSVのヘッダに「検査日」が見つかりません(文字コードを -CsvEncoding SJIS / UTF8 で指定して再試行してください): $path"
}

# ============================================================================
# 対応表
# ============================================================================

function Load-CodeMap {
    $p = Join-Path $MapDir 'ecg_code_map.csv'
    if (-not (Test-Path $p)) { throw "変換表が見つかりません: $p" }
    $map = @{}
    foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
        $cd = (Normalize-Text $r.DeviceCd).ToUpper()
        if ($cd -eq '') { continue }
        $map[$cd] = @{ KekkaCd = (Normalize-Text $r.KEKKA_CD); Text = (Normalize-Text $r.DeviceText) }
    }
    return $map
}

function Load-Slots {
    $p = Join-Path $MapDir 'ecg_items.csv'
    if (-not (Test-Path $p)) { throw "所見枠の設定が見つかりません: $p" }
    $slots = @()
    $hantei = ''
    foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
        $slot = (Normalize-Text $r.Slot).ToUpper()
        if ($slot -match '^\d+$') {
            $slots += @{ No = [int]$slot; KomokuCd = (Normalize-Text $r.KOMOKU_CD) }
        }
        elseif ($slot -eq 'HANTEI') { $hantei = Normalize-Text $r.KOMOKU_CD }
    }
    return @{ Slots = @($slots | Sort-Object { $_.No }); Hantei = $hantei }
}

# ============================================================================
# DB (form_import.ps1 と同じ方式)
# ============================================================================

function Resolve-ConnectionString {
    if ($ConnectionString) { return $ConnectionString }
    if (Test-Path $ConnFile) {
        $txt = Get-Content -Path $ConnFile -Raw
        $lines = @($txt -split "`r?`n" | Where-Object { (Normalize-Text $_) -ne '' })
        foreach ($ln in $lines) {
            if ($ln -match '(?i)(data source|server)\s*=') {
                $cs = Normalize-Text $ln
                if ($cs -notmatch '(?i)(initial catalog|database)\s*=') {
                    $cs = $cs.TrimEnd(';') + ';Initial Catalog=K166_SIBAURAUSER'
                }
                return $cs
            }
        }
        $kv = @{}
        foreach ($ln in $lines) {
            if ($ln -match '^\s*([^=:]+?)\s*[=:]\s*(.*?)\s*$') { $kv[$Matches[1].ToUpper()] = $Matches[2] }
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
    if ($params) { foreach ($k in $params.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $params[$k]) } }
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
    if ($params) { foreach ($k in $params.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $params[$k]) } }
    $n = $cmd.ExecuteNonQuery()
    $cmd.Dispose()
    return $n
}

function Resolve-PkSeq($conn, [string]$ymd, [string]$kenNo) {
    $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KANJA_G WHERE KEN_YMD = @ymd AND KEN_NO = @no' @{ ymd = $ymd; no = $kenNo }
    if ($dt.Rows.Count -eq 0 -and $kenNo -match '^\d+$') {
        $dt = Invoke-DbQuery $conn 'SELECT PK_SEQ FROM T_KANJA_G WHERE KEN_YMD = @ymd AND KEN_NO = @no' @{ ymd = $ymd; no = [int]$kenNo }
    }
    if ($dt.Rows.Count -eq 0) { return $null }
    if ($dt.Rows.Count -gt 1) { throw "受診者が複数見つかりました (KEN_YMD=$ymd, KEN_NO=$kenNo)。中止します。" }
    return $dt.Rows[0].PK_SEQ
}

$script:Zk011 = $null
function Get-Zk011($conn) {
    if ($null -eq $script:Zk011) {
        $script:Zk011 = Invoke-DbQuery $conn 'SELECT KEKKA_CD, SYOKEN, HANTEI_KIGO FROM T_SYOKEN2 WHERE SYOKEN_CD = @cd' @{ cd = $SyokenCd }
    }
    return $script:Zk011
}

function Find-ByKekkaCd($conn, [string]$kc) {
    foreach ($r in (Get-Zk011 $conn).Rows) {
        if ((Normalize-Text ([string]$r.KEKKA_CD)) -eq $kc) { return $r }
    }
    return $null
}

function Find-ByText($conn, [string]$text) {
    $t = Normalize-Text $text
    $dt = Get-Zk011 $conn
    foreach ($r in $dt.Rows) { if ((Normalize-Text ([string]$r.SYOKEN)) -eq $t) { return $r } }
    $t2 = $t -replace '[\s　]', ''
    foreach ($r in $dt.Rows) { if ((([string]$r.SYOKEN) -replace '[\s　]', '') -eq $t2) { return $r } }
    return $null
}

# ============================================================================
# CSV解析
# ============================================================================

if (-not $Csv) { throw '使い方: ecg_import.ps1 -Csv <ecg.csv> [-Only <ID>] [-Commit] / -ListCodes' }

$rows = Read-EcgCsv $Csv
$header = $rows[0]
$data = if ($rows.Count -gt 1) { $rows[1..($rows.Count - 1)] } else { @() }

# ヘッダ名 → 列index
$colIdx = @{}
for ($i = 0; $i -lt $header.Count; $i++) { $colIdx[(Normalize-Text $header[$i])] = $i }
foreach ($need in @('検査日','ID','判定記号')) {
    if (-not $colIdx.ContainsKey($need)) { throw "CSVに「$need」列がありません。ヘッダ: $($header -join ', ')" }
}

function Get-Col($fields, [string]$name) {
    if (-not $colIdx.ContainsKey($name)) { return '' }
    $i = $colIdx[$name]
    if ($i -ge $fields.Count) { return '' }
    return Normalize-Text ([string]$fields[$i])
}

# 1行 → person オブジェクト
function Parse-Row($fields) {
    $findings = @()
    for ($i = 1; $i -le 5; $i++) {
        $cd = (Get-Col $fields "所見コード$i").ToUpper()
        $tx = Get-Col $fields "所見$i"
        if ($cd -ne '' -or $tx -ne '') { $findings += @{ Cd = $cd; Text = $tx } }
    }
    return @{
        Ymd    = Normalize-Ymd (Get-Col $fields '検査日')
        Id     = Normalize-KenNo (Get-Col $fields 'ID')
        Hantei = Get-Col $fields '判定記号'
        HanteiName = Get-Col $fields '判定'
        Findings = $findings
    }
}

$persons = @()
foreach ($f in $data) {
    $p = Parse-Row $f
    if ($p.Id -eq '') { continue }
    if ($Only -and $p.Id -ne (Normalize-KenNo $Only)) { continue }
    $persons += $p
}
if ($persons.Count -eq 0) { throw "対象データがありません$(if ($Only) { " (ID=$Only)" })。" }

$codeMap = Load-CodeMap

# ---- モード: コード一覧 (DB不要) ----
if ($ListCodes) {
    Write-Host "=== CSV内の装置コードと変換表の状況 (対象 $($persons.Count) 人) ===" -ForegroundColor Cyan
    $seen = @{}
    foreach ($p in $persons) {
        foreach ($fd in $p.Findings) {
            $key = $fd.Cd
            if (-not $seen.ContainsKey($key)) { $seen[$key] = @{ Text = $fd.Text; Count = 0 } }
            $seen[$key].Count++
        }
        if ($p.Findings.Count -eq 0) {
            if (-not $seen.ContainsKey('NONE')) { $seen['NONE'] = @{ Text = '(所見なし→異常なし)'; Count = 0 } }
            $seen['NONE'].Count++
        }
    }
    $out = foreach ($k in ($seen.Keys | Sort-Object)) {
        $kc = ''
        if ($codeMap.ContainsKey($k)) { $kc = $codeMap[$k].KekkaCd }
        New-Object PSObject -Property @{
            'コード' = $k; '所見文(装置)' = $seen[$k].Text; '件数' = $seen[$k].Count
            'KEKKA_CD' = $kc; '状態' = $(if ($kc -ne '') { 'OK' } else { '未設定' })
        }
    }
    $out | Select-Object コード, '所見文(装置)', 件数, KEKKA_CD, 状態 | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    Write-Host '「未設定」の行は form\ecg_code_map.csv の KEKKA_CD を記入してください (-DumpSyoken ZK011 の一覧と突合)。'
    return
}

# ---- 通常モード: プレビュー / 書込 ----
$slotCfg = Load-Slots
$validSlots = @($slotCfg.Slots | Where-Object { $_.KomokuCd -ne '' })
$hanteiKomoku = $slotCfg.Hantei
if ($validSlots.Count -eq 0) {
    throw 'form\ecg_items.csv の KOMOKU_CD が未設定です。-DumpItems (form_import.ps1) で心電図所見1〜のKOMOKU_CDを確認して記入してください。'
}
if ($hanteiKomoku -eq '') {
    Write-Warning 'ecg_items.csv の HANTEI 行が未設定のため、判定は書き込みません(自動判定に任せます)。'
}

$conn = Open-Db
try {
    $hadError = $false
    foreach ($p in $persons) {
        $ymd = $p.Ymd
        if ($KenYmd) { $ymd = Normalize-Ymd $KenYmd }
        if (-not $ymd) { Write-Warning "ID=$($p.Id): 検査日が解釈できません → スキップ"; $hadError = $true; continue }

        $pk = Resolve-PkSeq $conn $ymd $p.Id
        if ($null -eq $pk) {
            Write-Warning "受診者が見つかりません (KEN_YMD=$ymd, KEN_NO=$($p.Id)) → スキップ"
            $hadError = $true
            continue
        }
        $curDt = Invoke-DbQuery $conn 'SELECT KOMOKU_CD, KEKKA, KEKKA_CD, HANTEI_KIGO FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pk }
        $current = @{}
        foreach ($r in $curDt.Rows) { $current[(Normalize-Text ([string]$r.KOMOKU_CD))] = $r }

        # 書き込む所見リストを確定
        $toWrite = @()
        $rowErr = @()
        $fds = $p.Findings
        if ($fds.Count -eq 0) {
            if ($codeMap.ContainsKey('NONE')) { $fds = @(@{ Cd = 'NONE'; Text = '異常なし' }) }
            else { $rowErr += 'NONE(異常なし)の変換が未設定' }
        }
        if ($fds.Count -gt $validSlots.Count) {
            $rowErr += "所見が $($fds.Count) 件ありますが枠は $($validSlots.Count) 個です"
        }
        for ($i = 0; $i -lt [Math]::Min($fds.Count, $validSlots.Count); $i++) {
            $fd = $fds[$i]
            $slot = $validSlots[$i]
            $hit = $null
            $via = ''
            if ($fd.Cd -ne '' -and $codeMap.ContainsKey($fd.Cd) -and $codeMap[$fd.Cd].KekkaCd -ne '') {
                $hit = Find-ByKekkaCd $conn $codeMap[$fd.Cd].KekkaCd
                $via = "コード$($fd.Cd)→$($codeMap[$fd.Cd].KekkaCd)"
                if (-not $hit) { $rowErr += "$($fd.Cd): KEKKA_CD=$($codeMap[$fd.Cd].KekkaCd) がZK011に存在しません" }
            }
            elseif ($fd.Text -ne '') {
                $hit = Find-ByText $conn $fd.Text
                $via = '所見文一致'
                if (-not $hit) { $rowErr += "$($fd.Cd) 「$($fd.Text)」: 変換表未設定かつマスタに同文なし" }
            }
            else { $rowErr += "所見$($i+1): コードも所見文も空" }

            if ($hit) {
                $komoku = $slot.KomokuCd
                if (-not $current.ContainsKey($komoku)) { $rowErr += "枠なし: KOMOKU_CD=$komoku (所見$($slot.No))" }
                else {
                    $toWrite += @{
                        SlotName = "所見$($slot.No)"; Komoku = $komoku; Via = $via; IsHantei = $false
                        Now = Normalize-Text ([string]$current[$komoku].KEKKA)
                        Kekka = [string]$hit.SYOKEN
                        KekkaCd = Normalize-Text ([string]$hit.KEKKA_CD)
                        Hantei = Normalize-Text ([string]$hit.HANTEI_KIGO)
                    }
                }
            }
        }

        # 判定 (装置の判定記号をそのまま書込。ecg_items.csv の HANTEI 行設定時のみ)
        if ($hanteiKomoku -ne '' -and $p.Hantei -ne '') {
            if (-not $current.ContainsKey($hanteiKomoku)) { $rowErr += "枠なし: KOMOKU_CD=$hanteiKomoku (判定)" }
            else {
                $toWrite += @{
                    SlotName = '判定'; Komoku = $hanteiKomoku; Via = '装置判定記号(HANTEI_KIGOのみ更新)'; IsHantei = $true
                    Now = Normalize-Text ([string]$current[$hanteiKomoku].HANTEI_KIGO)
                    Kekka = ''
                    KekkaCd = ''
                    Hantei = $p.Hantei
                }
            }
        }

        Write-Host ''
        Write-Host ("--- ID {0} / 受診日 {1} / PK_SEQ {2} / 装置判定 {3} {4} ---" -f $p.Id, $ymd, $pk, $p.Hantei, $p.HanteiName) -ForegroundColor Cyan
        if ($toWrite.Count -gt 0) {
            $toWrite | ForEach-Object {
                New-Object PSObject -Property @{
                    '枠' = $_.SlotName; 'KOMOKU_CD' = $_.Komoku; '現在値' = $_.Now
                    '書込所見' = $_.Kekka; 'KEKKA_CD' = $_.KekkaCd; '判定' = $_.Hantei; '根拠' = $_.Via
                }
            } | Select-Object 枠, KOMOKU_CD, 現在値, 書込所見, KEKKA_CD, 判定, 根拠 |
                Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host
        }
        foreach ($e in $rowErr) { Write-Host "  [エラー] $e" -ForegroundColor Red }

        if ($Commit) {
            if ($rowErr.Count -gt 0 -and -not $Force) {
                Write-Host "エラーがあるため ID=$($p.Id) の書込を中止 (-Force でOK分のみ書込可)" -ForegroundColor Red
                $hadError = $true
                continue
            }
            if ($toWrite.Count -eq 0) { Write-Host '書込対象なし'; continue }
            if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
            $full = Invoke-DbQuery $conn 'SELECT * FROM T_KENSA WHERE PK_SEQ = @p' @{ p = $pk }
            $bfile = Join-Path $BackupDir ("T_KENSA_{0}_{1}.csv" -f $pk, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $full | Export-Csv -Path $bfile -NoTypeInformation -Encoding UTF8
            Write-Host "[バックアップ] $bfile" -ForegroundColor DarkGray
            $tran = $conn.BeginTransaction()
            try {
                $done = 0
                foreach ($w in $toWrite) {
                    if ($w.IsHantei) {
                        # 判定は HANTEI_KIGO のみ更新 (同じ行の所見文 KEKKA を上書きしないため)
                        $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET HANTEI_KIGO = @h WHERE PK_SEQ = @p AND KOMOKU_CD = @cd' @{
                            h = $w.Hantei; p = $pk; cd = $w.Komoku
                        }
                    }
                    else {
                        $n = Invoke-DbExec $conn $tran 'UPDATE T_KENSA SET KEKKA = @k, KEKKA_CD = @kc, HANTEI_KIGO = @h WHERE PK_SEQ = @p AND KOMOKU_CD = @cd' @{
                            k = $w.Kekka; kc = $w.KekkaCd; h = $w.Hantei; p = $pk; cd = $w.Komoku
                        }
                    }
                    if ($n -ne 1) { throw ("UPDATE影響行数が {0} (KOMOKU_CD={1})。ロールバックします。" -f $n, $w.Komoku) }
                    $done++
                }
                $tran.Commit()
                Write-Host ("[書込完了] ID={0}: {1} 件の所見を書込" -f $p.Id, $done) -ForegroundColor Green
            }
            catch { $tran.Rollback(); throw }
        }
        if ($rowErr.Count -gt 0) { $hadError = $true }
    }
    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ プレビューのみ。書込むには -Commit を付けて再実行してください。' -ForegroundColor Yellow
    }
    Write-Host '※ 書込後は健診ナビの「自動判定」を実行してください。'
    if ($hadError) { exit 1 }
}
finally { $conn.Close() }
