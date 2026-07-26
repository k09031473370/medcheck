<#
.SYNOPSIS
  健診ナビDB 調査・掃除ツール (db_tool.ps1)

.DESCRIPTION
  アプリの取込テストで作ってしまった不要データ(検査センター「院内」など)を
  探して安全に削除するためのツール。
  - 検索: DB内の全テーブルの文字列列から指定文字列を含む行を探す
  - 確認: テーブルの中身を表示 (WHERE指定可)
  - 削除: 対象行をCSVにバックアップしてから削除 (-Commit を付けるまで実削除しない)

.EXAMPLE
  # 1) 「院内」がどのテーブルに入っているか探す
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -FindText 院内

  # 2) 見つかったテーブルの中身を確認 (例: テーブル名 M_CENTER だった場合)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -Table M_CENTER

  # 3) 削除プレビュー (まだ消えない。対象行が表示される)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -Table M_CENTER -Where "CENTER_NM = N'院内'" -Delete

  # 4) 問題なければ実削除 (削除前に backup\ へCSV保存される)
  powershell -ExecutionPolicy Bypass -File C:\kenshin-navi\db_tool.ps1 -Table M_CENTER -Where "CENTER_NM = N'院内'" -Delete -Commit
#>
[CmdletBinding()]
param(
    [string]$FindText,            # DB全体からこの文字列を含む行を探す
    [string]$Table,               # 対象テーブル名 (表示/削除)
    [string]$Where,               # 絞り込み条件 (SQLのWHERE句。例: "CENTER_NM = N'院内'")
    [string]$Sql,                 # 任意のSELECT文を実行 (SELECT以外は拒否)
    [string]$Database,            # 接続先DBを一時的に変更 (例: -Database master)
    [switch]$Delete,              # 削除モード (Table と Where が必須)
    [switch]$Commit,              # 付けると実削除。付けなければプレビューのみ
    [int]$MaxRows = 30,           # 表示する最大行数
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)

$ErrorActionPreference = 'Stop'
$BackupDir = Join-Path $PSScriptRoot 'backup'

function Normalize-Text([string]$s) {
    if ($null -eq $s) { return '' }
    return $s.Trim()
}

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
        Write-Warning "接続ファイルを解釈できませんでした: $ConnFile → 既定値で続行します。"
    }
    return 'Data Source=KNSV\SQLEXPRESS;Initial Catalog=K166_SIBAURAUSER;Integrated Security=True'
}

function Open-Db {
    $cs = Resolve-ConnectionString
    if ($Database) {
        if ($Database -notmatch '^[A-Za-z0-9_]+$') { throw "データベース名が不正です: $Database" }
        if ($cs -match '(?i)(initial catalog|database)\s*=') {
            $cs = $cs -replace '(?i)(initial catalog|database)\s*=\s*[^;]+', ('Initial Catalog=' + $Database)
        } else {
            $cs = $cs.TrimEnd(';') + ';Initial Catalog=' + $Database
        }
    }
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

# テーブル名の安全化 ([ ] で囲む。壊れた名前は拒否)
function Safe-Table([string]$name) {
    $n = Normalize-Text $name
    if ($n -notmatch '^[A-Za-z0-9_]+$') { throw "テーブル名が不正です: $name" }
    return "[$n]"
}

function Show-Table($dt, [int]$max) {
    $total = $dt.Rows.Count
    $view = $dt | Select-Object -First $max
    $view | Format-Table -AutoSize -Wrap | Out-String -Width 400 | Write-Host
    if ($total -gt $max) { Write-Host "... 他 $($total - $max) 行 (全 $total 行)" -ForegroundColor DarkGray }
    else { Write-Host "(全 $total 行)" -ForegroundColor DarkGray }
}

function Backup-Rows($dt, [string]$tag) {
    if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
    $file = Join-Path $BackupDir ("{0}_{1}.csv" -f $tag, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $dt | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Host "[バックアップ] $file" -ForegroundColor DarkGray
    return $file
}

# ============================================================================
# メイン
# ============================================================================

if (-not $FindText -and -not $Table -and -not $Sql) {
    Write-Host '使い方:'
    Write-Host '  検索: db_tool.ps1 -FindText 院内 [-Database <DB名>]'
    Write-Host '  表示: db_tool.ps1 -Table <テーブル名> [-Where "<条件>"]'
    Write-Host '  SQL : db_tool.ps1 -Sql "SELECT ..." (SELECTのみ)'
    Write-Host '  削除: db_tool.ps1 -Table <テーブル名> -Where "<条件>" -Delete [-Commit]'
    return
}

$conn = Open-Db
try {
    # ---- モード0: 任意SELECT ----
    if ($Sql) {
        $s = $Sql.Trim()
        if ($s -notmatch '^(?i)\s*SELECT\b') { throw '-Sql で実行できるのは SELECT 文のみです。' }
        $dt = Invoke-DbQuery $conn $s $null
        Show-Table $dt $MaxRows
        return
    }

    # ---- モード1: 全テーブル文字列検索 ----
    if ($FindText) {
        $needle = $FindText -replace "'", "''"
        Write-Host "=== DB全体から「$FindText」を検索します (文字列列のみ) ===" -ForegroundColor Cyan
        $cols = Invoke-DbQuery $conn @"
SELECT t.name AS TBL, c.name AS COL
FROM sys.tables t
JOIN sys.columns c ON c.object_id = t.object_id
JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE ty.name IN ('char','varchar','nchar','nvarchar','text','ntext')
ORDER BY t.name, c.column_id
"@ $null
        $hitTables = @{}
        foreach ($r in $cols.Rows) {
            $tbl = [string]$r.TBL; $col = [string]$r.COL
            try {
                $cnt = Invoke-DbQuery $conn "SELECT COUNT(*) AS N FROM [$tbl] WHERE [$col] LIKE N'%$needle%'" $null
                $n = [int]$cnt.Rows[0].N
                if ($n -gt 0) {
                    Write-Host ("  {0}.{1} : {2} 行" -f $tbl, $col, $n) -ForegroundColor Yellow
                    if (-not $hitTables.ContainsKey($tbl)) { $hitTables[$tbl] = @() }
                    $hitTables[$tbl] += $col
                }
            } catch { }  # text型など比較できない列はスキップ
        }
        if ($hitTables.Count -eq 0) { Write-Host "見つかりませんでした。"; return }
        foreach ($tbl in $hitTables.Keys) {
            $conds = @($hitTables[$tbl] | ForEach-Object { "[$_] LIKE N'%$needle%'" }) -join ' OR '
            Write-Host ''
            Write-Host "--- $tbl (該当行の内容) ---" -ForegroundColor Cyan
            $dt = Invoke-DbQuery $conn "SELECT TOP $MaxRows * FROM [$tbl] WHERE $conds" $null
            Show-Table $dt $MaxRows
        }
        Write-Host ''
        Write-Host '次の手順: 削除したい行が見つかったら' -ForegroundColor Green
        Write-Host '  db_tool.ps1 -Table <テーブル名> -Where "<列> = N''院内''" -Delete       ← プレビュー'
        Write-Host '  db_tool.ps1 -Table <テーブル名> -Where "<列> = N''院内''" -Delete -Commit ← 実削除'
        return
    }

    # ---- モード2/3: テーブル表示・削除 ----
    $tbl = Safe-Table $Table
    $whereSql = ''
    if ($Where) { $whereSql = ' WHERE ' + $Where }

    $dt = Invoke-DbQuery $conn ("SELECT * FROM $tbl" + $whereSql) $null

    if (-not $Delete) {
        Write-Host "=== $tbl$whereSql ===" -ForegroundColor Cyan
        Show-Table $dt $MaxRows
        return
    }

    # 削除モード
    if (-not $Where) { throw '削除には -Where が必須です (全行削除の事故防止)。' }
    Write-Host "=== 削除対象: $tbl WHERE $Where ===" -ForegroundColor Yellow
    Show-Table $dt $MaxRows
    if ($dt.Rows.Count -eq 0) { Write-Host '対象行がありません。'; return }

    if (-not $Commit) {
        Write-Host ''
        Write-Host "※ プレビューのみ。上記 $($dt.Rows.Count) 行を削除するには -Commit を付けて再実行してください。" -ForegroundColor Yellow
        return
    }

    [void](Backup-Rows $dt ("DELETE_" + (Normalize-Text $Table)))
    $tran = $conn.BeginTransaction()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.Transaction = $tran
        $cmd.CommandText = "DELETE FROM $tbl WHERE $Where"
        $n = $cmd.ExecuteNonQuery()
        $cmd.Dispose()
        if ($n -ne $dt.Rows.Count) {
            throw "削除行数($n)がプレビュー($($dt.Rows.Count))と一致しません。ロールバックしました。"
        }
        $tran.Commit()
        Write-Host "[削除完了] $n 行を削除しました。" -ForegroundColor Green
    }
    catch {
        $tran.Rollback()
        throw
    }
}
finally { $conn.Close() }
