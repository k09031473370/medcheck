<#
.SYNOPSIS
  練習用DBを作るためのSQLファイルを書き出す (export_testdb.ps1)

.DESCRIPTION
  本番DBを「読むだけ」で、テスト用DBを作り直すためのSQLファイルを1つ作る。
  バックアップ(.bak)が取り出せない環境のための代替手段。

  ★本番には一切書き込まない。SELECT しかしない。
  ★書き出す時点で氏名などを架空の値に置き換えるので、
    出来上がるファイルに患者の個人情報は入らない。

  出来たSQLファイルを自分のPCへ持って行き、空のDBに対して流せば
  取込・請求・帳票の試験ができる環境になる。

.EXAMPLE
  # 直近6か月分を書き出す (既定)
  powershell -ExecutionPolicy Bypass -File export_testdb.ps1

  # 期間と出力先を指定
  powershell -ExecutionPolicy Bypass -File export_testdb.ps1 -Months 12 -Out D:\testdb.sql

  # 仮名化しない (院内のテスト機用。持ち出し厳禁)
  powershell -ExecutionPolicy Bypass -File export_testdb.ps1 -NoMask
#>
[CmdletBinding()]
param(
    [int]$Months = 6,             # 何か月前までの受診データを持っていくか
    [string]$Out,                 # 出力先。省略時はデスクトップ
    [switch]$NoMask,              # 付けると仮名化しない (院内専用)
    [switch]$Ask,                 # バッチから使う: 期間を対話で聞く
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)

$ErrorActionPreference = 'Stop'

if ($Ask) {
    Write-Host '================================================' -ForegroundColor Cyan
    Write-Host ' 練習用DBを作るためのSQLファイルを書き出します' -ForegroundColor Cyan
    Write-Host ' (健診ナビは読むだけです。データは変わりません)' -ForegroundColor Cyan
    Write-Host '================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host ' 氏名・住所などは架空の値に置き換えて書き出すので、'
    Write-Host ' 出来上がるファイルに患者さんの個人情報は入りません。'
    Write-Host ''
    $a = Read-Host '何か月分の受診データを入れますか? (空Enter=6)'
    if ($a -match '^\d+$') { $Months = [int]$a }
    Write-Host ''
    Write-Host '書き出しています... (数分かかることがあります)'
    Write-Host ''
}

$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Normalize-KenNo')
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Get-CurrentKensa')

# ---- 持っていくテーブル ----
# Filter: 空=全件 / それ以外は WHERE 句。@CUT は期間の開始日に置き換わる。
# 受診データは「その期間の T_KENSIN にぶら下がるもの」だけを持っていく。
$SUB_PK    = "SELECT PK_SEQ FROM T_KENSIN WHERE D_KENSIN >= '@CUT'"
$SUB_KOJIN = "SELECT KOJIN_ID FROM T_KENSIN WHERE D_KENSIN >= '@CUT'"
$TABLES = @(
    # --- マスタ (全件。ここが無いと判定も料金も動かない) ---
    @{ 名 = 'T_KOMOKU';      区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_SYOKEN1';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_SYOKEN2';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_KIJUN1';      区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_KIJUN2';      区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_COURSE1';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_COURSE2';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_COURSE3';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_COURSE4';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'M_DANTAI';      区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'M_OPTION';      区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'M_KENPO';       区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_DANTAI1';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_DANTAI2';     区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'T_BUSYO';       区分 = 'マスタ'; Filter = '' }
    @{ 名 = 'M_HENKAN_IRAI'; 区分 = 'マスタ'; Filter = '' }
    # --- 受診データ (期間で絞る) ---
    @{ 名 = 'T_KENSIN';   区分 = '受診'; Filter = "D_KENSIN >= '@CUT'" }
    @{ 名 = 'T_KANJA_G';  区分 = '受診'; Filter = "PK_SEQ IN ($SUB_PK)" }
    @{ 名 = 'T_KENSA';    区分 = '受診'; Filter = "PK_SEQ IN ($SUB_PK)" }
    @{ 名 = 'T_RYOUKIN';  区分 = '受診'; Filter = "PK_SEQ IN ($SUB_PK)" }
    @{ 名 = 'T_KOJIN1';   区分 = '受診'; Filter = "KOJIN_ID IN ($SUB_KOJIN)" }
    @{ 名 = 'T_KOJIN2';   区分 = '受診'; Filter = "KOJIN_ID IN ($SUB_KOJIN)" }
)

# ---- 仮名化する列 (テーブル名.列名 → 置き換えるSQL式) ----
# 読み出す SELECT の中で置き換えるので、生の個人情報はこのPCから出ない。
$MASK = @{
    'T_KOJIN1.KANJI_SIMEI'     = "N'検査' + RIGHT('0000'+CAST(KOJIN_ID AS varchar(10)),4) + N'太郎'"
    'T_KOJIN1.KANA_SIMEI'      = "N'ｹﾝｻ' + RIGHT('0000'+CAST(KOJIN_ID AS varchar(10)),4) + N'ﾀﾛｳ'"
    'T_KOJIN1.KANJI_SIMEI_OLD' = 'NULL'
    'T_KOJIN1.KANA_SIMEI_OLD'  = 'NULL'
    'T_KOJIN1.EIGO_SIMEI'      = 'NULL'
    'T_KOJIN1.KOJIN_NO'        = "CASE WHEN LTRIM(RTRIM(ISNULL(KOJIN_NO,'')))='' THEN KOJIN_NO ELSE 'E'+CAST(KOJIN_ID AS varchar(10)) END"
    'T_KOJIN1.KARUTE_NO'       = "CASE WHEN LTRIM(RTRIM(ISNULL(KARUTE_NO,'')))='' THEN KARUTE_NO ELSE 'K'+CAST(KOJIN_ID AS varchar(10)) END"
    'T_KOJIN1.TECHO_NO'        = 'NULL'
    'T_KOJIN1.HIBAKU_NO'       = 'NULL'
    'T_KOJIN1.D_SEINEN'        = "CASE WHEN LEN(LTRIM(RTRIM(ISNULL(D_SEINEN,''))))>=8 THEN LEFT(D_SEINEN,4)+'/01/01' ELSE D_SEINEN END"
    'T_KOJIN2.YUUBIN'          = 'NULL'
    'T_KOJIN2.JYUSYO1'         = "N'（練習用）'"
    'T_KOJIN2.JYUSYO2'         = 'NULL'
    'T_KOJIN2.JYUSYO3'         = 'NULL'
    'T_KOJIN2.TEL'             = 'NULL'
    'T_KOJIN2.FAX'             = 'NULL'
    'T_KOJIN2.TEL_REN'         = 'NULL'
    'T_KOJIN2.TEL2'            = 'NULL'
    'T_KOJIN2.E_MAIL'          = 'NULL'
    'T_KOJIN2.CYUUI_MEMO'      = 'NULL'
    'T_KOJIN2.BIKOU'           = 'NULL'
    'T_KOJIN2.KENPO_KIGO'      = 'NULL'
    'T_KOJIN2.KENPO_BANGO'     = 'NULL'
    'T_KOJIN2.KENPO_BANGO_EDA' = 'NULL'
    'T_KOJIN2.SYUSSEKI_NO'     = 'NULL'
    'T_KOJIN2.SOUFU_JYUSYO1'   = 'NULL'
    'T_KOJIN2.SOUFU_JYUSYO2'   = 'NULL'
    'T_KOJIN2.SOUFU_JYUSYO3'   = 'NULL'
    'T_KOJIN2.SOUFU_YUUBIN'    = 'NULL'
    'T_KENSIN.BIKO'            = 'NULL'
    'T_KENSIN.TYUUI'           = 'NULL'
    'T_KENSIN.IKANSEN_MEMO'    = 'NULL'
}

$INV = [System.Globalization.CultureInfo]::InvariantCulture

# ---- SQL用のリテラルに変換 ----
function ConvertTo-SqlLiteral($v, [string]$type) {
    if ($null -eq $v -or [DBNull]::Value.Equals($v)) { return 'NULL' }
    switch -Regex ($type) {
        '^bit$' { if ([bool]$v) { return '1' } else { return '0' } }
        '^(int|bigint|smallint|tinyint)$' { return ([string]$v) }
        '^(numeric|decimal|money|smallmoney)$' { return ([decimal]$v).ToString($INV) }
        '^(float|real)$' { return ([double]$v).ToString('R', $INV) }
        '^(date)$' { return "'" + ([datetime]$v).ToString('yyyy-MM-dd') + "'" }
        '^(datetime|datetime2|smalldatetime|datetimeoffset)$' {
            return "'" + ([datetime]$v).ToString('yyyy-MM-ddTHH:mm:ss.fff') + "'"
        }
        '^time$' { return "'" + ([timespan]$v).ToString('hh\:mm\:ss') + "'" }
        '^uniqueidentifier$' { return "'" + ([guid]$v).ToString() + "'" }
        '^(binary|varbinary|image|timestamp|rowversion)$' { return 'NULL' }   # 画像類は持っていかない
        default {
            return "N'" + ([string]$v).Replace("'", "''") + "'"
        }
    }
}

function Get-ColumnTypeSql($col) {
    $t = [string]$col.DATA_TYPE
    $len = $col.CHARACTER_MAXIMUM_LENGTH
    $p = $col.NUMERIC_PRECISION; $s = $col.NUMERIC_SCALE
    switch -Regex ($t) {
        # rowversion は入れ直せないので、ただの varbinary にしておく
        '^(timestamp|rowversion)$' { return 'varbinary(8)' }
        '^(char|varchar|nchar|nvarchar|binary|varbinary)$' {
            $l = if ($null -eq $len -or [DBNull]::Value.Equals($len) -or [int]$len -lt 0) { 'MAX' } else { [string][int]$len }
            return "$t($l)"
        }
        '^(numeric|decimal)$' { return "$t($([int]$p),$([int]$s))" }
        '^(datetime2|time|datetimeoffset)$' {
            if ($null -ne $col.DATETIME_PRECISION -and -not [DBNull]::Value.Equals($col.DATETIME_PRECISION)) {
                return "$t($([int]$col.DATETIME_PRECISION))"
            }
            return $t
        }
        default { return $t }
    }
}

$cut = (Get-Date).AddMonths(-$Months).ToString('yyyy/MM/dd')
if (-not $Out) {
    $desk = [Environment]::GetFolderPath('Desktop')
    if (-not $desk) { $desk = $PSScriptRoot }
    $Out = Join-Path $desk ('testdb_{0}.sql' -f (Get-Date -Format 'yyyyMMdd'))
}

$conn = Open-Db
$sw = $null
try {
    if ($NoMask) {
        Write-Host '[注意] -NoMask 指定です。実名のまま書き出します。持ち出さないでください。' -ForegroundColor Red
    } else {
        Write-Host '[仮名化] 氏名・住所などは架空の値に置き換えて書き出します。' -ForegroundColor Green
    }
    Write-Host ("[期間] {0} 以降の受診データ ({1} か月分)" -f $cut, $Months) -ForegroundColor Cyan

    $cnt = Invoke-DbQuery $conn "SELECT COUNT(*) AS N, COUNT(DISTINCT KOJIN_ID) AS H FROM T_KENSIN WHERE D_KENSIN >= @c" @{ c = $cut }
    $nKen = [int]$cnt.Rows[0].N
    if ($nKen -eq 0) { throw "この期間に受診データがありません: $cut 以降" }
    Write-Host ("[対象] 受診 {0} 件 / 人 {1} 人" -f $nKen, [int]$cnt.Rows[0].H) -ForegroundColor Cyan

    $sw = New-Object System.IO.StreamWriter($Out, $false, (New-Object System.Text.UTF8Encoding($true)))
    $sw.WriteLine("-- 練習用DB 作成スクリプト")
    $sw.WriteLine("-- 作成日時: $(Get-Date -Format 'yyyy/MM/dd HH:mm')")
    $sw.WriteLine("-- 期間    : $cut 以降 / 受診 $nKen 件")
    $sw.WriteLine("-- 仮名化  : $(if ($NoMask) { 'していない (持ち出し厳禁)' } else { '済み' })")
    $sw.WriteLine("--")
    $sw.WriteLine("-- 使い方: 空のDBを作ってから、このファイルを丸ごと実行する")
    $sw.WriteLine("--   CREATE DATABASE K166_SIBAURAUSER COLLATE Japanese_CI_AS")
    $sw.WriteLine("--   USE K166_SIBAURAUSER")
    $sw.WriteLine("--   (このファイルを開いて実行)")
    $sw.WriteLine("SET NOCOUNT ON;")
    $sw.WriteLine("GO")
    $sw.WriteLine("")

    $total = 0
    foreach ($t in $TABLES) {
        $name = $t.名
        $ex = Invoke-DbQuery $conn "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @t AND TABLE_TYPE = 'BASE TABLE'" @{ t = $name }
        if ([int]$ex.Rows[0].N -eq 0) {
            # 実体が別DBにあるシノニムやビューのことがある。読めるなら中身を写し取る。
            if (Invoke-DbQuery $conn "SELECT 1 AS N WHERE OBJECT_ID(@t) IS NOT NULL" @{ t = $name } | ForEach-Object { $_.Rows.Count } | Where-Object { $_ -gt 0 }) {
                Write-Host ("  {0,-16} (シノニム/ビュー)" -f $name) -ForegroundColor DarkCyan
                $t.実体なし = $true
            } else {
                Write-Host ("  [skip] {0} … 無い" -f $name) -ForegroundColor DarkGray
                continue
            }
        }

        $cols = Invoke-DbQuery $conn @"
SELECT c.COLUMN_NAME, c.DATA_TYPE, c.CHARACTER_MAXIMUM_LENGTH,
       c.NUMERIC_PRECISION, c.NUMERIC_SCALE, c.DATETIME_PRECISION, c.IS_NULLABLE,
       COLUMNPROPERTY(OBJECT_ID(QUOTENAME(c.TABLE_SCHEMA)+'.'+QUOTENAME(c.TABLE_NAME)), c.COLUMN_NAME, 'IsIdentity') AS IS_IDENT,
       COLUMNPROPERTY(OBJECT_ID(QUOTENAME(c.TABLE_SCHEMA)+'.'+QUOTENAME(c.TABLE_NAME)), c.COLUMN_NAME, 'IsComputed') AS IS_COMP
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_NAME = @t ORDER BY c.ORDINAL_POSITION
"@ @{ t = $name }

        # 主キー (無いと健診ナビ側の更新でこけることがある)
        $pk = Invoke-DbQuery $conn @"
SELECT k.COLUMN_NAME
FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE k
  ON k.CONSTRAINT_NAME = tc.CONSTRAINT_NAME AND k.TABLE_NAME = tc.TABLE_NAME
WHERE tc.TABLE_NAME = @t AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
ORDER BY k.ORDINAL_POSITION
"@ @{ t = $name }

        $defs = @(); $sel = @(); $colNames = @(); $colTypes = @{}
        $hasIdent = $false
        foreach ($c in $cols.Rows) {
            $cn = [string]$c.COLUMN_NAME
            $ty = ([string]$c.DATA_TYPE).ToLower()
            $isIdent = ($c.IS_IDENT -isnot [DBNull]) -and ([int]$c.IS_IDENT -eq 1)
            if ($isIdent) { $hasIdent = $true }

            $nl = if ([string]$c.IS_NULLABLE -eq 'NO') { 'NOT NULL' } else { 'NULL' }
            # rowversion は入れ直せないので NULL 可にしておく
            if ($ty -eq 'timestamp' -or $ty -eq 'rowversion') { $nl = 'NULL' }
            $id = if ($isIdent) { ' IDENTITY(1,1)' } else { '' }
            $defs += ("  [{0}] {1}{2} {3}" -f $cn, (Get-ColumnTypeSql $c), $id, $nl)

            $colNames += $cn
            $colTypes[$cn] = $ty
            $key = "$name.$cn"
            if ((-not $NoMask) -and $MASK.ContainsKey($key)) { $sel += ("{0} AS [{1}]" -f $MASK[$key], $cn) }
            else { $sel += "[$cn]" }
        }
        if ($pk.Rows.Count -gt 0) {
            $pkc = @($pk.Rows | ForEach-Object { '[' + [string]$_.COLUMN_NAME + ']' }) -join ','
            $defs += ("  CONSTRAINT [PK_{0}] PRIMARY KEY ({1})" -f $name, $pkc)
        }

        $sw.WriteLine("-- ==== $name ($($t.区分)) ====")
        $sw.WriteLine("IF OBJECT_ID('$name') IS NOT NULL DROP TABLE [$name];")
        $sw.WriteLine("GO")
        $sw.WriteLine("CREATE TABLE [$name] (")
        $sw.WriteLine(($defs -join ",`r`n"))
        $sw.WriteLine(");")
        $sw.WriteLine("GO")
        if ($hasIdent) { $sw.WriteLine("SET IDENTITY_INSERT [$name] ON;") ; $sw.WriteLine("GO") }

        # データ (件数が多いので、ためこまずに流しながら書く)
        $where = [string]$t.Filter
        if ($where -ne '') { $where = ' WHERE ' + $where.Replace('@CUT', $cut) }
        $q = "SELECT {0} FROM [{1}]{2}" -f ($sel -join ', '), $name, $where

        $head = "INSERT INTO [$name] ([" + ($colNames -join '],[') + "]) VALUES"
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $q
        $cmd.CommandTimeout = 0
        $rd = $cmd.ExecuteReader()
        $n = 0; $inBatch = 0
        try {
            while ($rd.Read()) {
                $vals = New-Object 'System.Collections.Generic.List[string]'
                for ($i = 0; $i -lt $colNames.Count; $i++) {
                    $vals.Add((ConvertTo-SqlLiteral $rd.GetValue($i) $colTypes[$colNames[$i]]))
                }
                if ($inBatch -eq 0) { $sw.WriteLine($head) } else { $sw.WriteLine(',') }
                $sw.Write('(' + ($vals -join ',') + ')')
                $n++; $inBatch++
                if ($inBatch -ge 200) { $sw.WriteLine(';'); $sw.WriteLine('GO'); $inBatch = 0 }
            }
        }
        finally { $rd.Close(); $cmd.Dispose() }
        if ($inBatch -gt 0) { $sw.WriteLine(';'); $sw.WriteLine('GO') }

        if ($hasIdent) { $sw.WriteLine("SET IDENTITY_INSERT [$name] OFF;"); $sw.WriteLine("GO") }
        $sw.WriteLine("")
        $total += $n
        Write-Host ("  {0,-16} {1,8} 行" -f $name, $n) -ForegroundColor Green
    }

    $sw.WriteLine("-- 合計 $total 行")
    $sw.Flush()
}
finally {
    if ($sw) { $sw.Close() }
    $conn.Close()
}

$mb = [math]::Round((Get-Item $Out).Length / 1MB, 1)
Write-Host ''
Write-Host ("出力: {0}  ({1} MB)" -f $Out, $mb) -ForegroundColor Cyan
Write-Host ''
Write-Host '自分のPCでの使い方:' -ForegroundColor Yellow
Write-Host '  1. SSMS で localhost\SQLEXPRESS に接続し、新しいクエリで次を実行'
Write-Host '       CREATE DATABASE K166_SIBAURAUSER COLLATE Japanese_CI_AS'
Write-Host '  2. このSQLファイルを流し込む (ファイルが大きいと SSMS で開けないので sqlcmd が確実)'
Write-Host ('       sqlcmd -S localhost\SQLEXPRESS -d K166_SIBAURAUSER -i "{0}"' -f $Out)
Write-Host '  3. kenshin-navi-tool フォルダに 接続先.txt を置いて、そのDBを向ける'
