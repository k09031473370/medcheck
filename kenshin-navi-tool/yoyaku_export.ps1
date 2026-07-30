<#
.SYNOPSIS
  当日フォーム → 予約取込ファイル(Excel) 生成 (yoyaku_export.ps1)

.DESCRIPTION
  当日健診フォーム(106列)から、健診ナビの「予約データ取込」で読み込める
  Excelファイルを作る。テンプレート(予約取込フォーマット.xlsx)の
  「予約取込(原本)」シートを複製し、3行目以降にデータを書き込む。

  DBには一切アクセスしない。生成したファイルをアプリの予約データ取込に読ませる。

  設定: form\yoyaku_settings.csv (住所・電話など全員共通の値)
        form\yoyaku_course.csv   (フォームのコース → コース名/コード)

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File yoyaku_export.ps1 -Csv <form.xlsx or .csv>
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Csv,   # 当日フォーム (.xlsx / .csv)
    [string]$Template,                            # 予約取込フォーマット.xlsx
    [string]$Out,                                 # 出力先 (既定: デスクトップ)
    [switch]$NoHeader,                            # フォームに見出し行が無い場合
    [string]$MapDir,
    [ValidateSet('SJIS','UTF8')][string]$CsvEncoding = 'SJIS'
)

$ErrorActionPreference = 'Stop'
if (-not $MapDir) { $MapDir = Join-Path $PSScriptRoot 'form' }
try { [void][System.Text.Encoding]::GetEncoding(932) }
catch { [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance) }

# .xlsx / .xlsm の読み込み (Excelを使わない共通部品)
$XlsxLib = Join-Path $PSScriptRoot 'xlsx_read.ps1'
if (-not (Test-Path $XlsxLib)) { throw "xlsx_read.ps1 が見つかりません: $XlsxLib" }
. $XlsxLib

function Normalize-Text([string]$s) { if ($null -eq $s) { return '' } return $s.Trim() }

# ---- 設定の読み込み ----
function Load-Settings {
    $p = Join-Path $MapDir 'yoyaku_settings.csv'
    $h = @{}
    if (Test-Path $p) {
        foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
            $k = Normalize-Text $r.Key
            if ($k -ne '') { $h[$k] = Normalize-Text $r.Value }
        }
    }
    return $h
}
function Load-CourseMap {
    $p = Join-Path $MapDir 'yoyaku_course.csv'
    $h = @{}
    if (Test-Path $p) {
        foreach ($r in (Import-Csv -Path $p -Encoding UTF8)) {
            $k = Normalize-Text $r.FromCourse
            if ($k -ne '') { $h[$k] = @{ Name = (Normalize-Text $r.CourseName); Code = (Normalize-Text $r.CourseCode) } }
        }
    }
    return $h
}

# ---- フォームの読み込み ----
function Read-FormRows([string]$path) {
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    if ($ext -eq '.xlsx' -or $ext -eq '.xlsm') {
        # Excelを起動せずに読む (日付は yyyy/MM/dd の文字列で返る)
        return ,(Read-Xlsx $path)
    }
    if ($ext -eq '.xls') {
        throw "古い形式(.xls)は読めません。Excelで開いて .xlsx として保存し直してください。`n$path"
    }
    # CSV
    $enc = if ($CsvEncoding -eq 'UTF8') { New-Object System.Text.UTF8Encoding($false) } else { [System.Text.Encoding]::GetEncoding(932) }
    $text = [System.IO.File]::ReadAllText($path, $enc)
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    $rows = @()
    foreach ($ln in ($text -split "`r?`n")) {
        if ((Normalize-Text $ln) -eq '') { continue }
        $rows += ,($ln -split ',')
    }
    return ,$rows
}

# テンプレートの見出し名から列番号を引く (前方一致も許す)
#   例: 「住所1」→「住所1(都道府県)」にも一致する
function ColIdx($colOf, [string]$name) {
    if ($colOf.ContainsKey($name)) { return [int]$colOf[$name] }
    foreach ($k in $colOf.Keys) { if ($k -like ($name + '*')) { return [int]$colOf[$k] } }
    return 0
}

function F($row, [int]$col) {
    if ($col -le 0 -or $col -gt $row.Count) { return '' }
    return Normalize-Text ([string]$row[$col - 1])
}

# ============================================================================
# メイン
# ============================================================================

if (-not $Template) {
    $cand = @(Get-ChildItem -Path $PSScriptRoot -Filter '*予約取込*.xlsx' -File -ErrorAction SilentlyContinue)
    if ($cand.Count -eq 0) {
        throw "予約取込フォーマットのExcelが見つかりません。ツールと同じフォルダに置くか -Template で指定してください。"
    }
    $Template = $cand[0].FullName
}
if (-not (Test-Path $Template)) { throw "テンプレートが見つかりません: $Template" }
if (-not (Test-Path $Csv)) { throw "フォームが見つかりません: $Csv" }

$settings = Load-Settings
$courses  = Load-CourseMap

$rows = Read-FormRows $Csv
if ($rows.Count -eq 0) { throw 'フォームにデータがありません。' }
$data = if ($NoHeader) { $rows } else { $rows[1..($rows.Count - 1)] }

# 出力先
if (-not $Out) {
    $ymd = (F $data[0] 3) -replace '[/\-]', ''
    if ($ymd -eq '') { $ymd = Get-Date -Format 'yyyyMMdd' }
    $Out = Join-Path ([Environment]::GetFolderPath('Desktop')) ("予約取込_{0}.xlsx" -f $ymd)
}

Copy-Item -Path $Template -Destination $Out -Force
Write-Host "[生成] $Out" -ForegroundColor Cyan

$excel = $null; $wb = $null
$warn = @()
$no = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    # ダイアログ抑止。環境によって受け付けないプロパティがあるので、失敗しても続行する
    foreach ($kv in @(@{N='AskToUpdateLinks';V=$false}, @{N='EnableEvents';V=$false}, @{N='AutomationSecurity';V=3})) {
        try { $excel.($kv.N) = $kv.V } catch { Write-Host ("  (設定 {0} はこの環境では使えません)" -f $kv.N) -ForegroundColor DarkGray }
    }
    # 自分でコピーしたファイルなので、余計な引数は付けずに開く
    $wb = $excel.Workbooks.Open($Out)
    $ws = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -like '*原本*') { $ws = $s; break } }
    if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
    Write-Host "[シート] $($ws.Name)" -ForegroundColor DarkGray

    # ---- テンプレートの見出しを読んで、列の位置を名前で決める ----
    #  テンプレートによって「年齢」「社員番号」「所属名」の有無や見出し行が違うため、
    #  列番号を決め打ちにしない。
    $probe = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(6, 40)).Value2
    $hdrRow = 0
    $colOf = @{}
    for ($pr = 1; $pr -le 6; $pr++) {
        $found = @{}
        for ($pc = 1; $pc -le 40; $pc++) {
            $v = $probe.GetValue($pr, $pc)
            if ($null -eq $v) { continue }
            $t = ([string]$v).Trim()
            if ($t -ne '') { $found[$t] = $pc }
        }
        if ($found.ContainsKey('氏名') -and (@($found.Keys | Where-Object { $_ -like '生年月日*' }).Count -gt 0)) {
            $hdrRow = $pr
            $colOf = $found
            break
        }
    }
    if ($hdrRow -eq 0) {
        throw "予約取込フォーマットの見出し行が見つかりません。1〜6行目に「氏名」「生年月日」を含む行が必要です。"
    }
    $startRow = $hdrRow + 1
    Write-Host ("[見出し] {0} 行目 / {1} 列を認識 (データは {2} 行目から)" -f $hdrRow, $colOf.Count, $startRow) -ForegroundColor DarkGray

    # 既存のサンプル行を消す
    $last = $ws.UsedRange.Rows.Count
    if ($last -ge $startRow) { [void]$ws.Range($ws.Cells.Item($startRow, 1), $ws.Cells.Item($last, 40)).ClearContents() }

    # 1人1行ぶんの値を「見出し名 → 値」で組み立てる
    $lines = @()
    foreach ($row in $data) {
        $name = F $row 7
        if ($name -eq '') { continue }
        $no++

        $courseRaw = F $row 106
        $cName = ''; $cCode = ''
        if ($courseRaw -ne '') {
            if ($courses.ContainsKey($courseRaw)) { $cName = $courses[$courseRaw].Name; $cCode = $courses[$courseRaw].Code }
            else { $warn += "コース「$courseRaw」の変換が未設定 ($name)"; $cName = $courseRaw }
        }
        else { $warn += "コースが空欄 ($name)" }

        $hokensya = F $row 14
        $kenpoName = ''
        if ($hokensya -ne '') { $kenpoName = $settings['健保名'] }

        $lines += ,@{
            'No'         = [int]$no
            '氏名'       = [string]$name
            '氏名カナ'   = [string](F $row 8)
            '性別'       = [string](F $row 9)
            '生年月日'   = [string](F $row 11)
            '年齢'       = [string](F $row 10)
            '郵便番号'   = [string]$settings['郵便番号']
            '住所1'      = [string]$settings['住所1']
            '住所2'      = [string]$settings['住所2']
            '住所3'      = [string]$settings['住所3']
            '電話番号'   = [string]$settings['電話番号']
            'カルテ番号' = ''
            '社員番号'   = [string](F $row 1)
            '保険者番号' = [string]$hokensya
            '保険証記号' = [string](F $row 15)
            '保険証番号' = [string](F $row 16)
            '保険証枝番号' = ''
            '健保コード' = [string]$settings['健保コード']
            '健保名'     = [string]$kenpoName
            '事業所コード' = [string]$settings['事業所コード']
            '事業所名'   = [string](F $row 5)
            '所属名'     = [string](F $row 13)
            'コースコード' = [string]$cCode
            'コース名'   = [string]$cName
            '予約日'     = [string](F $row 3)
            '予約時間'   = [string]$settings['予約時間']
        }
    }

    # 見出し名 → 列番号 に置き換えて、まとめて1回で書き込む
    if ($lines.Count -gt 0) {
        $place = @{}
        $maxCol = 0
        foreach ($k in $lines[0].Keys) {
            $ci = ColIdx $colOf $k
            if ($ci -gt 0) {
                $place[$k] = $ci
                if ($ci -gt $maxCol) { $maxCol = $ci }
            }
            else { $warn += "テンプレートに「$k」の列がないため、書き出していません" }
        }
        $arr = New-Object 'object[,]' $lines.Count, $maxCol
        for ($i = 0; $i -lt $lines.Count; $i++) {
            foreach ($k in $place.Keys) { $arr[$i, ($place[$k] - 1)] = $lines[$i][$k] }
        }
        $rng = $ws.Range($ws.Cells.Item($startRow, 1), $ws.Cells.Item($startRow + $lines.Count - 1, $maxCol))
        $rng.Value2 = $arr
    }

    $wb.Save()
    Write-Host ("[完了] {0} 人分を書き出しました。" -f $no) -ForegroundColor Green
}
catch {
    $inv = $_.InvocationInfo
    Write-Host '' 
    Write-Host '[失敗] 予約取込ファイルの書き出しでエラーが起きました。' -ForegroundColor Red
    Write-Host ("  内容 : {0}" -f $_.Exception.Message) -ForegroundColor Red
    if ($inv) {
        Write-Host ("  場所 : {0} 行目" -f $inv.ScriptLineNumber) -ForegroundColor Red
        Write-Host ("  該当 : {0}" -f ($inv.Line).Trim()) -ForegroundColor Red
    }
    Write-Host ("  途中まで書けた人数: {0}" -f $no) -ForegroundColor DarkGray
    throw
}
finally {
    if ($wb) { $wb.Close($true) | Out-Null }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}

if ($warn.Count -gt 0) {
    Write-Host ''
    Write-Host '--- 確認が必要な項目 ---' -ForegroundColor Yellow
    $warn | Select-Object -Unique | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
}
Write-Host ''
Write-Host '※ 生成したExcelを開いて内容を確認してから、健診ナビの「予約データ取込」で読み込んでください。'
