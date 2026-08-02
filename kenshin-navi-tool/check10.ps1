<#
  帳票で使える差込み文字を全部あつめる (check10.ps1)
  check10.bat をダブルクリックすると実行され、結果 r_check10.txt がメモ帳で開きます。
  帳票テンプレートを読むだけで、一切変更しません。

  なぜ必要か:
    個人結果票(101)には「社員番号」の差込み文字が無かった。
    差込み文字の名前は健診ナビが決めているので、こちらで勝手に書いても出ない。
    他の帳票に社員番号や部署名の差込み文字が用意されていれば、
    それを結果票に足すだけで印字できる。
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$out = Join-Path $dir 'r_check10.txt'
$reader = Join-Path $dir 'xlsx_read.ps1'

function W($t) { $t | Out-File $out -Append -Encoding Default }

"=== 帳票で使える差込み文字 一覧 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default

if (-not (Test-Path $reader)) { W "xlsx_read.ps1 がありません: $dir"; notepad $out; return }
. $reader

# 帳票が置いてある場所をさがす
$roots = @(
    '\\KNSV\KenshinNavi\11_健康診断結果報告書',
    '\\KNSV\KenshinNavi',
    'C:\KenshinNavi', 'D:\KenshinNavi'
)
$root = $null
foreach ($r in $roots) { if (Test-Path $r) { $root = $r; break } }
if (-not $root) { W '帳票フォルダが見つかりませんでした。パスを教えてください。'; notepad $out; return }

W "対象フォルダ: $root"
$files = @(Get-ChildItem -Path $root -Filter '*.xls*' -Recurse -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -notlike '~$*' } | Sort-Object FullName)
W ("帳票ファイル: {0} 個" -f $files.Count)

# 差込み文字 → それが使われている帳票
$tokens = @{}
$failed = @()
foreach ($f in $files) {
    try { $rows = Read-Xlsx $f.FullName } catch { $failed += ('{0} ({1})' -f $f.Name, $_.Exception.Message); continue }
    foreach ($row in $rows) {
        foreach ($cell in $row) {
            $v = [string]$cell
            if ($v -eq '' -or $v -notlike '*`**') { continue }
            foreach ($mm in [regex]::Matches($v, '\*\*[^\s|]+')) {
                $t = $mm.Value
                if (-not $tokens.ContainsKey($t)) { $tokens[$t] = New-Object System.Collections.Generic.HashSet[string] }
                [void]$tokens[$t].Add($f.Name)
            }
        }
    }
}

# 結果値・判定は数が多いので、それ以外(人や事業所の情報)を先に出す
$keys = @($tokens.Keys | Sort-Object)
$person = @($keys | Where-Object { $_ -notmatch '結果値|判定|基準値|単位|項目名称|HL_' })
$rest   = @($keys | Where-Object { $_ -match '結果値|判定|基準値|単位|項目名称|HL_' })

W ''
W '--- 1. 人・事業所まわりの差込み文字 (ここに社員番号があれば使えます) ---'
foreach ($k in $person) { W ('{0,-40} {1}' -f $k, (($tokens[$k] | Sort-Object) -join ', ')) }

W ''
W ('--- 2. 検査結果まわりの差込み文字 ({0}種。参考) ---' -f $rest.Count)
foreach ($k in $rest) { W ('  {0}' -f $k) }

if ($failed.Count -gt 0) {
    W ''
    W '--- 読めなかったファイル ---'
    foreach ($x in $failed) { W ("  {0}" -f $x) }
}

W ''
W '※ 1に「社員番号」「個人番号」「部署」らしいものがあれば、それを結果票に足せば印字できます。'
W '=== 完了。この内容(特に1)をチャットに貼り付けてください ==='
notepad $out
