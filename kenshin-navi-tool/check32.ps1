<#
  受付処理が何を書いているかを調べる (check32.ps1)
  check32.bat をダブルクリックすると実行されます。
  DBは読むだけで、一切変更しません。

  使い方 (2回実行します):
    1回目 … 受付する前に実行。状態を控えます (r_check32_before.txt)
    2回目 … 健診ナビで1人だけ受付してから実行。
             前回との違いを r_check32.txt に出します。

  これが分かれば、ツールから98人まとめて受付する仕組みを作れます。
  分からないまま真似すると、健診ナビが期待する形と食い違って
  あとで面倒なことになるので、必ず実物を見てから作ります。
#>
[CmdletBinding()]
param(
    [string]$Ymd = '2026/08/21',   # 対象の受診日
    [switch]$Reset                 # 控えを取り直したいとき
)
$ErrorActionPreference = 'Continue'
$dir    = $PSScriptRoot
$tool   = Join-Path $dir 'db_tool.ps1'
$out    = Join-Path $dir 'r_check32.txt'
$before = Join-Path $dir 'r_check32_before.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }

if (-not (Test-Path $tool)) { "db_tool.ps1 がありません: $dir" | Out-File $out -Encoding Default; notepad $out; return }
if ($Reset -and (Test-Path $before)) { Remove-Item $before -Force }

# 受診・受付まわりの状態を1行1人で書き出す
$sql = @"
SELECT s.PK_SEQ,
       '[' + ISNULL(s.SEQ1,'') + ']' AS SEQ1,
       '[' + ISNULL(s.SEQ2,'') + ']' AS SEQ2,
       '[' + ISNULL(CONVERT(varchar(20), s.UKE_NO_KENSA),'') + ']' AS 受付番号,
       s.F_UKETUKE AS 受付F,
       (SELECT COUNT(*) FROM T_KANJA_G g WHERE g.PK_SEQ = s.PK_SEQ) AS 受付行数,
       (SELECT TOP 1 '[' + ISNULL(g.KEN_NO,'') + ']' FROM T_KANJA_G g WHERE g.PK_SEQ = s.PK_SEQ) AS 受付KEN_NO,
       (SELECT TOP 1 '[' + ISNULL(g.KEN_YMD,'') + ']' FROM T_KANJA_G g WHERE g.PK_SEQ = s.PK_SEQ) AS 受付KEN_YMD,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
          AND LTRIM(RTRIM(ISNULL(k.SEQ1,''))) <> '') AS 検査SEQ1入り,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS 検査行数
FROM T_KENSIN s
WHERE s.D_KENSIN = '$Ymd' AND s.F_TORIKESI = 0
ORDER BY s.PK_SEQ
"@

$snap = & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows 500 *>&1

if (-not (Test-Path $before)) {
    $snap | Out-File $before -Encoding Default
    "=== 受付処理の調査 1回目 (控えを取りました) $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
    W ''
    W "受診日 $Ymd の状態を控えました。"
    W "  控えの場所: $before"
    W ''
    W '=== 次にやること ==='
    W '  1. 健診ナビの予約画面で、誰か1人だけ「受付」してください。'
    W '     (あとで「取消」で戻せます)'
    W '  2. もう一度 check32.bat をダブルクリックしてください。'
    W '     受付で何が変わったかが出ます。'
    W ''
    W '  ※ 控えを取り直したいときは r_check32_before.txt を消してから実行してください。'
    W ''
    W '--- いまの状態 (先頭のほう) ---'
    $snap | Select-Object -First 25 | Out-File $out -Append -Encoding Default
    notepad $out
    return
}

# 2回目: 差分を出す
"=== 受付処理で何が変わったか $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
W ''

$oldLines = @(Get-Content $before -Encoding Default)
$newLines = @($snap | ForEach-Object { [string]$_ })

# PK_SEQ で始まる行だけを拾って突き合わせる
function ToMap($lines) {
    $h = @{}
    foreach ($l in $lines) {
        if ($l -match '^\s*(\d{6,})\s') { $h[$Matches[1]] = ($l -replace '\s+', ' ').Trim() }
    }
    return $h
}
$o = ToMap $oldLines
$n = ToMap $newLines

W ("控え {0} 人 / いま {1} 人" -f $o.Count, $n.Count)
W ''

$diff = 0
foreach ($pk in ($n.Keys | Sort-Object)) {
    if (-not $o.ContainsKey($pk)) { W "★ 新しく増えた: $pk"; $diff++; continue }
    if ($o[$pk] -ne $n[$pk]) {
        $diff++
        W "★ 変わった人 PK_SEQ $pk"
        W ("   受付前: {0}" -f $o[$pk])
        W ("   受付後: {0}" -f $n[$pk])
        W ''
    }
}
if ($diff -eq 0) {
    W '違いはありませんでした。'
    W '  健診ナビでまだ受付していないか、別の受診日を見ている可能性があります。'
    W ("  いま見ている受診日: {0}" -f $Ymd)
}
else {
    W ("違いのあった人: {0} 人" -f $diff)
}

W ''
W '=== 読み方 ==='
W '  SEQ1 に値が入った          → 受付で SEQ1 が決まる。値の形(受診日8桁+4桁)を見ます。'
W '  受付番号 が変わった          → 受付で採番される。ツールで入れた番号が上書きされる恐れあり。'
W '  受付行数 が 0 から 1 になった  → T_KANJA_G に行が作られる。ここが受付の本体です。'
W '  検査SEQ1入り が増えた        → T_KENSA の全行にも SEQ1 が書かれる。'
W '  受付F が 0 から 1 になった    → フラグも立つ。'
W ''
W '  この4つが分かれば、ツールから98人まとめて受付する仕組みを作れます。'
W '  ただし T_KANJA_G には他にも列があるはずなので、'
W '  作る前にもう一度その中身を見せてもらいます。'

notepad $out
