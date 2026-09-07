<#
  健診ナビの「自動判定 → 登録」をくりかえす (自動判定くりかえし.ps1)
  自動判定くりかえし.bat をダブルクリックして使います。

  やっていること
    人がマウスで押すのと同じことを、決めた回数だけくりかえすだけです。
    データベースには一切触れません。判定の中身は健診ナビが出したものそのままです。

  前提
    ・健診ナビの「結果入力」画面を開いていること
    ・「編集モード」にチェックが入っていること
    ・「登録後、次へ」にチェックが入っていること (登録すると次の人に進む)
    ・健診ナビの画面を動かさないこと (ボタンの位置がずれます)

  止め方
    ESCキーを押す。または黒い画面を閉じる。
#>
[CmdletBinding()]
param(
    [int]$Count = 0,          # 何人分くりかえすか
    [int]$WaitHantei = 1800,  # 自動判定を押したあと待つ時間(ミリ秒)
    [int]$WaitToroku = 2200   # 登録を押したあと待つ時間(ミリ秒)
)
$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, int e);
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(120);
        mouse_event(0x0002, 0, 0, 0, 0);   // 左ボタンを押す
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, 0, 0, 0, 0);   // 左ボタンを離す
    }
    public static POINT Where() { POINT p; GetCursorPos(out p); return p; }
}
'@

function Line { Write-Host ('-' * 66) -ForegroundColor DarkGray }

Write-Host ''
Write-Host ('=' * 66) -ForegroundColor Cyan
Write-Host ' 健診ナビ 自動判定のくりかえし' -ForegroundColor Cyan
Write-Host ('=' * 66) -ForegroundColor Cyan
Write-Host ''
Write-Host '  人がマウスで押すのと同じことを、決めた回数くりかえします。'
Write-Host '  データベースには触りません。判定は健診ナビが出したものです。'
Write-Host ''
Write-Host '  始める前に、健診ナビの結果入力画面で次を確かめてください。' -ForegroundColor Yellow
Write-Host '    ・「編集モード」にチェックが入っている' -ForegroundColor Yellow
Write-Host '    ・「登録後、次へ」にチェックが入っている' -ForegroundColor Yellow
Write-Host '    ・これから判定したい人が画面に出ている' -ForegroundColor Yellow
Write-Host ''
Write-Host '  途中で止めたいときは ESC キーを押してください。' -ForegroundColor Yellow
Write-Host ''

# ---- ボタンの位置を覚える ----
function Read-Point([string]$name) {
    Write-Host ''
    Write-Host "  ★「$name」ボタンの上にマウスを置いて、そのまま待ってください。" -ForegroundColor White
    Write-Host '    クリックはしないでください。5秒後の位置をそのまま覚えます。'
    Write-Host '    のこり...' -NoNewline
    for ($i = 5; $i -ge 1; $i--) { Write-Host " $i" -NoNewline; Start-Sleep -Seconds 1 }
    $p = [Mouse]::Where()
    Write-Host ''
    Write-Host ("    覚えました: X={0} Y={1}" -f $p.X, $p.Y) -ForegroundColor Green
    return $p
}

Line
$pHantei = Read-Point '自動判定'
$pToroku = Read-Point '登録'

if ($pHantei.X -eq $pToroku.X -and $pHantei.Y -eq $pToroku.Y) {
    Write-Host ''
    Write-Host '  ★ 2つの位置が同じです。マウスを動かせていません。やり直してください。' -ForegroundColor Red
    Write-Host ''
    return
}

# ---- 回数 ----
if ($Count -le 0) {
    Write-Host ''
    Line
    Write-Host '  何人分くりかえしますか。'
    Write-Host '  はじめは 3 と入れて、ちゃんと動くか目で確かめてください。' -ForegroundColor Yellow
    $ans = Read-Host '  人数'
    $Count = 0
    [void][int]::TryParse(($ans -replace '[^\d]', ''), [ref]$Count)
}
if ($Count -le 0) { Write-Host '  人数が読み取れませんでした。終わります。'; return }
if ($Count -gt 200) { Write-Host '  200人までにしてください。'; return }

Write-Host ''
Line
Write-Host ('  これから {0} 人分くりかえします。' -f $Count) -ForegroundColor Cyan
Write-Host ('    自動判定 X={0} Y={1}' -f $pHantei.X, $pHantei.Y)
Write-Host ('    登録     X={0} Y={1}' -f $pToroku.X, $pToroku.Y)
Write-Host ''
Write-Host '  健診ナビの画面を前に出して、そのまま触らずに待ってください。' -ForegroundColor Yellow
Write-Host '  10秒後に始めます。やめるなら今この画面を閉じてください。' -ForegroundColor Yellow
Write-Host '  カウント中...' -NoNewline
for ($i = 10; $i -ge 1; $i--) { Write-Host " $i" -NoNewline; Start-Sleep -Seconds 1 }
Write-Host ''
Write-Host ''

# ---- くりかえし ----
$done = 0
for ($n = 1; $n -le $Count; $n++) {
    # ESCで中断
    try {
        if ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq 'Escape') {
                Write-Host ''
                Write-Host ('  ESCが押されました。{0} 人まで進んだところで止めます。' -f $done) -ForegroundColor Yellow
                break
            }
        }
    } catch { }   # 入力が使えない環境ではESC中断だけ効かない。窓を閉じれば止まる。
    Write-Host ("  {0,3} / {1}  自動判定..." -f $n, $Count) -NoNewline
    [Mouse]::Click($pHantei.X, $pHantei.Y)
    Start-Sleep -Milliseconds $WaitHantei
    Write-Host ' 登録...' -NoNewline
    [Mouse]::Click($pToroku.X, $pToroku.Y)
    Start-Sleep -Milliseconds $WaitToroku
    $done++
    Write-Host ' 済' -ForegroundColor Green
}

Write-Host ''
Line
Write-Host ('  {0} 人分やりました。' -f $done) -ForegroundColor Green
Write-Host ''
Write-Host '  健診ナビの画面を見て、ちゃんと人が進んでいるか確かめてください。' -ForegroundColor Yellow
Write-Host '  おかしければ、その人を開いて手で直してください。'
Write-Host '  (このツールは押しただけなので、健診ナビ側で直せば元に戻ります)'
Write-Host ''
