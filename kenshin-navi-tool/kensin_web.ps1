<#
.SYNOPSIS
  健診ナビのDBを見る自作画面 (kensin_web.ps1)

.DESCRIPTION
  健診ナビと同じDBを読んで、ブラウザで見られる画面を出す。
  インストールは要らない。Windowsに元から入っている機能だけで動く。

  この版は「読むだけ」。DBには一切書き込まない。
  使い勝手を確かめてから、書き込む画面を足していく。

  安全のため 127.0.0.1 (自分のPC) からしか繋がらないようにしてある。
  他のPCから見えることはない。

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File kensin_web.ps1
  powershell -ExecutionPolicy Bypass -File kensin_web.ps1 -Port 8080
#>
[CmdletBinding()]
param(
    [int]$Port = 8080,
    [switch]$NoBrowser,
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)

$ErrorActionPreference = 'Stop'

# form_import.ps1 から接続まわりを借りる (同じ接続先を使うため)
$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Normalize-KenNo')
Invoke-Expression (Get-Part 'function Normalize-Ymd' 'function Parse-CsvText')
Invoke-Expression (Get-Part 'function Resolve-ConnectionString' 'function Get-CurrentKensa')

# ============================================================================
# HTML の部品
# ============================================================================

function HtmlEnc([string]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

$Style = @'
<style>
* { box-sizing: border-box; }
body { font-family: "Yu Gothic UI","Meiryo",sans-serif; margin:0; background:#f4f6f8; color:#222; }
header { background:#1f4e79; color:#fff; padding:10px 16px; display:flex; align-items:center; gap:16px; }
header h1 { font-size:16px; margin:0; font-weight:600; }
header .env { margin-left:auto; font-size:12px; opacity:.85; }
nav { background:#2d6ca2; padding:0 16px; }
nav a { color:#fff; display:inline-block; padding:8px 14px; text-decoration:none; font-size:14px; }
nav a:hover, nav a.on { background:#1f4e79; }
main { padding:16px; }
.bar { background:#fff; border:1px solid #d7dde3; border-radius:6px; padding:10px 12px; margin-bottom:12px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
.bar input[type=text], .bar input[type=date] { padding:6px 8px; border:1px solid #bcc6d0; border-radius:4px; font-size:14px; }
.bar button { padding:6px 16px; border:0; border-radius:4px; background:#1f4e79; color:#fff; font-size:14px; cursor:pointer; }
.bar button:hover { background:#163a5a; }
table { border-collapse:collapse; width:100%; background:#fff; font-size:13px; }
th, td { border:1px solid #dde3e9; padding:5px 8px; text-align:left; white-space:nowrap; }
th { background:#eaf0f6; position:sticky; top:0; font-weight:600; }
tr:nth-child(even) td { background:#fafbfc; }
tr:hover td { background:#fff6e0; }
td.num { text-align:right; }
a { color:#1f4e79; }
.pill { display:inline-block; padding:1px 8px; border-radius:10px; font-size:12px; }
.ok { background:#e2efda; color:#375623; }
.warn { background:#fde9d9; color:#833c00; }
.ng { background:#fbdada; color:#8b1a1a; }
.muted { color:#7a8797; }
.count { font-size:13px; color:#555; margin:8px 2px; }
.err { background:#fbdada; border:1px solid #e0a0a0; padding:12px; border-radius:6px; }
.note { font-size:12px; color:#666; margin-top:10px; line-height:1.7; }
.wrap { max-height:calc(100vh - 230px); overflow:auto; border:1px solid #d7dde3; border-radius:6px; }
</style>
'@

function Page([string]$title, [string]$body, [string]$nav, [string]$envLabel) {
    return @"
<!doctype html><html lang="ja"><head><meta charset="utf-8">
<title>$(HtmlEnc $title)</title>$Style</head><body>
<header><h1>健診ビューア</h1><span class="env">$(HtmlEnc $envLabel)</span></header>
<nav>$nav</nav>
<main>$body</main>
</body></html>
"@
}

# ============================================================================
# 画面
# ============================================================================

function Render-List($conn, [string]$ymd, [string]$kw) {
    $where = @('s.F_TORIKESI = 0')
    $prm = @{}
    if ($ymd -ne '') { $where += 's.D_KENSIN = @y'; $prm['y'] = $ymd }
    if ($kw  -ne '') { $where += "(j.KANJI_SIMEI LIKE '%' + @k + '%' OR j.KANA_SIMEI LIKE '%' + @k + '%' OR s.UKE_NO_KENSA LIKE '%' + @k + '%' OR d.MEISYO1 LIKE '%' + @k + '%')"; $prm['k'] = $kw }
    $sql = @"
SELECT TOP 500 s.PK_SEQ, s.D_KENSIN, s.UKE_NO_KENSA, j.KANJI_SIMEI, j.KANA_SIMEI,
       j.SEIBETU, j.D_SEINEN, d.MEISYO1 AS DANTAI, s.COURSE_CD, c1.MEISYO AS COURSE_MEI,
       s.D_JIDOHANTEI,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ) AS WAKU,
       (SELECT COUNT(*) FROM T_KENSA k WHERE k.PK_SEQ = s.PK_SEQ
          AND LTRIM(RTRIM(ISNULL(k.KEKKA,''))) <> '') AS DONE
FROM T_KENSIN s
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c1 ON c1.COURSE_CD = s.COURSE_CD AND c1.DANTAI_CD1 = s.DANTAI_CD1
WHERE $($where -join ' AND ')
ORDER BY s.D_KENSIN DESC, LEN(LTRIM(RTRIM(s.UKE_NO_KENSA))), s.UKE_NO_KENSA
"@
    $dt = Invoke-DbQuery $conn $sql $prm

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(@"
<form class="bar" method="get" action="/list">
  <label>受診日 <input type="text" name="ymd" value="$(HtmlEnc $ymd)" placeholder="2026/06/26" size="12"></label>
  <label>さがす <input type="text" name="kw" value="$(HtmlEnc $kw)" placeholder="氏名・受付番号・会社名" size="24"></label>
  <button type="submit">表示</button>
  <a href="/list">クリア</a>
</form>
"@)

    if ($dt.Rows.Count -eq 0) {
        [void]$sb.Append('<p class="count">該当する受診者がいません。</p>')
        return $sb.ToString()
    }

    $nDone = 0; $nPart = 0; $nNone = 0; $nHantei = 0
    $rows = New-Object System.Text.StringBuilder
    foreach ($r in $dt.Rows) {
        $waku = [int]$r.WAKU; $done = [int]$r.DONE
        $pct = if ($waku -gt 0) { [math]::Floor($done * 100.0 / $waku) } else { 0 }
        if ($waku -gt 0 -and $done -ge $waku) { $cls='ok'; $lbl='入力済'; $nDone++ }
        elseif ($done -gt 0)                  { $cls='warn'; $lbl="$pct%"; $nPart++ }
        else                                  { $cls='ng'; $lbl='未入力'; $nNone++ }
        $hantei = Normalize-Text ([string]$r.D_JIDOHANTEI)
        if ($hantei -ne '') { $nHantei++ }
        $sex = switch (Normalize-Text ([string]$r.SEIBETU)) { '1' {'男'} '2' {'女'} default {''} }
        [void]$rows.Append(@"
<tr>
<td>$(HtmlEnc ([string]$r.D_KENSIN))</td>
<td class="num">$(HtmlEnc (Normalize-Text ([string]$r.UKE_NO_KENSA)))</td>
<td><a href="/person?pk=$([int]$r.PK_SEQ)">$(HtmlEnc ([string]$r.KANJI_SIMEI))</a></td>
<td class="muted">$(HtmlEnc ([string]$r.KANA_SIMEI))</td>
<td>$sex</td>
<td>$(HtmlEnc ([string]$r.DANTAI))</td>
<td>$(HtmlEnc ([string]$r.COURSE_MEI))</td>
<td class="num">$done / $waku</td>
<td><span class="pill $cls">$lbl</span></td>
<td>$(if ($hantei -ne '') { '<span class="pill ok">済</span>' } else { '<span class="muted">-</span>' })</td>
</tr>
"@)
    }

    [void]$sb.Append("<p class='count'>$($dt.Rows.Count) 名 &nbsp;/&nbsp; 入力済 $nDone &nbsp; 途中 $nPart &nbsp; 未入力 $nNone &nbsp;/&nbsp; 自動判定済 $nHantei</p>")
    [void]$sb.Append('<div class="wrap"><table><thead><tr>')
    foreach ($h in @('受診日','受付番号','氏名','カナ','性別','事業所','コース','入力','進捗','自動判定')) {
        [void]$sb.Append("<th>$h</th>")
    }
    [void]$sb.Append('</tr></thead><tbody>')
    [void]$sb.Append($rows.ToString())
    [void]$sb.Append('</tbody></table></div>')
    [void]$sb.Append('<p class="note">「入力」は 結果が入っている項目数 / その人の検査枠の数 です。<br>氏名をクリックすると、その人の検査項目が見られます。<br>この画面は読むだけです。健診ナビのデータは変わりません。</p>')
    return $sb.ToString()
}

function Render-Person($conn, [int]$pk) {
    $h = Invoke-DbQuery $conn @'
SELECT TOP 1 s.PK_SEQ, s.D_KENSIN, s.UKE_NO_KENSA, j.KANJI_SIMEI, j.KANA_SIMEI, j.D_SEINEN,
       j.SEIBETU, j.KOJIN_NO, d.MEISYO1 AS DANTAI, c1.MEISYO AS COURSE_MEI, s.D_JIDOHANTEI
FROM T_KENSIN s
JOIN T_KOJIN1 j ON j.KOJIN_ID = s.KOJIN_ID
LEFT JOIN M_DANTAI d ON d.DANTAI_CD1 = s.DANTAI_CD1
LEFT JOIN T_COURSE1 c1 ON c1.COURSE_CD = s.COURSE_CD AND c1.DANTAI_CD1 = s.DANTAI_CD1
WHERE s.PK_SEQ = @p
'@ @{ p = $pk }
    if ($h.Rows.Count -eq 0) { return '<p class="err">その受診者が見つかりません。</p>' }
    $p0 = $h.Rows[0]

    $k = Invoke-DbQuery $conn @'
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, km.MEISYO1 AS 項目名, k.KEKKA, k.KEKKA_CD, k.HANTEI_KIGO
FROM T_KENSA k
LEFT JOIN T_KOMOKU km ON LTRIM(RTRIM(km.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.PK_SEQ = @p
ORDER BY k.KOMOKU_CD
'@ @{ p = $pk }

    $sex = switch (Normalize-Text ([string]$p0.SEIBETU)) { '1' {'男'} '2' {'女'} default {''} }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(@"
<div class="bar">
<b style="font-size:16px">$(HtmlEnc ([string]$p0.KANJI_SIMEI))</b>
<span class="muted">$(HtmlEnc ([string]$p0.KANA_SIMEI))</span>
<span>$sex</span>
<span>生年月日 $(HtmlEnc ([string]$p0.D_SEINEN))</span>
<span>受診日 <b>$(HtmlEnc ([string]$p0.D_KENSIN))</b></span>
<span>受付番号 <b>$(HtmlEnc (Normalize-Text ([string]$p0.UKE_NO_KENSA)))</b></span>
<span>社員番号 $(HtmlEnc (Normalize-Text ([string]$p0.KOJIN_NO)))</span>
<span style="margin-left:auto"><a href="javascript:history.back()">← 戻る</a></span>
</div>
<div class="bar">
<span>$(HtmlEnc ([string]$p0.DANTAI))</span><span>$(HtmlEnc ([string]$p0.COURSE_MEI))</span>
<span>自動判定 $(if ((Normalize-Text ([string]$p0.D_JIDOHANTEI)) -ne '') { '<span class="pill ok">済</span>' } else { '<span class="pill ng">まだ</span>' })</span>
</div>
"@)

    $filled = 0
    $rows = New-Object System.Text.StringBuilder
    foreach ($r in $k.Rows) {
        $v = Normalize-Text ([string]$r.KEKKA)
        if ($v -ne '') { $filled++ }
        $nm = Normalize-Text ([string]$r.項目名)
        [void]$rows.Append(@"
<tr>
<td class="muted">$(HtmlEnc ([string]$r.CD))</td>
<td>$(HtmlEnc $nm)</td>
<td>$(if ($v -eq '') { '<span class="muted">-</span>' } else { HtmlEnc $v })</td>
<td class="muted">$(HtmlEnc (Normalize-Text ([string]$r.KEKKA_CD)))</td>
<td>$(HtmlEnc (Normalize-Text ([string]$r.HANTEI_KIGO)))</td>
</tr>
"@)
    }
    [void]$sb.Append("<p class='count'>検査枠 $($k.Rows.Count) / 結果あり $filled</p>")
    [void]$sb.Append('<div class="wrap"><table><thead><tr><th>項目CD</th><th>項目名</th><th>結果</th><th>結果CD</th><th>判定</th></tr></thead><tbody>')
    [void]$sb.Append($rows.ToString())
    [void]$sb.Append('</tbody></table></div>')
    return $sb.ToString()
}

# ============================================================================
# サーバー
# ============================================================================

$cs = Resolve-ConnectionString
$envLabel = '接続先: ' + (($cs -split ';' | Where-Object { $_ -match '(?i)(data source|initial catalog)' }) -join ' / ')

$prefix = "http://127.0.0.1:$Port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
try { $listener.Start() }
catch {
    Write-Host "ポート $Port を使えませんでした。別のポートで試してください (例: -Port 8081)" -ForegroundColor Red
    throw
}

Write-Host ''
Write-Host '=== 健診ビューア ===' -ForegroundColor Cyan
Write-Host "  $prefix" -ForegroundColor Green
Write-Host "  $envLabel" -ForegroundColor DarkGray
Write-Host ''
Write-Host '  この画面は読むだけです。健診ナビのデータは変わりません。' -ForegroundColor DarkGray
Write-Host '  終わるときは、この黒い画面で Ctrl+C を押すか、ウィンドウを閉じてください。' -ForegroundColor Yellow
Write-Host ''

if (-not $NoBrowser) { Start-Process $prefix }

$nav = '<a href="/list" class="on">受診者一覧</a>'

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $req = $ctx.Request; $res = $ctx.Response
        $path = $req.Url.AbsolutePath
        $html = ''
        $conn = $null
        try {
            if ($path -eq '/favicon.ico') { $res.StatusCode = 404; $res.Close(); continue }
            $conn = Open-Db
            switch -Regex ($path) {
                '^/person$' {
                    $pk = 0
                    [void][int]::TryParse(($req.QueryString['pk']), [ref]$pk)
                    $html = Page '受診者' (Render-Person $conn $pk) $nav $envLabel
                    break
                }
                default {
                    $ymdIn = Normalize-Text ($req.QueryString['ymd'])
                    $ymd = ''
                    if ($ymdIn -ne '') {
                        $ymd = Normalize-Ymd $ymdIn
                        if (-not $ymd) { $ymd = $ymdIn }   # 解釈できなくてもそのまま検索させる
                    }
                    $kw = Normalize-Text ($req.QueryString['kw'])
                    $html = Page '受診者一覧' (Render-List $conn $ymd $kw) $nav $envLabel
                    break
                }
            }
        }
        catch {
            $msg = HtmlEnc ($_.Exception.Message)
            $html = Page 'エラー' "<div class='err'><b>エラーが起きました</b><br><pre>$msg</pre></div>" $nav $envLabel
        }
        finally { if ($conn) { $conn.Close() } }

        $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
        $res.ContentType = 'text/html; charset=utf-8'
        $res.ContentLength64 = $bytes.Length
        $res.OutputStream.Write($bytes, 0, $bytes.Length)
        $res.OutputStream.Close()
    }
}
finally {
    $listener.Stop()
    $listener.Close()
    Write-Host '終了しました。' -ForegroundColor DarkGray
}
