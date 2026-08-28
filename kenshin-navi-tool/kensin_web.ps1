<#
.SYNOPSIS
  健診ナビのDBを見る自作画面 (kensin_web.ps1)

.DESCRIPTION
  健診ナビと同じDBを読んで、ブラウザで見られる画面を出す。
  インストールは要らない。Windowsに元から入っている機能だけで動く。

  一覧・個人の閲覧に加えて、個人画面から結果を入力して保存できる。
  保存は変更した欄だけを UPDATE し、書込前に backup フォルダへ控えを取る。
  判定・総合所見は計算しない (健診ナビの「自動判定」で行う)。

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
Invoke-Expression (Get-Part 'function Get-LocalConnFile' 'function Get-CurrentKensa')
Invoke-Expression (Get-Part 'function Backup-Kensa' 'function Backup-Kojin1')
Invoke-Expression (Get-Part 'function Acquire-AppLock' 'function Commit-Plan')

# バックアップの置き場 (form_import.ps1 と同じ)
$BackupDir = Join-Path $PSScriptRoot 'backup'

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
td input[type=text], td select { padding:3px 5px; border:1px solid #bcc6d0; border-radius:3px; font-size:13px; font-family:inherit; }
td input[type=text]:focus, td select:focus { outline:2px solid #2d6ca2; border-color:#2d6ca2; }
.saved { background:#e2efda; border:1px solid #a9c48c; color:#375623; padding:10px 12px; border-radius:6px; margin-bottom:12px; }
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

# DBの版によって列の有無が違うので、あるかどうかを一度だけ調べて覚えておく
$script:ColCache = @{}
function Test-DbColumn($conn, [string]$table, [string]$col) {
    $key = "$table.$col"
    if ($script:ColCache.ContainsKey($key)) { return $script:ColCache[$key] }
    $dt = Invoke-DbQuery $conn @'
SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = @t AND COLUMN_NAME = @c
'@ @{ t = $table; c = $col }
    $has = ([int]$dt.Rows[0].N -gt 0)
    $script:ColCache[$key] = $has
    return $has
}

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
    [void]$sb.Append('<p class="note">「入力」は 結果が入っている項目数 / その人の検査枠の数 です。<br>氏名をクリックすると、その人の結果を入力できます。</p>')
    return $sb.ToString()
}

function Render-Person($conn, [int]$pk, [string]$msg) {
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

    # 単位・所見コードの列はDBの版で有無が違うので、あるものだけ読む
    $selTani = if (Test-DbColumn $conn 'T_KOMOKU' 'TANI') { 'km.TANI' } else { "''" }
    $selSho  = if (Test-DbColumn $conn 'T_KOMOKU' 'SYOKEN_CD') { "LTRIM(RTRIM(ISNULL(km.SYOKEN_CD,'')))" } else { "''" }
    $k = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(k.KOMOKU_CD)) AS CD, km.MEISYO1 AS 項目名, $selTani AS 単位,
       $selSho AS SYOKEN_CD,
       k.KEKKA, k.KEKKA_CD, k.HANTEI_KIGO
FROM T_KENSA k
LEFT JOIN T_KOMOKU km ON LTRIM(RTRIM(km.KOMOKU_CD)) = LTRIM(RTRIM(k.KOMOKU_CD))
WHERE k.PK_SEQ = @p
ORDER BY k.KOMOKU_CD
"@ @{ p = $pk }

    # この人が使う選択肢だけをまとめて引く (所見コード表 T_SYOKEN2)
    $choices = @{}
    $shoCds = @($k.Rows | ForEach-Object { Normalize-Text ([string]$_.SYOKEN_CD) } |
                Where-Object { $_ -ne '' } | Sort-Object -Unique)
    if ($shoCds.Count -gt 0) {
        $inList = ($shoCds | ForEach-Object { "'" + $_.Replace("'", "''") + "'" }) -join ','
        $sh = Invoke-DbQuery $conn @"
SELECT LTRIM(RTRIM(SYOKEN_CD)) AS SC, SYOKEN, KEKKA_CD
FROM T_SYOKEN2 WHERE LTRIM(RTRIM(SYOKEN_CD)) IN ($inList)
ORDER BY SYOKEN_CD, KEKKA_CD
"@ $null
        foreach ($r in $sh.Rows) {
            $sc = [string]$r.SC
            $txt = Normalize-Text ([string]$r.SYOKEN)
            if ($txt -eq '') { continue }
            if (-not $choices.ContainsKey($sc)) { $choices[$sc] = @() }
            $choices[$sc] += @{ SYOKEN = $txt; KEKKA_CD = Normalize-Text ([string]$r.KEKKA_CD) }
        }
    }

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

    if ($msg -ne '') { [void]$sb.Append("<div class='saved'>$(HtmlEnc $msg)</div>") }

    # ---- 入力欄。選択肢のある項目はプルダウンにする ----
    $filled = 0
    $rows = New-Object System.Text.StringBuilder
    foreach ($r in $k.Rows) {
        $cd = [string]$r.CD
        $v = Normalize-Text ([string]$r.KEKKA)
        if ($v -ne '') { $filled++ }
        $nm = Normalize-Text ([string]$r.項目名)
        $tani = Normalize-Text ([string]$r.単位)
        $sid = HtmlEnc $cd

        $sho = Normalize-Text ([string]$r.SYOKEN_CD)
        if ($sho -ne '' -and $choices.ContainsKey($sho)) {
            # 所見・選択式の項目 → プルダウン
            $opt = New-Object System.Text.StringBuilder
            [void]$opt.Append('<option value=""></option>')
            $hit = $false
            foreach ($c in $choices[$sho]) {
                $sel = if ($c.SYOKEN -eq $v) { $hit = $true; ' selected' } else { '' }
                [void]$opt.Append("<option value=""$(HtmlEnc $c.SYOKEN)""$sel>$(HtmlEnc $c.SYOKEN)</option>")
            }
            # マスタに無い値が既に入っている場合も消さずに残す
            if (-not $hit -and $v -ne '') {
                [void]$opt.Append("<option value=""$(HtmlEnc $v)"" selected>$(HtmlEnc $v)</option>")
            }
            $cell = "<select name=""v_$sid"">$($opt.ToString())</select>"
        }
        else {
            $cell = "<input type=""text"" name=""v_$sid"" value=""$(HtmlEnc $v)"" size=""14"">"
        }

        [void]$rows.Append(@"
<tr>
<td class="muted">$(HtmlEnc $cd)</td>
<td>$(HtmlEnc $nm)</td>
<td>$cell</td>
<td class="muted">$(HtmlEnc $tani)</td>
<td class="muted">$(HtmlEnc (Normalize-Text ([string]$r.KEKKA_CD)))</td>
<td>$(HtmlEnc (Normalize-Text ([string]$r.HANTEI_KIGO)))</td>
</tr>
"@)
    }
    [void]$sb.Append("<p class='count'>検査枠 $($k.Rows.Count) / 結果あり $filled</p>")
    [void]$sb.Append("<form method='post' action='/save'><input type='hidden' name='pk' value='$pk'>")
    [void]$sb.Append('<div class="wrap"><table><thead><tr><th>項目CD</th><th>項目名</th><th>結果</th><th>単位</th><th>結果CD</th><th>判定</th></tr></thead><tbody>')
    [void]$sb.Append($rows.ToString())
    [void]$sb.Append('</tbody></table></div>')
    [void]$sb.Append(@"
<div class="bar" style="margin-top:12px">
  <button type="submit">保存する</button>
  <span class="muted">変更した欄だけ書き込みます。書き込む前に backup フォルダへ自動で控えを取ります。</span>
</div></form>
<p class="note">保存しても<b>自動判定は動きません</b>。判定・総合所見は健診ナビ側で「自動判定」を実行してください。<br>
元に戻したいときは <b>元に戻す.bat</b> で、backup フォルダの控えから戻せます。</p>
"@)
    return $sb.ToString()
}

# 入力された結果を T_KENSA に書き戻す。
# 変更のあった項目だけ UPDATE する (触っていない欄は上書きしない)。
function Save-Person($conn, [int]$pk, [hashtable]$posted) {
    $cur = Invoke-DbQuery $conn @'
SELECT LTRIM(RTRIM(KOMOKU_CD)) AS CD, KEKKA FROM T_KENSA WHERE PK_SEQ = @p
'@ @{ p = $pk }
    if ($cur.Rows.Count -eq 0) { throw "その受診者の検査枠がありません (PK_SEQ=$pk)" }

    $changes = @()
    foreach ($r in $cur.Rows) {
        $cd = [string]$r.CD
        if (-not $posted.ContainsKey($cd)) { continue }      # 画面に無かった項目は触らない
        $new = Normalize-Text $posted[$cd]
        $old = Normalize-Text ([string]$r.KEKKA)
        if ($new -eq $old) { continue }
        $changes += @{ CD = $cd; Old = $old; New = $new }
    }
    if ($changes.Count -eq 0) { return '変更はありませんでした。' }

    [void](Backup-Kensa $conn $pk)

    $tran = $conn.BeginTransaction()
    try {
        [void](Acquire-AppLock $conn $tran "T_KENSA:$pk")
        foreach ($c in $changes) {
            $n = Invoke-DbExec $conn $tran `
                'UPDATE T_KENSA SET KEKKA = @k WHERE PK_SEQ = @p AND KOMOKU_CD = @cd' `
                @{ k = $c.New; p = $pk; cd = $c.CD }
            if ($n -ne 1) { throw "項目 $($c.CD) の更新で $n 行が対象になりました。中止します。" }
        }
        $tran.Commit()
    }
    catch { $tran.Rollback(); throw }

    Write-Host ("[保存] PK_SEQ={0} {1} 項目" -f $pk, $changes.Count) -ForegroundColor Green
    foreach ($c in $changes) {
        Write-Host ("    {0}: 「{1}」→「{2}」" -f $c.CD, $c.Old, $c.New) -ForegroundColor DarkGray
    }
    return ("{0} 項目を保存しました。" -f $changes.Count)
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
Write-Host '  個人画面で結果を入力して保存できます。書込前に backup へ控えを取ります。' -ForegroundColor DarkGray
Write-Host '  終わるときは、この黒い画面で Ctrl+C を押すか、ウィンドウを閉じてください。' -ForegroundColor Yellow
Write-Host ''

if (-not $NoBrowser) { Start-Process $prefix }

$nav = '<a href="/list" class="on">受診者一覧</a>'

try {
    while ($listener.IsListening) {
        try { $ctx = $listener.GetContext() }
        catch { Write-Host "[受付できませんでした] $($_.Exception.Message)" -ForegroundColor DarkYellow; continue }
        $req = $ctx.Request; $res = $ctx.Response
        $path = $req.Url.AbsolutePath
        $html = ''
        $conn = $null
        $sent = $false      # レスポンスを自前で返した (リダイレクト等) 場合に立てる
        try {
            if ($path -eq '/favicon.ico') { $res.StatusCode = 404; $res.Close(); continue }
            $conn = Open-Db
            switch -Regex ($path) {
                '^/person$' {
                    $pk = 0
                    [void][int]::TryParse(($req.QueryString['pk']), [ref]$pk)
                    $msg = Normalize-Text ($req.QueryString['msg'])
                    $html = Page '受診者' (Render-Person $conn $pk $msg) $nav $envLabel
                    break
                }
                '^/save$' {
                    if ($req.HttpMethod -ne 'POST') { $res.StatusCode = 405; $sent = $true; break }
                    $rdr = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
                    $body = $rdr.ReadToEnd(); $rdr.Close()
                    $posted = @{}
                    $pk = 0
                    foreach ($pair in ($body -split '&')) {
                        if ($pair -eq '') { continue }
                        $eq = $pair.IndexOf('=')
                        if ($eq -lt 0) { continue }
                        $nameRaw = [System.Uri]::UnescapeDataString(($pair.Substring(0, $eq)).Replace('+', ' '))
                        $valRaw  = [System.Uri]::UnescapeDataString(($pair.Substring($eq + 1)).Replace('+', ' '))
                        if ($nameRaw -eq 'pk') { [void][int]::TryParse($valRaw, [ref]$pk); continue }
                        if ($nameRaw -like 'v_*') { $posted[$nameRaw.Substring(2)] = $valRaw }
                    }
                    $done = Save-Person $conn $pk $posted
                    # 二重送信を防ぐため、保存後は GET に戻す
                    $res.StatusCode = 303
                    $res.RedirectLocation = "/person?pk=$pk&msg=" + [System.Uri]::EscapeDataString($done)
                    $sent = $true
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
            # 画面にも黒い画面にも出す。何があってもサーバーは止めない。
            $emsg = $_.Exception.Message
            $where = $_.InvocationInfo.PositionMessage
            Write-Host "[エラー] $path : $emsg" -ForegroundColor Red
            if ($where) { Write-Host $where -ForegroundColor DarkGray }
            $html = Page 'エラー' ("<div class='err'><b>エラーが起きました</b><br><pre>{0}</pre><pre class='muted'>{1}</pre></div>" -f (HtmlEnc $emsg), (HtmlEnc $where)) $nav $envLabel
            $sent = $false
        }
        finally { if ($conn) { $conn.Close() } }

        try {
            if ($sent) { $res.Close() }
            else {
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
                $res.ContentType = 'text/html; charset=utf-8'
                $res.ContentLength64 = $bytes.Length
                $res.OutputStream.Write($bytes, 0, $bytes.Length)
                $res.OutputStream.Close()
            }
        }
        catch {
            # ブラウザ側が先に切った場合など。ここで止まる理由はない。
            Write-Host "[送信できませんでした] $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
    Write-Host '終了しました。' -ForegroundColor DarkGray
}
