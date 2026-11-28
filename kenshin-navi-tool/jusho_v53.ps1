<#
.SYNOPSIS
  受診票 OCR 読み取りアプリ v5.3 の Excel (88 列) の「郵便番号」「住所」を、健診ナビの個人マスタ2 (T_KOJIN2) に入れる (jusho_v53.ps1)

.DESCRIPTION
  2026/11/28 check123 で確認した作り:
  ・T_KOJIN2 は 1 人 1 行 (主キー KOJIN_ID)。YUUBIN varchar(7) = ハイフンなしの 7 桁、
    JYUSYO1 nvarchar(10) = 都道府県だけ、JYUSYO2 nvarchar(100) = 市区町村・番地、JYUSYO3 nvarchar(30) = 建物名・部屋番号。
  ・新島 3 日間の 378 人のうち、住所が入っているのは 54 人だけ。
  入れ方:
  ・人の特定は 受診日 + 受付番号 (T_KENSIN)。受診日は -Ymd で指定する (聞かれる)。同じ受付番号が 2 日以上にあればその人は入れない。
  ・受診票の住所 (1 つながり) を、都道府県 / 市区町村・番地 / 建物名 に分ける。数字は半角、数字の間の横棒は「-」にそろえる。
  ・健診ナビの郵便番号の辞書 (T_JYUSYO: JS_POST / JS_JYUSYO1 = 都道府県 / JS_JYUSYO2 = 市区町村・町名。約 14.7 万行) を引く:
    都道府県が書かれていないときは辞書の都道府県を使う。「本村1-2-3」のように市区町村が無いときは、辞書の「新島村本村」と重なるところから「新島村」を補う。
    郵便番号が辞書に無い・都道府県が辞書と合わないときは、一覧の「注意」に ★ を付ける (読み違いに気づくため)。
    辞書で決まらないときは、同じ郵便番号で健診ナビに入っている人の都道府県、島の町村名の順に使う。それでも無ければ空のまま入れて「都道府県なし」と出す。
  ・T_KOJIN2 にトリガは無く、更新日時の欄も無い。予約 (T_KENSIN) に住所の写しは無い (2026/11/28 check123)。
  ・-Mode empty     = 空の人だけ入れる (健診ナビに住所が入っている人には触らない)
    -Mode overwrite = 受診票の住所で上書きする
    指定しなければ聞く。
  ・受診票の住所が空欄・「斜線」の人は何もしない (前の住所を消さない)。読めない字 (?・■) がある住所は入れない。
  ・送付先 (SOUFU_*) には触らない。
  ・プレビュー → y で書く。書く前に T_KOJIN2 の元の行を backup\ に保存し、変更を backup\jusho_v53_<日時>.csv に残す。-Undo で戻せる。
  ・全員分を 1 つのトランザクションで書く (途中で失敗したら全部取り消す)。
  例:
    jusho_v53.ps1 -File "...\1127_1127.xlsx" -Ymd 2026/11/23 -Mode empty
    jusho_v53.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$File,
    [string]$Ymd,                 # 受診日 (例 2026/11/24。複数はカンマ区切り)
    [string]$Mode,                # empty = 空の人だけ / overwrite = 上書き
    [switch]$PreviewOnly,         # 付けると、一覧を出すだけで終わる (書き込まない・聞かない)
    [switch]$Commit,
    [string]$Undo,
    [string]$ConnFile = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt',
    [string]$ConnectionString
)
$ErrorActionPreference = 'Stop'

$Core = Join-Path $PSScriptRoot 'form_import.ps1'
if (-not (Test-Path $Core)) { throw "form_import.ps1 が同じフォルダにありません: $PSScriptRoot" }
$src = Get-Content $Core -Raw
function Get-Part([string]$from, [string]$to) {
    $i = $src.IndexOf($from); $j = $src.IndexOf($to, $i)
    if ($i -lt 0 -or $j -lt 0) { throw "form_import.ps1 の構成が変わっています ($from)" }
    return $src.Substring($i, $j - $i)
}
Invoke-Expression (Get-Part 'function Normalize-Text'     '$XlsxLib = Join-Path')
. (Join-Path $PSScriptRoot 'xlsx_read.ps1')
Invoke-Expression (Get-Part 'function Test-XlsxEncrypted' 'function Detect-MappingPath')
Invoke-Expression (Get-Part 'function Get-LocalConnFile'  'function Resolve-PkSeq')
Invoke-Expression (Get-Part 'function Test-Locked'        'function Commit-Plan')
$BackupDir = Join-Path $PSScriptRoot 'backup'
if (-not (Test-Path $BackupDir)) { [void](New-Item -ItemType Directory -Path $BackupDir) }
$OutDir = Join-Path $PSScriptRoot 'out'
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir) }
function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
function Split-List([string]$s) { return @(($s -split '[,、\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }
function DbV($v) { if ($null -eq $v -or [string]$v -eq '') { return [DBNull]::Value } else { return $v } }
# 元の値を記録に残すとき、NULL と空文字を区別する
function RawV($v) { if ($null -eq $v -or $v -is [DBNull]) { return '<NULL>' } else { return [string]$v } }
function UnRaw($v) { if ([string]$v -eq '<NULL>') { return [DBNull]::Value } else { return [string]$v } }
function StrV($v) { if ($null -eq $v -or $v -is [DBNull]) { return '' } else { return ([string]$v).Trim() } }
# 記録の値を、SQL の ISNULL(列,'') と比べる形にする (NULL は空文字。前後の空白はそのまま)
function RawS($v) { if ($null -eq $v -or $v -is [DBNull] -or [string]$v -eq '<NULL>') { return '' } else { return [string]$v } }

# ---- 文字のそろえ方 ----
# 横棒に見える字 (‐ ‑ ‒ – — ― − － ー ｰ)。数字の間にあるものだけ「-」にする (カタカナの長音は変えない)
$HYCHARS = -join (@(0x2010, 0x2011, 0x2012, 0x2013, 0x2014, 0x2015, 0x2212, 0xFF0D, 0x30FC, 0xFF70) | ForEach-Object { [char]$_ })
$HYCLASS = '[\-' + $HYCHARS + ']'
function To-Half([string]$s) {
    if ($null -eq $s) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        $cp = [int]$ch
        if ($cp -ge 0xFF10 -and $cp -le 0xFF19) { [void]$sb.Append([char]($cp - 0xFEE0)) }   # 全角数字
        elseif ($cp -eq 0x3000) { [void]$sb.Append(' ') }                                       # 全角空白
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}
function Fix-Hyphen([string]$s) { return [regex]::Replace($s, '(?<=[0-9])\s*' + $HYCLASS + '+\s*(?=[0-9])', '-') }
function Addr-Key([string]$s) { return ((Fix-Hyphen (To-Half $s)) -replace '\s', '').ToUpper() }
function Clean-Zip([string]$s) { return ((To-Half (Normalize-KenNo $s)) -replace '[^0-9]', '') }

$PREFS = @('北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
           '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県',
           '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県',
           '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県')
# 都道府県を書かないことが多い、東京の島の町村
$TOKYO_ISLANDS = @('新島村', '神津島村', '大島町', '利島村', '三宅村', '御蔵島村', '八丈町', '青ヶ島村', '小笠原村')
$BLANKS = @('', '斜線', 'なし', '無し', '同上', '空欄', '未記入', '-')

# 受診票の住所 (1 つながり) を 都道府県 / 市区町村・番地 / 建物名 に分ける
function Split-Address([string]$raw) {
    $res = @{ J1 = ''; J2 = ''; J3 = ''; Notes = @(); Bad = ''; Zip = '' }
    $a = Fix-Hyphen (To-Half (Normalize-Text $raw))
    $a = ($a -replace '\s+', ' ').Trim()
    # 先頭に郵便番号が書いてあれば外す
    $mz = [regex]::Match($a, '^(?:〒\s*([0-9]{3})-?([0-9]{4})|([0-9]{3})-([0-9]{4}))(?![0-9\-])\s*')
    if ($mz.Success) {
        $res.Zip = $mz.Groups[1].Value + $mz.Groups[2].Value + $mz.Groups[3].Value + $mz.Groups[4].Value
        $a = $a.Substring($mz.Length).Trim()
    }
    if ($a -match '[?？■□●〓]' -or $a -match '読めな|判読|不明') { $res.Bad = '読めない字がある'; return $res }
    $rest = $a
    foreach ($pf in $PREFS) { if ($a.StartsWith($pf)) { $res.J1 = $pf; $rest = $a.Substring($pf.Length).Trim(); break } }
    # 番地のあとの空白で、建物名を分ける。「1丁目 2-3」のように番地の途中に空白があるときは続きとみなす
    $toks = @($rest -split ' ' | Where-Object { $_ -ne '' })
    $j2 = ''; $idx = 0; $seenDigit = $false
    while ($idx -lt $toks.Count) {
        $tk = $toks[$idx]
        if ($seenDigit -and $tk -notmatch '^[0-9\-]' -and $tk -notmatch '^(丁目|番|号|の[0-9])') { break }
        if ($j2 -match '[0-9]$' -and $tk -match '^[0-9]') { $j2 += '-' }
        $j2 += $tk
        if ($tk -match '[0-9]') { $seenDigit = $true }
        $idx++
    }
    $j3 = ''
    if ($idx -lt $toks.Count) { $j3 = ($toks[$idx..($toks.Count - 1)] -join ' ') }
    # 番地のすぐあとに空白なしで建物名が続くとき (「117-3シーグリーン」) も、建物名を分ける。
    # 番地らしい形 (横棒か 番・号 がある) のときだけ。「北12条西3丁目」のような町名の途中では分けない
    $mb = [regex]::Match($j2, '^(.*?)([0-9]+(?:(?:-|丁目|番地|番|号|の)[0-9]*)*)([^0-9\-].*)$')
    if ($mb.Success -and $mb.Groups[2].Value -match '-|番|号' -and $mb.Groups[3].Value.Length -ge 2 -and $mb.Groups[3].Value -notmatch '^(条|丁|番|号|線|地割|地先|先|以|の|区|町|村)') {
        $j2 = $mb.Groups[1].Value + $mb.Groups[2].Value
        $j3 = ($mb.Groups[3].Value + ' ' + $j3).Trim()
    }
    if ($j3.Length -gt 30) {
        if (($j2.Length + 1 + $j3.Length) -le 100) { $j2 = $j2 + ' ' + $j3; $j3 = ''; $res.Notes += '建物名が長いので住所2にまとめた' }
        else { $res.Bad = '建物名が長すぎる (30 字まで)'; return $res }
    }
    if ($j2.Length -gt 100) { $res.Bad = '住所が長すぎる (100 字まで)'; return $res }
    if ($j2.Length -lt 3) { $res.Bad = '住所が短すぎる'; return $res }
    $res.J2 = $j2; $res.J3 = $j3
    return $res
}

# ============================================================
# 元に戻す
# ============================================================
if ($Undo) {
    if ($Undo -eq 'ask') {
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'jusho_v53_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\jusho_v53_*.csv) がありません。'; return }
        for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
        $ans = Read-Host '戻す記録の番号 (それ以外 = やめる)'
        if ($ans -notmatch '^\d+$' -or [int]$ans -lt 1 -or [int]$ans -gt $files.Count) { Write-Host 'やめました。'; return }
        $Undo = $files[[int]$ans - 1].FullName
    }
    $urows = @(Import-Csv -Path $Undo -Encoding UTF8)
    Write-Host ("この記録で入れた住所 {0} 人分を、入れる前の内容に戻します: {1}" -f $urows.Count, $Undo) -ForegroundColor Yellow
    Write-Host '  入れたあとで人が書き換えた住所には触りません。'
    $ans = Read-Host '戻しますか (y = 戻す / それ以外 = やめる)'
    if (($ans -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。'; return }
    $conn = Open-Db
    try {
        $back = 0; $kept = 0
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran 'JUSHO_IMPORT'
            foreach ($r in $urows) {
                # 今の中身が「入れた内容」のままの行だけ戻す
                $n = Invoke-DbExec $conn $tran "UPDATE T_KOJIN2 SET YUUBIN = @py, JYUSYO1 = @p1, JYUSYO2 = @p2, JYUSYO3 = @p3 WHERE KOJIN_ID = @k AND ISNULL(YUUBIN,'') = @ny AND ISNULL(JYUSYO1,'') = @n1 AND ISNULL(JYUSYO2,'') = @n2 AND ISNULL(JYUSYO3,'') = @n3" `
                    @{ py = (UnRaw $r.前の郵便番号); p1 = (UnRaw $r.前の住所1); p2 = (UnRaw $r.前の住所2); p3 = (UnRaw $r.前の住所3); k = $r.KOJIN_ID
                       ny = (RawS $r.新しい郵便番号); n1 = (RawS $r.新しい住所1); n2 = (RawS $r.新しい住所2); n3 = (RawS $r.新しい住所3) }
                if ($n -eq 1) { $back++ } else { $kept++ }
            }
            $tran.Commit()
        } catch { $tran.Rollback(); throw }
        Write-Host ("[完了] 戻した人 {0} / 中身が変わっていたので触らなかった人 {1}" -f $back, $kept) -ForegroundColor Green
    } finally { $conn.Close() }
    return
}

# ============================================================
# 入れる
# ============================================================
if (-not $File) { $File = Read-Host '受診票 OCR v5.3 の Excel (住所の列つき) をドラッグ＆ドロップして Enter' }
$File = ($File -replace '^"|"$', '').Trim()
if (-not (Test-Path $File)) { throw "ファイルが見つかりません: $File" }
$rows = Read-FormCsv $File 'SJIS'
$head = $rows[0]
# 見出しの名前で列をさがす (列の順番が変わっても読める)
$colOf = @{}
for ($ci = 1; $ci -le $head.Count; $ci++) { $hn = Normalize-Text (Get-Field $head $ci); if ($hn -ne '' -and -not $colOf.ContainsKey($hn)) { $colOf[$hn] = $ci } }
if (-not $colOf.ContainsKey('受付NO')) { throw '受診票 OCR の Excel ではないようです (見出し「受付NO」がありません)。' }
$colAddr = 0; $colZip = 0
if ($colOf.ContainsKey('住所')) { $colAddr = $colOf['住所'] }
else { foreach ($hk in $colOf.Keys) { if ($hk -like '*住所*' -and $hk -notlike '*郵便*') { $colAddr = $colOf[$hk]; break } } }
if ($colOf.ContainsKey('郵便番号')) { $colZip = $colOf['郵便番号'] }
else { foreach ($hk in $colOf.Keys) { if ($hk -like '*郵便*') { $colZip = $colOf[$hk]; break } } }
if ($colAddr -le 0) { throw 'この Excel には「住所」の列がありません。住所は読み取りアプリ v5.3 で作った Excel (88 列) にだけ入っています。' }

if (-not $Ymd) { $Ymd = Read-Host '受診日 (例 2026/11/24。複数はカンマ区切り。このファイルの人が受けた日)' }
$dateList = @(Split-List $Ymd | ForEach-Object { Normalize-Ymd $_ } | Where-Object { $_ })
if ($dateList.Count -eq 0) { throw ("受診日「{0}」を解釈できません。2026/11/24 のように入力してください。" -f $Ymd) }

if (-not $Mode) {
    Write-Host ''
    Write-Host '健診ナビにもう住所が入っている人を、どうしますか。'
    Write-Host '  1 = 空の人だけ入れる (入っている人には触らない)'
    Write-Host '  2 = 受診票の住所で上書きする'
    $ans = Read-Host '番号 (1 か 2。それ以外 = やめる)'
    $ans = (To-Half ($ans -as [string])).Trim()
    if ($ans -eq '1') { $Mode = 'empty' } elseif ($ans -eq '2') { $Mode = 'overwrite' } else { Write-Host 'やめました。何も変えていません。'; return }
}
$Mode = $Mode.Trim().ToLower()
if ($Mode -ne 'empty' -and $Mode -ne 'overwrite') { throw "-Mode は empty (空の人だけ) か overwrite (上書き) のどちらかです。" }
$modeName = $(if ($Mode -eq 'empty') { '空の人だけ入れる' } else { '受診票の住所で上書きする' })

# ファイルにある受付番号と住所
$entries = @(); $ukeList = @(); $nBlank = 0
foreach ($f in $rows[1..($rows.Count - 1)]) {
    $u = Normalize-KenNo (Get-Field $f $colOf['受付NO'])
    if ($u -notmatch '^\d+$') { continue }
    $u = [string][int]$u
    $ukeList += [int]$u
    $rawAddr = Normalize-Text (Get-Field $f $colAddr)
    $rawZip = ''
    if ($colZip -gt 0) { $rawZip = Normalize-Text (Get-Field $f $colZip) }
    if ($BLANKS -contains ($rawAddr -replace '[\s　]', '')) { $nBlank++; continue }
    $entries += @{ Uke = $u; Addr = $rawAddr; Zip = $rawZip }
}
if ($ukeList.Count -eq 0) { throw '受付番号のある行がありません。' }
$dupInFile = @($ukeList | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($dupInFile.Count -gt 0) { throw ("ファイルの中に同じ受付番号が 2 行以上あります: {0}。直してからやり直してください。" -f ($dupInFile -join ', ')) }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 住所の取込 (受診票 OCR v5.3 → T_KOJIN2)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($File))
Write-Host ("  受診日: {0}  受付番号 {1} 人分 / 住所が書いてある人 {2} 人 / 住所が空欄の人 {3} 人" -f ($dateList -join ', '), $ukeList.Count, $entries.Count, $nBlank)
Write-Host ("  入れ方: {0}" -f $modeName) -ForegroundColor Yellow
if ($entries.Count -eq 0) { Write-Host '住所が書いてある人がいません。入れるものがありません。'; return }

$conn = Open-Db
try {
    $dLit = "'" + ($dateList -join "','") + "'"
    $people = @{}; $ambiguous = @{}
    $uLit = ($ukeList | Sort-Object -Unique) -join ','
    foreach ($r in (Invoke-DbQuery $conn "SELECT s.PK_SEQ, s.UKE_NO_KENSA AS UKE, s.KOJIN_ID, s.D_KENSIN AS YMD, LTRIM(RTRIM(ISNULL(g.KANJI_SIMEI,''))) AS NAME FROM T_KENSIN s LEFT JOIN T_KOJIN1 g ON g.KOJIN_ID = s.KOJIN_ID WHERE s.F_TORIKESI = 0 AND s.D_KENSIN IN ($dLit) AND s.UKE_NO_KENSA IN ($uLit)").Rows) {
        $uk = [string][int]$r.UKE
        # 同じ受付番号が 2 件 (別の日・同じ日) あれば、どの人か決められないので入れない
        if ($people.ContainsKey($uk) -or $ambiguous.ContainsKey($uk)) { $ambiguous[$uk] = 1; [void]$people.Remove($uk); continue }
        $people[$uk] = @{ Pk = [string]$r.PK_SEQ; Kojin = [string]$r.KOJIN_ID; Name = [string]$r.NAME; Ymd = [string]$r.YMD }
    }
    # 今の住所
    $cur = @{}
    if ($people.Count -gt 0) {
        $kLit = (@($people.Values | ForEach-Object { $_.Kojin }) | Sort-Object -Unique) -join ','
        foreach ($r in (Invoke-DbQuery $conn "SELECT KOJIN_ID, YUUBIN, JYUSYO1, JYUSYO2, JYUSYO3 FROM T_KOJIN2 WHERE KOJIN_ID IN ($kLit)").Rows) {
            $cur[[string]$r.KOJIN_ID] = @{ Y = (StrV $r.YUUBIN); A1 = (StrV $r.JYUSYO1); A2 = (StrV $r.JYUSYO2); A3 = (StrV $r.JYUSYO3)
                RawY = (RawV $r.YUUBIN); Raw1 = (RawV $r.JYUSYO1); Raw2 = (RawV $r.JYUSYO2); Raw3 = (RawV $r.JYUSYO3) }
        }
    }
    # 住所を分ける。郵便番号の辞書 (T_JYUSYO) と、同じ郵便番号の人の都道府県を引く
    $zipWant = @{}; $zipAll = @{}
    foreach ($en in $entries) {
        $en.Sp = Split-Address $en.Addr
        $zp = Clean-Zip $en.Zip
        if ($zp -eq '' -and $en.Sp.Zip -ne '') { $zp = $en.Sp.Zip }
        $en.ZipNew = $zp; $en.ZipOk = ($zp -match '^[0-9]{7}$')
        if ($en.ZipOk -and $en.Sp.Bad -eq '') { $zipAll[$zp] = 1 }
        if ($en.ZipOk -and $en.Sp.Bad -eq '' -and $en.Sp.J1 -eq '') { $zipWant[$zp] = 1 }
    }
    $zipDict = @{}; $dictOn = $false
    if ($zipAll.Count -gt 0) {
        try {
            $zaLit = "'" + ((@($zipAll.Keys) | Sort-Object) -join "','") + "'"
            foreach ($r in (Invoke-DbQuery $conn "SELECT JS_POST AS Z, LTRIM(RTRIM(ISNULL(JS_JYUSYO1,''))) AS P, LTRIM(RTRIM(ISNULL(JS_JYUSYO2,''))) AS C FROM T_JYUSYO WHERE JS_POST IN ($zaLit)").Rows) {
                $zk = ([string]$r.Z).Trim()
                if (-not $zipDict.ContainsKey($zk)) { $zipDict[$zk] = @{ Pref = ''; Cities = @() } }
                $dp = ([string]$r.P) -replace '[\s　]', ''
                if ($zipDict[$zk].Pref -eq '') { foreach ($pf in $PREFS) { if ($dp.StartsWith($pf)) { $zipDict[$zk].Pref = $pf; break } } }
                $dc = ([string]$r.C) -replace '[\s　]', ''
                if ($dc -ne '') { $zipDict[$zk].Cities += $dc }
            }
            $dictOn = $true
        } catch {
            Write-Host ("  (郵便番号の辞書 T_JYUSYO を引けませんでした。辞書なしで続けます: {0})" -f $_.Exception.Message) -ForegroundColor DarkGray
        }
    }
    $zipPref = @{}
    if ($zipWant.Count -gt 0) {
        $zLit = "'" + ((@($zipWant.Keys) | Sort-Object) -join "','") + "'"
        foreach ($r in (Invoke-DbQuery $conn "SELECT YUUBIN AS Z, LTRIM(RTRIM(JYUSYO1)) AS P, COUNT(*) AS CNT FROM T_KOJIN2 WHERE YUUBIN IN ($zLit) AND LTRIM(RTRIM(ISNULL(JYUSYO1,''))) <> '' GROUP BY YUUBIN, LTRIM(RTRIM(JYUSYO1)) ORDER BY 3 DESC").Rows) {
            $zk = ([string]$r.Z).Trim()
            if (-not $zipPref.ContainsKey($zk) -and ($PREFS -contains [string]$r.P)) { $zipPref[$zk] = [string]$r.P }
        }
    }

    # 計画
    $plan = @(); $notFound = @(); $ambList = @()
    foreach ($en in ($entries | Sort-Object { [int]$_.Uke })) {
        $uke = $en.Uke
        if ($ambiguous.ContainsKey($uke)) { $ambList += $uke; continue }
        if (-not $people.ContainsKey($uke)) { $notFound += $uke; continue }
        $p = $people[$uke]; $kj = $p.Kojin
        $row = New-Object PSObject -Property @{ 受付番号 = $uke; 受診日 = $p.Ymd; 氏名 = $p.Name; 操作 = ''; 状態 = ''; 郵便番号 = ''; 住所1 = ''; 住所2 = ''; 住所3 = ''
            前の郵便番号 = ''; 前の住所 = ''; 受診票の住所 = $en.Addr; 注意 = ''; KOJIN_ID = $kj; PkSeq = $p.Pk; Prev = $null; Fin = $null }
        if (-not $cur.ContainsKey($kj)) { $row.状態 = '個人マスタ2 に行が無い (入れません)'; $plan += $row; continue }
        $c0 = $cur[$kj]
        $row.前の郵便番号 = $c0.Y; $row.前の住所 = (($c0.A1 + ' ' + $c0.A2 + ' ' + $c0.A3) -replace '\s+', ' ').Trim()
        $sp = $en.Sp
        if ($sp.Bad -ne '') { $row.状態 = ('要確認: ' + $sp.Bad + ' (入れません)'); $plan += $row; continue }
        $notes = @($sp.Notes)
        $n1 = $sp.J1; $n2 = $sp.J2; $n3 = $sp.J3
        $dz = $null
        if ($en.ZipOk -and $zipDict.ContainsKey($en.ZipNew)) { $dz = $zipDict[$en.ZipNew] }
        if ($en.ZipOk -and $dictOn -and $null -eq $dz) { $notes += '★郵便番号が辞書に無い (読み違いかもしれない)' }
        if ($n1 -ne '' -and $null -ne $dz -and $dz.Pref -ne '' -and $dz.Pref -ne $n1) { $notes += ('★郵便番号の都道府県 (' + $dz.Pref + ') と合わない') }
        # 市区町村が書かれていない住所 (「本村1-2-3」) は、辞書の「新島村本村」と重なるところから「新島村」を補う
        if ($null -ne $dz -and (($n2 -split '[0-9]')[0]) -notmatch '[市区町村郡]') {
            $headPart = ''
            foreach ($cty in $dz.Cities) {
                for ($ov = [Math]::Min($cty.Length - 1, $n2.Length); $ov -ge 2; $ov--) {
                    if ($n2.StartsWith($cty.Substring($cty.Length - $ov))) { $headPart = $cty.Substring(0, $cty.Length - $ov); break }
                }
                if ($headPart -ne '') { break }
            }
            if ($headPart -match '[市区町村郡]$' -and ($headPart.Length + $n2.Length) -le 100) { $n2 = $headPart + $n2; $notes += '市区町村は郵便番号から補った' }
        }
        if ($n1 -eq '') {
            if ($null -ne $dz -and $dz.Pref -ne '') { $n1 = $dz.Pref; $notes += '都道府県は郵便番号から' }
            elseif ($en.ZipOk -and $zipPref.ContainsKey($en.ZipNew)) { $n1 = $zipPref[$en.ZipNew]; $notes += '都道府県は同じ郵便番号の人から' }
            else {
                foreach ($isl in $TOKYO_ISLANDS) { if ($n2.StartsWith($isl)) { $n1 = '東京都'; $notes += '都道府県は町村名から'; break } }
                if ($n1 -eq '') { $notes += '★都道府県なし' }
            }
        }
        # 番地の前に 市・区・町・村・郡 の字が無い住所 (「本村1-2-3」だけ、など) は知らせる
        if ((($n2 -split '[0-9]')[0]) -notmatch '[市区町村郡]') { $notes += '★市区町村が書かれていないかもしれない' }
        if ($en.ZipNew -ne '' -and -not $en.ZipOk) { $notes += ('★郵便番号が 7 桁でない (' + $en.ZipNew + ') → 郵便番号は入れない') }
        $curEmpty = (($c0.A1 + $c0.A2 + $c0.A3) -eq '')
        $same = ((Addr-Key ($n1 + $n2 + $n3)) -eq (Addr-Key ($c0.A1 + $c0.A2 + $c0.A3)))
        # 都道府県を補えなかったときは、都道府県を除いて比べる
        if (-not $same -and $sp.J1 -eq '' -and (Addr-Key ($n2 + $n3)) -eq (Addr-Key ($c0.A2 + $c0.A3))) { $same = $true }
        $wAddr = $false; $wZip = $false
        if ($curEmpty) { $wAddr = $true; $row.操作 = '新しく入れる' }
        elseif ($same) { $row.操作 = '同じ' }
        elseif ($Mode -eq 'overwrite') { $wAddr = $true; $row.操作 = '上書き' }
        else { $row.操作 = '触らない' }
        if ($en.ZipOk -and ($wAddr -or $same)) {
            if ($c0.Y -eq '') { $wZip = $true }
            elseif ($c0.Y -ne $en.ZipNew) {
                if ($Mode -eq 'overwrite') { $wZip = $true } else { $notes += ('郵便番号は前のまま (受診票は ' + $en.ZipNew + ')') }
            }
        }
        if ($wAddr -and -not $en.ZipOk -and $c0.Y -ne '') { $notes += '郵便番号は前のまま (受診票に無い)' }
        $row.注意 = ($notes -join ' / ')
        if (-not $wAddr -and -not $wZip) {
            $row.状態 = $(if ($same) { '同じ (変更なし)' } else { '入っています (触らない)' })
            $plan += $row; continue
        }
        if (-not $wAddr -and $wZip) { $row.操作 = '郵便番号だけ' }
        # 書いたあとの 4 つの欄 (変えない欄は今のまま)
        $fin = @{ Y = $c0.RawY; J1 = $c0.Raw1; J2 = $c0.Raw2; J3 = $c0.Raw3 }
        if ($wZip) { $fin.Y = $en.ZipNew }
        if ($wAddr) { $fin.J1 = $(if ($n1 -eq '') { '<NULL>' } else { $n1 }); $fin.J2 = $n2; $fin.J3 = $(if ($n3 -eq '') { '<NULL>' } else { $n3 }) }
        $row.Prev = $c0; $row.Fin = $fin
        $row.郵便番号 = (StrV (UnRaw $fin.Y)); $row.住所1 = (StrV (UnRaw $fin.J1)); $row.住所2 = (StrV (UnRaw $fin.J2)); $row.住所3 = (StrV (UnRaw $fin.J3))
        $row.状態 = 'OK'
        $plan += $row
    }

    $ok = @($plan | Where-Object { $_.状態 -eq 'OK' })
    $ng = @($plan | Where-Object { $_.状態 -ne 'OK' -and $_.状態 -ne '同じ (変更なし)' -and $_.状態 -ne '入っています (触らない)' })
    Write-Host ''
    Write-Host '=== 入れる住所 ===' -ForegroundColor Cyan
    if ($ok.Count -gt 0) {
        $ok | Select-Object 受付番号, 受診日, 氏名, 操作, 郵便番号, 住所1, 住所2, 住所3, 注意 | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
    } else { Write-Host '  (ありません)' }
    $upd = @($ok | Where-Object { $_.操作 -eq '上書き' })
    if ($upd.Count -gt 0) {
        Write-Host ("=== ★前から入っている住所を、受診票の住所で上書きするもの: {0} 人 ===" -f $upd.Count) -ForegroundColor Yellow
        $upd | Select-Object 受付番号, 氏名, 前の郵便番号, 前の住所, @{ n = '新しい住所'; e = { (($_.郵便番号 + ' ' + $_.住所1 + ' ' + $_.住所2 + ' ' + $_.住所3) -replace '\s+', ' ').Trim() } } | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
        Write-Host '   読み取りの読み違いで前の正しい住所を消さないよう、ここに出た人は紙と見比べてから y を押してください。' -ForegroundColor Yellow
    }
    if ($ng.Count -gt 0) {
        Write-Host ("=== ★入れない人 (要確認): {0} 人 → 紙を見て Excel の住所を直してから、もう一度流す ===" -f $ng.Count) -ForegroundColor Yellow
        $ng | Select-Object 受付番号, 氏名, 状態, 受診票の住所 | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
    }
    $warn = @($ok | Where-Object { $_.注意 -like '*★*' })
    if ($warn.Count -gt 0) { Write-Host ("  ★注意の付いた人 (入れますが、一覧の「注意」を見てください): 受付番号 {0}" -f (($warn | ForEach-Object { $_.受付番号 }) -join ', ')) -ForegroundColor Yellow }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $previewFile = Join-Path $OutDir ("住所取込_{0}_確認.csv" -f $stamp)
    $plan | Select-Object 受付番号, 受診日, 氏名, 操作, 状態, 郵便番号, 住所1, 住所2, 住所3, 前の郵便番号, 前の住所, 受診票の住所, 注意 | Export-Csv -Path $previewFile -NoTypeInformation -Encoding UTF8

    Line
    Write-Host ("書く: {0} 人 (新しく入れる {1}・上書き {2}・郵便番号だけ {3}) / 同じなので何もしない: {4} 人 / 入っているので触らない: {5} 人 / 要確認で入れない: {6} 人 / 住所が空欄: {7} 人" -f $ok.Count,
        @($ok | Where-Object { $_.操作 -eq '新しく入れる' }).Count, @($ok | Where-Object { $_.操作 -eq '上書き' }).Count, @($ok | Where-Object { $_.操作 -eq '郵便番号だけ' }).Count,
        @($plan | Where-Object { $_.状態 -eq '同じ (変更なし)' }).Count, @($plan | Where-Object { $_.状態 -eq '入っています (触らない)' }).Count, $ng.Count, $nBlank) -ForegroundColor Green
    Write-Host ("  健診ナビに見つからない受付番号: {0}" -f $(if ($notFound.Count -gt 0) { $notFound -join ', ' } else { 'なし' })) -ForegroundColor Green
    if ($ambList.Count -gt 0) { Write-Host ("  ★同じ受付番号の予約が 2 件以上あって決められない番号 (入れません): {0}  → 受診日を 1 日だけ指定してやり直してください。" -f ($ambList -join ', ')) -ForegroundColor Red }
    if ($notFound.Count -gt 0) { Write-Host '  見つからない番号は、受診日が違うか、健診ナビに受付番号がまだ入っていません。' -ForegroundColor Yellow }
    Write-Host ("  一覧を Excel で見るとき: {0}  (氏名と住所が入っています。外に出さないでください)" -f $previewFile)
    Line
    if ($ok.Count -eq 0) { Write-Host '入れるものがありません。'; return }
    if ($PreviewOnly) { Write-Host '※ 一覧を出しただけです。まだ何も書き込んでいません。' -ForegroundColor Yellow; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
        $ans = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($ans -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }

    # 健診ナビの画面で開かれている人は飛ばす
    $targets = @(); $skipped = @()
    foreach ($t in $ok) {
        $lock = Test-Locked $conn $t.PkSeq
        if ($lock) { Write-Host ("  [とばしました] {0} ({1}) は健診ナビで編集中です ({2})" -f $t.氏名, $t.受付番号, $lock) -ForegroundColor Yellow; $skipped += $t.受付番号; continue }
        $targets += $t
    }
    if ($targets.Count -eq 0) { Write-Host '書ける人がいません (全員が健診ナビで開かれています)。'; return }

    # 書く前の控え (元の行そのもの) と、元に戻すための記録
    $tLit = (@($targets | ForEach-Object { $_.KOJIN_ID }) | Sort-Object -Unique) -join ','
    $bkFile = Join-Path $BackupDir ("T_KOJIN2_{0}.csv" -f $stamp)
    $bk = Invoke-DbQuery $conn "SELECT * FROM T_KOJIN2 WHERE KOJIN_ID IN ($tLit)"
    $bkCols = @($bk.Columns | ForEach-Object { $_.ColumnName })
    $bk.Rows | Select-Object $bkCols | Export-Csv -Path $bkFile -NoTypeInformation -Encoding UTF8
    $undoFile = Join-Path $BackupDir ("jusho_v53_{0}.csv" -f $stamp)
    $targets | ForEach-Object {
        New-Object PSObject -Property @{ KOJIN_ID = $_.KOJIN_ID; 受付番号 = $_.受付番号; 受診日 = $_.受診日; 氏名 = $_.氏名; 操作 = $_.操作
            前の郵便番号 = $_.Prev.RawY; 前の住所1 = $_.Prev.Raw1; 前の住所2 = $_.Prev.Raw2; 前の住所3 = $_.Prev.Raw3
            新しい郵便番号 = $_.Fin.Y; 新しい住所1 = $_.Fin.J1; 新しい住所2 = $_.Fin.J2; 新しい住所3 = $_.Fin.J3 }
    } | Select-Object KOJIN_ID, 受付番号, 受診日, 氏名, 操作, 前の郵便番号, 前の住所1, 前の住所2, 前の住所3, 新しい郵便番号, 新しい住所1, 新しい住所2, 新しい住所3 |
        Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8

    $done = 0
    $tran = $conn.BeginTransaction()
    try {
        Acquire-AppLock $conn $tran 'JUSHO_IMPORT'
        foreach ($t in $targets) {
            # 一覧を出したときの中身のままの行だけ書く (その間に人が書き換えていたら止める)
            $n = Invoke-DbExec $conn $tran "UPDATE T_KOJIN2 SET YUUBIN = @y, JYUSYO1 = @j1, JYUSYO2 = @j2, JYUSYO3 = @j3 WHERE KOJIN_ID = @k AND ISNULL(YUUBIN,'') = @oy AND ISNULL(JYUSYO1,'') = @o1 AND ISNULL(JYUSYO2,'') = @o2 AND ISNULL(JYUSYO3,'') = @o3" `
                @{ y = (UnRaw $t.Fin.Y); j1 = (UnRaw $t.Fin.J1); j2 = (UnRaw $t.Fin.J2); j3 = (UnRaw $t.Fin.J3); k = $t.KOJIN_ID
                   oy = (RawS $t.Prev.RawY); o1 = (RawS $t.Prev.Raw1); o2 = (RawS $t.Prev.Raw2); o3 = (RawS $t.Prev.Raw3) }
            if ($n -ne 1) { throw ("書込の影響行数が {0} でした (受付番号 {1})。一覧を出したあとで住所が書き換えられたようです。全部取り消します。" -f $n, $t.受付番号) }
            $done++
        }
        $tran.Commit()
    } catch {
        $tran.Rollback()
        if (Test-Path $undoFile) { Remove-Item $undoFile -Force }
        Write-Host '★書き込みを全部取り消しました。何も変わっていません。' -ForegroundColor Red
        throw
    }
    Write-Host ''
    Write-Host ("[完了] {0} 人の住所を書きました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0}  → 画面を閉じてから、もう一度流してください" -f ($skipped -join ', ')) -ForegroundColor Red }
    Write-Host '   元に戻すときは 受診票住所_元に戻す.bat。' -ForegroundColor Yellow
}
finally { $conn.Close() }
