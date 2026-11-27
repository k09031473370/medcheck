<#
.SYNOPSIS
  受診票 OCR 読み取りツール v5 の Excel (80 列) の既往歴と家族歴を、健診ナビの既往歴 (T_KIOU) に入れる (kiou_v5.ps1)

.DESCRIPTION
  v5 の Excel の「既往歴1〜5_病名 (K01〜K64)・発症年齢・転帰 (T01〜T10)」を本人の既往歴として入れる。
  K の番号は健診ナビの病名コード (19 = 高血圧、47 = 高脂血症、48 = 糖尿病)、T の番号は転帰コード (05 = 薬剤治療中) と同じ。
  form\kiou_codes.csv で健診ナビの名前に直して入れる (受診票の「脂質異常症」は健診ナビの名前「高脂血症」になる。自動判定の対応表がこの名前で引くため)。
  入れ方は石屋の kiou123.ps1 と同じ:
  ・本人の既往歴として入れる。最初の 1 件は「特記事項なし」の行を置き換え、2 件目からは行を足す。
  ・同じ病名が既にあれば入れない。本人の既往歴は合計 -Max 件 (既定 4。結果票の枠が 4 つ) まで。
  ・病名の入った行は消さない・書き換えない。
  ・プレビュー → y で書く。書く前に本人の T_KIOU を backup\ に保存し、変更を backup\kiou_v5_<日時>.csv に残す。-Undo で戻せる。
  ・入れたあと自動判定をかけると J (治療中) などが付く。
  人の特定は 受診日 + 受付番号 (T_KENSIN)。ファイルに受診日と氏名の列が無いので、受診日は -Ymd で指定する (聞かれる)。
  ★1 つのファイルには同じ受診日の人だけを入れること。複数の日を指定したとき、同じ受付番号が 2 日以上にあればその人は入れない。
  家族歴 (家族歴1〜5_病名 B01〜B17・続柄 Z-1〜Z-4) も同じ表に「続柄つきの行」として入れる (2026/11/26 check122 で確認):
  ・続柄 Z-1 = 1 父 / Z-2 = 2 母 / Z-3 = 3 兄弟姉妹 / Z-4 = 4 祖父母。病名 B の番号は健診ナビの家族の病名コード (5 = 高血圧、13 = 糖尿病) と同じ。
  ・続柄が 2 つ書かれた病名 (Z-1,Z-2) は、続柄ごとに 1 行ずつ入れる。同じ病名・同じ続柄が既にあれば入れない。
  ・既往歴・家族歴は人ごとに 1 つ (受診のたびに別ではない) なので、前の年に入れた内容が残っている人がいる。
    同じ病名が前から入っていて、年齢・治療・続柄が受診票と違うときは、**受診票に合わせてその行を書き換える** (操作「更新」。2026/11/27 業務判断: 確認済みの Excel を正とする)。
    書き換える行はプレビューに「14→5」「父→祖父母」の形で別にまとめて出す。受診票に書かれていない病名の行は消さない・触らない。
    -KeepOld を付けると、前からある行は書き換えず、違いを知らせるだけにする。
  ・家族の「なし」の行 (続柄 未設定) があれば、最初の 1 件でその行を置き換える。年齢・治療の列は、健診ナビの家族の行でいちばん多い入り方に合わせる。
  自覚症状は start.bat (form_import) の方で入る。
  例:
    kiou_v5.ps1 -File "...\健診結果_2人_1126.xlsx" -Ymd 2026/11/24
    kiou_v5.ps1 -Undo ask
#>
[CmdletBinding()]
param(
    [string]$File,
    [string]$Ymd,                 # 受診日 (例 2026/11/24。複数はカンマ区切り)
    [int]$Max = 4,
    [switch]$KeepOld,             # 付けると、前から健診ナビにある行は書き換えない (受診票との違いを知らせるだけ)
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
function Line { Write-Host ('-' * 76) -ForegroundColor DarkGray }
function Split-List([string]$s) { return @(($s -split '[,、\s]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }
$PLACEHOLDER = @('', '特記事項なし', '特記事項無し', '特になし', '特に無し', 'なし', '無し')
function Test-Placeholder([string]$s) { return ($PLACEHOLDER -contains ((Normalize-Text $s) -replace '[\s　]', '')) }
function DbV($v) { if ($null -eq $v -or [string]$v -eq '') { return [DBNull]::Value } else { return $v } }

# ---- コード表 ----
$codeFile = Join-Path $PSScriptRoot 'form\kiou_codes.csv'
if (-not (Test-Path $codeFile)) { throw "form\kiou_codes.csv がありません。" }
$BY = @{}; $CH = @{}
foreach ($l in (Get-Content $codeFile -Encoding UTF8)) {
    if ($l -match '^\s*#' -or $l -match '^Kind,') { continue }
    $f = $l -split ','
    if ($f.Count -lt 3) { continue }
    if ($f[0] -eq 'BYOMEI') { $BY[[int]$f[1]] = $f[2] } elseif ($f[0] -eq 'CHIRYO') { $CH[[int]$f[1]] = $f[2] }
}

# ============================================================
# 元に戻す
# ============================================================
if ($Undo) {
    if ($Undo -eq 'ask') {
        $files = @(Get-ChildItem -Path $BackupDir -Filter 'kiou_v5_*.csv' -File | Sort-Object LastWriteTime -Descending)
        if ($files.Count -eq 0) { Write-Host '戻せる記録 (backup\kiou_v5_*.csv) がありません。'; return }
        for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  {0}: {1}  ({2})" -f ($i + 1), $files[$i].Name, $files[$i].LastWriteTime) }
        $a = Read-Host '戻す記録の番号 (それ以外 = やめる)'
        if ($a -notmatch '^\d+$' -or [int]$a -lt 1 -or [int]$a -gt $files.Count) { Write-Host 'やめました。'; return }
        $Undo = $files[[int]$a - 1].FullName
    }
    $rows = @(Import-Csv -Path $Undo -Encoding UTF8)
    Write-Host ("この記録で入れた既往歴・家族歴 {0} 行を元に戻します (追加した行は消し、置き換えた行は元の内容に戻す): {1}" -f $rows.Count, $Undo) -ForegroundColor Yellow
    $a = Read-Host '戻しますか (y = 戻す / それ以外 = やめる)'
    if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。'; return }
    $conn = Open-Db
    try {
        $back = 0; $kept = 0
        foreach ($r in $rows) {
            $tran = $conn.BeginTransaction()
            try {
                Acquire-AppLock $conn $tran ("KIOU_IMPORT:" + $r.KOJIN_ID)
                if ($r.種類 -eq '家族' -and $r.操作 -eq '追加') {
                    $n = Invoke-DbExec $conn $tran "DELETE FROM T_KIOU WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @b AND LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) = @z" @{ k = $r.KOJIN_ID; r = $r.RENBAN; b = $r.病名; z = $r.続柄CD }
                } elseif ($r.種類 -eq '家族') {
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KIOU SET BYOMEI = @ob, BYOMEI_CD = @oc, ZOKUGARA = @oz, ZOKUGARAMEI = @ozm WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @b AND LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) = @z" `
                        @{ ob = $r.前の病名; oc = $r.前の病名CD; oz = $r.前の続柄CD; ozm = $r.前の続柄名; k = $r.KOJIN_ID; r = $r.RENBAN; b = $r.病名; z = $r.続柄CD }
                } elseif ($r.操作 -eq '追加') {
                    $n = Invoke-DbExec $conn $tran "DELETE FROM T_KIOU WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @b" @{ k = $r.KOJIN_ID; r = $r.RENBAN; b = $r.病名 }
                } else {
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KIOU SET BYOMEI = @ob, BYOMEI_CD = @oc, CHIRYO = @och, CHIRYOMEI = @ochm, AGE = @oa WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @b" `
                        @{ ob = $r.前の病名; oc = $r.前の病名CD; och = $r.前の治療; ochm = $r.前の治療名; oa = (DbV $r.前の年齢); k = $r.KOJIN_ID; r = $r.RENBAN; b = $r.病名 }
                }
                $tran.Commit()
                if ($n -eq 1) { $back++ } else { $kept++ }
            } catch { $tran.Rollback(); throw }
        }
        Write-Host ("[完了] 戻した行 {0} / 中身が変わっていたので触らなかった行 {1}" -f $back, $kept) -ForegroundColor Green
    } finally { $conn.Close() }
    return
}

# ============================================================
# 入れる
# ============================================================
if (-not $File) { $File = Read-Host '受診票 OCR v5 の Excel (80 列) をドラッグ＆ドロップして Enter' }
$File = ($File -replace '^"|"$', '').Trim()
if (-not (Test-Path $File)) { throw "ファイルが見つかりません: $File" }
$rows = Read-FormCsv $File 'SJIS'
$head = $rows[0]
# 見出しの名前で列をさがす (列の順番が変わっても読める)
$colOf = @{}
for ($ci = 1; $ci -le $head.Count; $ci++) { $hn = Normalize-Text (Get-Field $head $ci); if ($hn -ne '' -and -not $colOf.ContainsKey($hn)) { $colOf[$hn] = $ci } }
foreach ($need in @('受付NO', '既往歴1_病名', '既往歴1_発症年齢', '既往歴1_転帰')) {
    if (-not $colOf.ContainsKey($need)) { throw ("受診票 OCR v5 の Excel ではないようです (見出し「{0}」がありません)。" -f $need) }
}
$kiouCols = @()
for ($ki = 1; $ki -le 9; $ki++) { if ($colOf.ContainsKey("既往歴${ki}_病名")) { $kiouCols += @{ B = $colOf["既往歴${ki}_病名"]; A = $colOf["既往歴${ki}_発症年齢"]; T = $colOf["既往歴${ki}_転帰"] } } }
$kazokuCols = @(); for ($ki = 1; $ki -le 9; $ki++) { if ($colOf.ContainsKey("家族歴${ki}_病名") -and $colOf.ContainsKey("家族歴${ki}_続柄")) { $kazokuCols += @{ B = $colOf["家族歴${ki}_病名"]; Z = $colOf["家族歴${ki}_続柄"] } } }
if (-not $Ymd) { $Ymd = Read-Host '受診日 (例 2026/11/24。このファイルの人が受けた日)' }
$dateList = @(Split-List $Ymd | ForEach-Object { Normalize-Ymd $_ } | Where-Object { $_ })
if ($dateList.Count -eq 0) { throw ("受診日「{0}」を解釈できません。2026/11/24 のように入力してください。" -f $Ymd) }
# ファイルにある受付番号
$ukeList = @(); $memoList = @(); $kazokuN = 0
foreach ($f in $rows[1..($rows.Count - 1)]) {
    $u = Normalize-KenNo (Get-Field $f $colOf['受付NO'])
    if ($u -notmatch '^\d+$') { continue }
    $ukeList += [int]$u
    if ($colOf.ContainsKey('確認メモ') -and (Normalize-Text (Get-Field $f $colOf['確認メモ'])) -ne '') { $memoList += $u }
    foreach ($kc in $kazokuCols) { if ((Normalize-Text (Get-Field $f $kc.B)) -ne '') { $kazokuN++ } }
}
if ($ukeList.Count -eq 0) { throw '受付番号のある行がありません。' }
$dupInFile = @($ukeList | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
if ($dupInFile.Count -gt 0) { throw ("ファイルの中に同じ受付番号が 2 行以上あります: {0}。直してからやり直してください。" -f ($dupInFile -join ', ')) }

Write-Host ''
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ' 既往歴・家族歴の取込 (受診票 OCR v5 → T_KIOU)' -ForegroundColor Cyan
Write-Host ('=' * 76) -ForegroundColor Cyan
Write-Host ("  ファイル: {0}" -f [System.IO.Path]::GetFileName($File))
Write-Host ("  受診日: {0}  受付番号 {1} 人分  本人の既往歴は 1 人 {2} 件まで / ファイルの家族歴 {3} 件" -f ($dateList -join ', '), $ukeList.Count, $Max, $kazokuN)
if ($memoList.Count -gt 0) { Write-Host ("  ★読み取りツールの「確認メモ」がある受付番号: {0}  → 紙と見比べてから入れてください" -f ($memoList -join ', ')) -ForegroundColor Yellow }

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
    # 本人の続柄名 (実データでいちばん多いもの)
    $selfMei = '本人'
    $z = Invoke-DbQuery $conn "SELECT TOP 1 LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) AS M, COUNT(*) AS C FROM T_KIOU WHERE LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) = '0' GROUP BY LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) ORDER BY 2 DESC"
    if ($z.Rows.Count -gt 0 -and [string]$z.Rows[0].M -ne '') { $selfMei = [string]$z.Rows[0].M }
    # 家族の行の作り (2026/11/26 check122: 続柄 1 父 / 2 母 / 3 兄弟姉妹 / 4 祖父母、病名コード 1〜17)。名前と列の入り方は実データでいちばん多いものに合わせる
    $ZM = @{ '1' = '父'; '2' = '母'; '3' = '兄弟姉妹'; '4' = '祖父母' }
    $FB = @{ 1 = '脳梗塞'; 2 = 'クモ膜下出血'; 3 = '脳出血'; 4 = '肺がん'; 5 = '高血圧'; 6 = '狭心症'; 7 = '心筋梗塞'; 8 = '食道がん'; 9 = '胃がん'; 10 = '大腸がん'
             11 = '肝がん'; 12 = '前立腺がん'; 13 = '糖尿病'; 14 = '子宮頸がん'; 15 = '子宮体がん'; 16 = '卵巣嚢腫（腫瘍）'; 17 = '乳がん' }
    $famAge = 0; $famCh = '0'; $famChM = ''; $famBk = ''
    if ($kazokuN -gt 0) {
        $FAMW = "LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) IN ('1','2','3','4')"
        $seenZ = @{}
        foreach ($r in (Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(CONVERT(varchar(20), ZOKUGARA))) AS Z, LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) AS M, COUNT(*) AS C FROM T_KIOU WHERE $FAMW AND LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) NOT IN ('', N'未設定') GROUP BY LTRIM(RTRIM(CONVERT(varchar(20), ZOKUGARA))), LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) ORDER BY 3 DESC").Rows) {
            if (-not $seenZ.ContainsKey([string]$r.Z)) { $seenZ[[string]$r.Z] = 1; $ZM[[string]$r.Z] = [string]$r.M }
        }
        $seenB = @{}
        foreach ($r in (Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(ISNULL(BYOMEI_CD,''))) AS BC, LTRIM(RTRIM(ISNULL(BYOMEI,''))) AS B, COUNT(*) AS C FROM T_KIOU WHERE $FAMW GROUP BY LTRIM(RTRIM(ISNULL(BYOMEI_CD,''))), LTRIM(RTRIM(ISNULL(BYOMEI,''))) ORDER BY 3 DESC").Rows) {
            $bcN = 0
            if ([string]$r.BC -match '^\d+$' -and [int]::TryParse([string]$r.BC, [ref]$bcN) -and $FB.ContainsKey($bcN) -and -not $seenB.ContainsKey($bcN) -and [string]$r.B -ne '') { $seenB[$bcN] = 1; $FB[$bcN] = [string]$r.B }
        }
        $tp = Invoke-DbQuery $conn "SELECT TOP 1 AGE, CHIRYO, CHIRYOMEI, BIKOU, COUNT(*) AS C FROM T_KIOU WHERE $FAMW GROUP BY AGE, CHIRYO, CHIRYOMEI, BIKOU ORDER BY 5 DESC"
        if ($tp.Rows.Count -gt 0) {
            $famAge = $tp.Rows[0].AGE; $famCh = [string]$tp.Rows[0].CHIRYO; $famChM = $tp.Rows[0].CHIRYOMEI; $famBk = $tp.Rows[0].BIKOU
        }
        Write-Host ("  家族の行の作り (健診ナビでいちばん多い形): 年齢 [{0}] / 治療CD [{1}] / 治療名 [{2}] / 備考 [{3}]" -f $(if ($famAge -is [DBNull]) { 'NULL' } else { $famAge }), $famCh, $(if ($famChM -is [DBNull]) { 'NULL' } else { $famChM }), $(if ($famBk -is [DBNull]) { 'NULL' } else { $famBk })) -ForegroundColor DarkGray
    }

    # 今の既往歴
    $exist = @{}
    if ($people.Count -gt 0) {
        $kLit = (@($people.Values | ForEach-Object { $_.Kojin }) | Sort-Object -Unique) -join ','
        foreach ($r in (Invoke-DbQuery $conn "SELECT KOJIN_ID, RENBAN, LTRIM(RTRIM(ISNULL(BYOMEI,''))) AS B, LTRIM(RTRIM(ISNULL(BYOMEI_CD,''))) AS BC, LTRIM(RTRIM(ISNULL(CHIRYO,''))) AS C, LTRIM(RTRIM(ISNULL(CHIRYOMEI,''))) AS CM, AGE, LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) AS ZK, LTRIM(RTRIM(ISNULL(ZOKUGARAMEI,''))) AS ZKM FROM T_KIOU WHERE KOJIN_ID IN ($kLit) ORDER BY KOJIN_ID, RENBAN").Rows) {
            $kj = [string]$r.KOJIN_ID
            if (-not $exist.ContainsKey($kj)) { $exist[$kj] = @{ MaxR = 0; Self = @{}; Count = 0; Holder = $null; Fam = @{}; FamHolder = $null; FamRows = @{} } }
            $rn = [int]$r.RENBAN; if ($rn -gt $exist[$kj].MaxR) { $exist[$kj].MaxR = $rn }
            if ([string]$r.ZK -eq '' -or [string]$r.ZK -eq '0') {
                if (Test-Placeholder ([string]$r.B)) {
                    if ($null -eq $exist[$kj].Holder) { $exist[$kj].Holder = @{ R = $rn; B = [string]$r.B; BC = [string]$r.BC; C = [string]$r.C; CM = [string]$r.CM; A = $(if ($r.AGE -eq [DBNull]::Value) { '' } else { [string]$r.AGE }) } }
                } else { $exist[$kj].Self[[string]$r.B] = @{ R = $rn; BC = [string]$r.BC; C = [string]$r.C; A = $(if ($r.AGE -eq [DBNull]::Value) { '' } else { [string][int]$r.AGE }); CM = [string]$r.CM; Old = $true }; $exist[$kj].Count++ }
            }
            else {
                # 家族の行。「なし」(続柄 未設定) は最初の 1 件で置き換える行、それ以外は 病名|続柄 で重複を見る
                if ((Test-Placeholder ([string]$r.B)) -or [string]$r.BC -eq '0') {
                    if ($null -eq $exist[$kj].FamHolder -and [string]$r.ZK -eq '9') { $exist[$kj].FamHolder = @{ R = $rn; B = [string]$r.B; BC = [string]$r.BC; ZK = [string]$r.ZK; ZKM = [string]$r.ZKM } }
                } else {
                    $exist[$kj].Fam[([string]$r.B + '|' + [string]$r.ZK)] = 1
                    # 同じ病名が別の続柄で入っているかを見るための控え (病名 → その行の一覧)
                    if (-not $exist[$kj].FamRows.ContainsKey([string]$r.B)) { $exist[$kj].FamRows[[string]$r.B] = @() }
                    $exist[$kj].FamRows[[string]$r.B] += @{ R = $rn; B = [string]$r.B; BC = [string]$r.BC; ZK = [string]$r.ZK; ZKM = [string]$r.ZKM; Used = $false }
                }
            }
        }
    }

    # 計画
    $plan = @(); $noHist = 0; $notFound = @()
    foreach ($f in $rows[1..($rows.Count - 1)]) {
        $uke = Normalize-KenNo (Get-Field $f $colOf['受付NO'])
        if ($uke -notmatch '^\d+$') { continue }
        $uke = [string][int]$uke
        $items = @()
        foreach ($kc in $kiouCols) {
            # K19 → 19、T05 → 5 (先頭の英字と空白を外す。数字だけでも読める)
            $bc = ((Normalize-Text (Get-Field $f $kc.B)).ToUpper() -replace '^[KＫ]\s*-?', '')
            if ($bc -eq '') { continue }
            $items += @{ BC = $bc; Age = (Normalize-KenNo (Get-Field $f $kc.A)); Ch = ((Normalize-Text (Get-Field $f $kc.T)).ToUpper() -replace '^[TＴ]\s*-?', '') }
        }
        # 家族歴: B13 → 13、続柄 Z-2 → 2 (Z-1,Z-2 のように 2 つあれば 1 つずつ)
        $famItems = @()
        foreach ($kc in $kazokuCols) {
            $famCd = ((Normalize-Text (Get-Field $f $kc.B)).ToUpper() -replace '^[BＢ]\s*-?', '')
            if ($famCd -eq '') { continue }
            $zs = @([regex]::Matches((Normalize-Text (Get-Field $f $kc.Z)), '\d+') | ForEach-Object { $_.Value })
            if ($zs.Count -eq 0) { $famItems += @{ BC = $famCd; Z = '' } } else { foreach ($zv in $zs) { $famItems += @{ BC = $famCd; Z = $zv } } }
        }
        if ($items.Count -eq 0 -and $famItems.Count -eq 0) { $noHist++; continue }
        if (-not $people.ContainsKey($uke)) { $notFound += $uke; continue }
        $p = $people[$uke]; $kj = $p.Kojin
        if (-not $exist.ContainsKey($kj)) { $exist[$kj] = @{ MaxR = 0; Self = @{}; Count = 0; Holder = $null; Fam = @{}; FamHolder = $null; FamRows = @{} } }
        $e = $exist[$kj]
        foreach ($it in $items) {
            $bn = 0; [void][int]::TryParse(($it.BC -replace '^0+(?=\d)', ''), [ref]$bn)
            $cn = -1; if ($it.Ch -ne '') { [void][int]::TryParse(($it.Ch -replace '^0+(?=\d)', ''), [ref]$cn) }
            $rep = New-Object PSObject -Property @{ 種類 = '本人'; 続柄 = ''; 続柄CD = '0'; 受付番号 = $uke; 受診日 = $p.Ymd; 氏名 = $p.Name; PkSeq = $p.Pk; KOJIN_ID = $kj
                病名CD = [string]$bn; 病名 = ''; 年齢 = $it.Age; 治療 = $(if ($cn -ge 0) { [string]$cn } else { '' }); 治療名 = ''
                操作 = ''; 連番 = ''; 前 = $null; 状態 = ''; 変更 = '' }
            if ($it.BC -notmatch '^\d+$') { $rep.病名CD = $it.BC; $rep.状態 = ("病名コード {0} を読めません" -f $it.BC); $plan += $rep; continue }
            if ($it.Ch -ne '' -and $it.Ch -notmatch '^\d+$') { $rep.状態 = ("転帰コード {0} を読めません" -f $it.Ch); $plan += $rep; continue }
            if (-not $BY.ContainsKey($bn)) { $rep.状態 = ("病名コード {0} が表にありません" -f $it.BC); $plan += $rep; continue }
            $rep.病名 = $BY[$bn]
            if ($cn -ge 0) { if ($CH.ContainsKey($cn)) { $rep.治療名 = $CH[$cn] } else { $rep.状態 = ("転帰コード {0} が表にありません" -f $it.Ch); $plan += $rep; continue } }
            if ($e.Self.ContainsKey($rep.病名)) {
                $rep.状態 = '既に入っています'
                $had = $e.Self[$rep.病名]
                if ($had -is [hashtable] -and $had.Old) {
                    # 前から健診ナビにある行。年齢・治療が受診票と違えば、受診票に合わせて書き換える (-KeepOld なら知らせるだけ)
                    $df = @(); $newAge = $had.A; $newC = $had.C; $newCM = $had.CM
                    if ($it.Age -match '^\d+$' -and ($had.A -eq '' -or [int]$it.Age -ne [int]$had.A)) {
                        $df += ("年齢 {0}→{1}" -f $(if ($had.A -eq '') { '(空)' } else { $had.A }), [int]$it.Age); $newAge = [string][int]$it.Age
                    }
                    if ($rep.治療名 -ne '' -and $rep.治療名 -ne $had.CM) {
                        $df += ("治療「{0}」→「{1}」" -f $had.CM, $rep.治療名); $newC = $rep.治療; $newCM = $rep.治療名
                    }
                    if ($df.Count -gt 0) {
                        if ($KeepOld) { $rep.状態 = '既にあり・中身が違う (書き換えない): ' + ($df -join '、') }
                        else {
                            $rep.操作 = '更新'; $rep.連番 = [string]$had.R; $rep.変更 = ($df -join '、')
                            $rep.前 = @{ R = $had.R; B = $rep.病名; BC = $had.BC; C = $had.C; CM = $had.CM; A = $had.A }
                            $rep.年齢 = $newAge; $rep.治療 = $newC; $rep.治療名 = $newCM
                            $had.A = $newAge; $had.C = $newC; $had.CM = $newCM; $had.Old = $false
                            $rep.状態 = 'OK'
                        }
                    }
                }
                $plan += $rep; continue
            }
            if ($e.Count -ge $Max) { $rep.状態 = ("本人の既往歴が既に {0} 件" -f $Max); $plan += $rep; continue }
            if ($null -ne $e.Holder) { $rep.操作 = '置換'; $rep.連番 = [string]$e.Holder.R; $rep.前 = $e.Holder; $e.Holder = $null }
            else { $e.MaxR++; $rep.操作 = '追加'; $rep.連番 = [string]$e.MaxR }
            $e.Self[$rep.病名] = @{ R = [int]$rep.連番; BC = $rep.病名CD; C = $rep.治療; A = $it.Age; CM = $rep.治療名; Old = $false }; $e.Count++
            $rep.状態 = 'OK'; $plan += $rep
        }
        $newRel = @{}
        foreach ($it in $famItems) {
            $bn2 = 0
            if ($it.BC -match '^\d+$' -and [int]::TryParse(($it.BC -replace '^0+(?=\d)', ''), [ref]$bn2) -and $FB.ContainsKey($bn2) -and $ZM.ContainsKey($it.Z)) {
                if (-not $newRel.ContainsKey($FB[$bn2])) { $newRel[$FB[$bn2]] = @{} }
                $newRel[$FB[$bn2]][$it.Z] = 1
            }
        }
        foreach ($it in $famItems) {
            $bn = 0; [void][int]::TryParse(($it.BC -replace '^0+(?=\d)', ''), [ref]$bn)
            $rep = New-Object PSObject -Property @{ 種類 = '家族'; 続柄 = ''; 続柄CD = $it.Z; 受付番号 = $uke; 受診日 = $p.Ymd; 氏名 = $p.Name; PkSeq = $p.Pk; KOJIN_ID = $kj
                病名CD = [string]$bn; 病名 = ''; 年齢 = ''; 治療 = ''; 治療名 = ''
                操作 = ''; 連番 = ''; 前 = $null; 状態 = ''; 変更 = '' }
            if ($it.BC -notmatch '^\d+$' -or -not $FB.ContainsKey($bn)) { $rep.病名CD = $it.BC; $rep.状態 = ("家族歴の病名コード {0} が表にありません" -f $it.BC); $plan += $rep; continue }
            $rep.病名 = $FB[$bn]
            if (-not $ZM.ContainsKey($it.Z)) { $rep.状態 = '続柄が読めていません (紙を見て Excel を直す)'; $plan += $rep; continue }
            $rep.続柄 = $ZM[$it.Z]
            $fk = $rep.病名 + '|' + $it.Z
            if ($e.Fam.ContainsKey($fk)) { $rep.状態 = '既に入っています'; $plan += $rep; continue }
            # 同じ病名が別の続柄で前から入っている行 (その続柄が受診票に無い行) があれば、その行の続柄を受診票に合わせる。
            # 足してしまうと「胃がん(父)」「胃がん(祖父母)」が並ぶため。-KeepOld なら知らせるだけ。
            $cand = $null
            if ($e.FamRows.ContainsKey($rep.病名)) {
                foreach ($fr in $e.FamRows[$rep.病名]) { if (-not $fr.Used -and -not $newRel[$rep.病名].ContainsKey($fr.ZK)) { $cand = $fr; break } }
            }
            if ($null -ne $cand) {
                $oldName = $(if ($cand.ZKM -ne '') { $cand.ZKM } else { '続柄CD ' + $cand.ZK })
                if ($KeepOld) { $rep.状態 = ("既にあり・続柄が違う (書き換えない): 健診ナビ「{0}」/ 受診票「{1}」" -f $oldName, $rep.続柄); $plan += $rep; continue }
                $cand.Used = $true
                $rep.操作 = '更新'; $rep.連番 = [string]$cand.R; $rep.変更 = ("続柄「{0}」→「{1}」" -f $oldName, $rep.続柄)
                $rep.前 = @{ R = $cand.R; B = $cand.B; BC = $cand.BC; ZK = $cand.ZK; ZKM = $cand.ZKM }
                [void]$e.Fam.Remove(($rep.病名 + '|' + $cand.ZK)); $e.Fam[$fk] = 1
                $rep.状態 = 'OK'; $plan += $rep; continue
            }
            if ($null -ne $e.FamHolder) { $rep.操作 = '置換'; $rep.連番 = [string]$e.FamHolder.R; $rep.前 = $e.FamHolder; $e.FamHolder = $null }
            else { $e.MaxR++; $rep.操作 = '追加'; $rep.連番 = [string]$e.MaxR }
            $e.Fam[$fk] = 1
            $rep.状態 = 'OK'; $plan += $rep
        }
    }

    Write-Host ''
    Write-Host '=== 入れる既往歴・家族歴 ===' -ForegroundColor Cyan
    $plan | Sort-Object { [int]$_.受付番号 }, 種類 | Select-Object 受付番号, 受診日, 氏名, 種類, 続柄, 病名CD, 病名, 年齢, 治療名, 操作, 変更, 連番, 状態 | Format-Table -AutoSize | Out-String -Width 300 | Write-Host
    $upd = @($plan | Where-Object { $_.操作 -eq '更新' -and $_.状態 -eq 'OK' })
    if ($upd.Count -gt 0) {
        Write-Host ("=== ★前から健診ナビに入っている内容を、受診票に合わせて書き換えるもの: {0} 件 ({1} 人) ===" -f $upd.Count, @($upd | Group-Object 受付番号).Count) -ForegroundColor Yellow
        $upd | Sort-Object { [int]$_.受付番号 }, 種類 | Select-Object 受付番号, 氏名, 種類, 病名, 変更 | Format-Table -AutoSize -Wrap | Out-String -Width 260 | Write-Host
        Write-Host '   読み取りの読み違いで前の正しい内容を消さないよう、ここに出た行は紙と見比べてから y を押してください。書き換えたくないときは -KeepOld を付けて実行します。' -ForegroundColor Yellow
    }
    $diff = @($plan | Where-Object { $_.状態 -like '既にあり・*' })
    if ($diff.Count -gt 0) {
        Write-Host ("=== 前から健診ナビに入っている内容と受診票が違うもの: {0} 件 ({1} 人) → 書き換えません。紙を見て、直すなら健診ナビの画面で ===" -f $diff.Count, @($diff | Group-Object 受付番号).Count) -ForegroundColor Yellow
        $diff | Sort-Object { [int]$_.受付番号 }, 種類 | Select-Object 受付番号, 種類, 病名, 状態 | Format-Table -AutoSize -Wrap | Out-String -Width 260 | Write-Host
    }
    $ok = @($plan | Where-Object { $_.状態 -eq 'OK' })
    Write-Host '=== 病名ごとの件数 ===' -ForegroundColor Cyan
    foreach ($g in ($ok | Group-Object 種類, 病名 | Sort-Object Count -Descending)) { Write-Host ("  {0,3} 件  {1}" -f $g.Count, $g.Name) }
    # 自動判定で変わるもの
    $names = @($ok | Where-Object { $_.種類 -eq '本人' } | ForEach-Object { $_.病名 } | Sort-Object -Unique)
    if ($names.Count -gt 0) {
        $inList = "N'" + (($names | ForEach-Object { $_ -replace "'", "''" }) -join "',N'") + "'"
        $hit = Invoke-DbQuery $conn "SELECT LTRIM(RTRIM(h.BYOMEI)) AS B, LTRIM(RTRIM(h.KOMOKU_CD)) AS CD, MAX(m.MEISYO1) AS NM, LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) AS H FROM T_KIOU_HANTEI h LEFT JOIN T_KOMOKU m ON LTRIM(RTRIM(m.KOMOKU_CD)) = LTRIM(RTRIM(h.KOMOKU_CD)) WHERE LTRIM(RTRIM(h.BYOMEI)) IN ($inList) GROUP BY LTRIM(RTRIM(h.BYOMEI)), LTRIM(RTRIM(h.KOMOKU_CD)), LTRIM(RTRIM(ISNULL(h.HANTEI_KIGO,''))) ORDER BY 1, 2"
        Write-Host ''
        Write-Host '=== 自動判定をかけたときに判定に効く病名 (健診ナビの対応表 T_KIOU_HANTEI) ===' -ForegroundColor Yellow
        foreach ($r in $hit.Rows) {
            $n = @($ok | Where-Object { $_.種類 -eq '本人' -and $_.病名 -eq [string]$r.B }).Count
            Write-Host ("  {0,-10} {1,3} 人 → {2} {3} が {4}" -f $r.B, $n, $r.CD, $r.NM, $r.H)
        }
    }
    Line
    Write-Host ("入れる: {0} 件 ({1} 人。本人 {5} 件・家族 {6} 件) / 既往歴も家族歴もなし: {2} 人 / 要確認: {3} 件 / 健診ナビに見つからない受付番号: {4}" -f $ok.Count, @($ok | Group-Object KOJIN_ID).Count, $noHist,
        @($plan | Where-Object { $_.状態 -ne 'OK' -and $_.状態 -ne '既に入っています' }).Count, $(if ($notFound.Count -gt 0) { $notFound -join ',' } else { 'なし' }),
        @($ok | Where-Object { $_.種類 -eq '本人' }).Count, @($ok | Where-Object { $_.種類 -eq '家族' }).Count) -ForegroundColor Green
    if ($ambiguous.Count -gt 0) { Write-Host ("★同じ受付番号の予約が 2 件以上あって決められない番号 (入れません): {0}  → 受診日を 1 日だけ指定してやり直してください" -f (($ambiguous.Keys | Sort-Object { [int]$_ }) -join ', ')) -ForegroundColor Red }
    if ($notFound.Count -gt 0) { Write-Host '  見つからない番号は、受診日が違うか、健診ナビに受付番号がまだ入っていません。' -ForegroundColor Yellow }
    Line
    if ($ok.Count -eq 0) { Write-Host '入れるものがありません。'; return }
    if (-not $Commit) {
        Write-Host '※ まだ何も書き込んでいません。' -ForegroundColor Yellow
        $a = Read-Host '書き込みますか (y = 書き込む / それ以外 = やめる)'
        if (($a -as [string]).Trim().ToLower() -ne 'y') { Write-Host 'やめました。何も変えていません。'; return }
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $undoFile = Join-Path $BackupDir ("kiou_v5_{0}.csv" -f $stamp)
    $undoRows = @(); $done = 0; $skipped = @()
    foreach ($grp in ($ok | Group-Object KOJIN_ID)) {
        $kj = $grp.Name; $pk = $grp.Group[0].PkSeq
        $lock = Test-Locked $conn $pk
        if ($lock) { Write-Host ("  [とばしました] {0} ({1}) は健診ナビで編集中です ({2})" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, $lock) -ForegroundColor Yellow; $skipped += $grp.Group[0].受付番号; continue }
        (Invoke-DbQuery $conn 'SELECT * FROM T_KIOU WHERE KOJIN_ID = @k' @{ k = $kj }) | Export-Csv -Path (Join-Path $BackupDir ("T_KIOU_{0}_{1}.csv" -f $kj, $stamp)) -NoTypeInformation -Encoding UTF8
        $tran = $conn.BeginTransaction()
        try {
            Acquire-AppLock $conn $tran ("KIOU_IMPORT:" + $kj)
            foreach ($t in $grp.Group) {
                $age = $(if ($t.年齢 -match '^\d+$') { [int]$t.年齢 } else { [DBNull]::Value })
                if ($t.種類 -eq '家族' -and $t.操作 -ne '追加') {
                    # 置換 = 家族の「なし」の行を書き換える / 更新 = 同じ病名の行の続柄を受診票に合わせる。病名と続柄だけ書き換える (年齢・治療の列はその行のまま)
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KIOU SET BYOMEI = @b, BYOMEI_CD = @bc, ZOKUGARA = @z, ZOKUGARAMEI = @zm WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @ob AND LTRIM(RTRIM(ISNULL(CONVERT(varchar(20), ZOKUGARA),''))) = @oz" `
                        @{ b = $t.病名; bc = $t.病名CD; z = $t.続柄CD; zm = $t.続柄; k = $kj; r = [int]$t.連番; ob = $t.前.B; oz = $t.前.ZK }
                }
                elseif ($t.種類 -eq '家族') {
                    $n = Invoke-DbExec $conn $tran "INSERT INTO T_KIOU (KOJIN_ID, RENBAN, BYOMEI, BYOMEI_CD, CHIRYO, CHIRYOMEI, AGE, ZOKUGARA, ZOKUGARAMEI, BIKOU) VALUES (@k, @r, @b, @bc, @c, @cm, @a, @z, @zm, @bk)" `
                        @{ k = $kj; r = [int]$t.連番; b = $t.病名; bc = $t.病名CD; c = $famCh; cm = $famChM; a = $famAge; z = $t.続柄CD; zm = $t.続柄; bk = $famBk }
                }
                elseif ($t.操作 -ne '追加') {
                    # 置換 = 「特記事項なし」の行を書き換える / 更新 = 同じ病名の行の年齢・治療を受診票に合わせる
                    $n = Invoke-DbExec $conn $tran "UPDATE T_KIOU SET BYOMEI = @b, BYOMEI_CD = @bc, CHIRYO = @c, CHIRYOMEI = @cm, AGE = @a WHERE KOJIN_ID = @k AND RENBAN = @r AND LTRIM(RTRIM(ISNULL(BYOMEI,''))) = @ob" `
                        @{ b = $t.病名; bc = $t.病名CD; c = $t.治療; cm = $t.治療名; a = $age; k = $kj; r = [int]$t.連番; ob = $t.前.B }
                } else {
                    $n = Invoke-DbExec $conn $tran "INSERT INTO T_KIOU (KOJIN_ID, RENBAN, BYOMEI, BYOMEI_CD, CHIRYO, CHIRYOMEI, AGE, ZOKUGARA, ZOKUGARAMEI, BIKOU) VALUES (@k, @r, @b, @bc, @c, @cm, @a, '0', @zm, '')" `
                        @{ k = $kj; r = [int]$t.連番; b = $t.病名; bc = $t.病名CD; c = $t.治療; cm = $t.治療名; a = $age; zm = $selfMei }
                }
                if ($n -ne 1) { throw ("書込の影響行数が {0} でした ({1} / {2})。ロールバックします。" -f $n, $t.氏名, $t.病名) }
                $undoRows += New-Object PSObject -Property @{ KOJIN_ID = $kj; RENBAN = $t.連番; 操作 = $t.操作; 病名 = $t.病名; 受付番号 = $t.受付番号; 氏名 = $t.氏名
                    種類 = $t.種類; 続柄CD = $t.続柄CD; 前の続柄CD = $(if ($t.前 -and $t.種類 -eq '家族') { $t.前.ZK } else { '' }); 前の続柄名 = $(if ($t.前 -and $t.種類 -eq '家族') { $t.前.ZKM } else { '' })
                    前の病名 = $(if ($t.前) { $t.前.B } else { '' }); 前の病名CD = $(if ($t.前) { $t.前.BC } else { '' }); 前の治療 = $(if ($t.前 -and $t.種類 -eq '本人') { $t.前.C } else { '' })
                    前の治療名 = $(if ($t.前 -and $t.種類 -eq '本人') { $t.前.CM } else { '' }); 前の年齢 = $(if ($t.前 -and $t.種類 -eq '本人') { $t.前.A } else { '' }) }
                $done++
            }
            $tran.Commit()
            Write-Host ("  [書込] {0} ({1}) : {2}" -f $grp.Group[0].氏名, $grp.Group[0].受付番号, (($grp.Group | ForEach-Object { "{0}({1})" -f $_.病名, $(if ($_.種類 -eq '家族') { $_.続柄 } else { $_.治療名 }) }) -join ' / ')) -ForegroundColor Green
        } catch { $tran.Rollback(); throw }
    }
    if ($undoRows.Count -gt 0) { $undoRows | Select-Object KOJIN_ID, RENBAN, 種類, 操作, 病名, 続柄CD, 前の病名, 前の病名CD, 前の治療, 前の治療名, 前の年齢, 前の続柄CD, 前の続柄名, 受付番号, 氏名 | Export-Csv -Path $undoFile -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host ("[完了] {0} 件の既往歴・家族歴を入れました。記録: {1}" -f $done, $undoFile) -ForegroundColor Green
    if ($skipped.Count -gt 0) { Write-Host ("  ★健診ナビで開かれていて書けなかった受付番号: {0}" -f ($skipped -join ', ')) -ForegroundColor Red }
    Write-Host '   このあと自動判定をかけると、高血圧・糖尿病・高脂血症 (治療中) の人の判定が J などになります。元に戻すときは 受診票既往歴_元に戻す.bat。' -ForegroundColor Yellow
}
finally { $conn.Close() }
