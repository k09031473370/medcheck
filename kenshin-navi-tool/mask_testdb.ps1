<#
.SYNOPSIS
  練習用DBの氏名などを架空のものに置き換える (mask_testdb.ps1)

.DESCRIPTION
  自分のPCへ復元したコピーDBに対して実行し、
  氏名・住所・電話・メール・生年月日などを架空の値に置き換える。
  検査結果や金額はそのまま残るので、取込や請求の試験にはそのまま使える。

  ★これは「復元したコピー」専用です。本番サーバーでは動きません。
    サーバー名が KNSV のときは、はっきり拒否します。

  既定は表示のみ。実際に書き換えるには -Commit を付ける。

.EXAMPLE
  # 何が置き換わるかを見る (書き換えません)
  powershell -ExecutionPolicy Bypass -File mask_testdb.ps1 -ConnectionString "Data Source=localhost\SQLEXPRESS;Initial Catalog=K166_SIBAURAUSER;Integrated Security=True"

  # 実際に置き換える
  powershell -ExecutionPolicy Bypass -File mask_testdb.ps1 -ConnectionString "..." -Commit
#>
[CmdletBinding()]
param(
    [switch]$Commit,
    [string[]]$BlockServer = @('KNSV'),   # このサーバー名を含むなら絶対に実行しない
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
Invoke-Expression (Get-Part 'function Normalize-Text' 'function Normalize-KenNo')
Invoke-Expression (Get-Part 'function Resolve-ConnectionString' 'function Get-CurrentKensa')

$conn = Open-Db
try {
    # ---- 安全確認: 本番では絶対に動かさない ----
    $info = Invoke-DbQuery $conn 'SELECT @@SERVERNAME AS SRV, DB_NAME() AS DB, HOST_NAME() AS HOSTNM' $null
    $srv = [string]$info.Rows[0].SRV
    $db  = [string]$info.Rows[0].DB

    Write-Host ''
    Write-Host '=== 練習用DBの仮名化 ===' -ForegroundColor Cyan
    Write-Host ("  サーバー : {0}" -f $srv)
    Write-Host ("  データベース: {0}" -f $db)
    Write-Host ''

    foreach ($b in $BlockServer) {
        if ($b -ne '' -and $srv -like "*$b*") {
            Write-Host "【中止】サーバー名に「$b」が含まれています。本番の可能性があるため実行しません。" -ForegroundColor Red
            Write-Host '        自分のPCへ復元したコピーに対して実行してください。' -ForegroundColor Red
            return
        }
    }

    # ---- 置き換える内容 ----
    # 検査結果・金額・コースはそのまま。人が特定できる情報だけを架空にする。
    $steps = @(
        @{ 名前 = '氏名(漢字・カナ・旧姓・英語)'; 表 = 'T_KOJIN1'
           Sql = @"
UPDATE T_KOJIN1 SET
  KANJI_SIMEI = N'検査' + RIGHT('0000' + CAST(KOJIN_ID AS varchar(10)), 4) + N'太郎',
  KANA_SIMEI  = N'ｹﾝｻ' + RIGHT('0000' + CAST(KOJIN_ID AS varchar(10)), 4) + N'ﾀﾛｳ',
  KANJI_SIMEI_OLD = NULL, KANA_SIMEI_OLD = NULL, EIGO_SIMEI = NULL
"@ }
        @{ 名前 = '社員番号・カルテNo・手帳No・被曝No'; 表 = 'T_KOJIN1'
           Sql = @"
UPDATE T_KOJIN1 SET
  KOJIN_NO  = CASE WHEN LTRIM(RTRIM(ISNULL(KOJIN_NO,'')))  = '' THEN KOJIN_NO  ELSE 'E' + CAST(KOJIN_ID AS varchar(10)) END,
  KARUTE_NO = CASE WHEN LTRIM(RTRIM(ISNULL(KARUTE_NO,''))) = '' THEN KARUTE_NO ELSE 'K' + CAST(KOJIN_ID AS varchar(10)) END,
  TECHO_NO = NULL, HIBAKU_NO = NULL
"@ }
        @{ 名前 = '生年月日 (年はそのまま、月日を1/1に寄せる)'; 表 = 'T_KOJIN1'
           Sql = @"
UPDATE T_KOJIN1 SET D_SEINEN = LEFT(D_SEINEN, 4) + '/01/01'
WHERE LEN(LTRIM(RTRIM(ISNULL(D_SEINEN,'')))) >= 8
"@ }
        @{ 名前 = '住所・電話・メール・保険証番号など'; 表 = 'T_KOJIN2'
           Sql = @"
UPDATE T_KOJIN2 SET
  YUUBIN = NULL, JYUSYO1 = N'（練習用）', JYUSYO2 = NULL, JYUSYO3 = NULL,
  TEL = NULL, FAX = NULL, TEL_REN = NULL, TEL2 = NULL, E_MAIL = NULL,
  CYUUI_MEMO = NULL, BIKOU = NULL,
  SOUFU_YUUBIN = NULL, SOUFU_JYUSYO1 = NULL, SOUFU_JYUSYO2 = NULL, SOUFU_JYUSYO3 = NULL,
  JYUSYO_EIGO = NULL, JYUSYO_KAIGAI = NULL, E_MAIL_KAIGAI = NULL,
  TEL_KAIGAI = NULL, FAX_KAIGAI = NULL, YUUBIN_KAIGAI = NULL,
  KENPO_KIGO = NULL, KENPO_BANGO = NULL, KENPO_BANGO_EDA = NULL,
  S_KENPO_KIGO = NULL, S_KENPO_HUGO = NULL, SYUSSEKI_NO = NULL
"@ }
        @{ 名前 = '受診の備考・注意メモ'; 表 = 'T_KENSIN'
           Sql = "UPDATE T_KENSIN SET BIKO = NULL, TYUUI = NULL, IKANSEN_MEMO = NULL" }
        @{ 名前 = '住所テーブル'; 表 = 'T_JYUSYO'; 任意 = $true
           Sql = "DELETE FROM T_JYUSYO" }
        @{ 名前 = '変更履歴ログ (氏名が残るため消す)'; 表 = 'T_KENSIN_LOG'; 任意 = $true
           Sql = "DELETE FROM T_KENSIN_LOG" }
        @{ 名前 = '結果変更ログ'; 表 = 'T_KEKKA_LOG'; 任意 = $true
           Sql = "DELETE FROM T_KEKKA_LOG" }
    )

    Write-Host '--- 置き換える内容 ---'
    foreach ($st in $steps) {
        $t = $st.表
        $exists = Invoke-DbQuery $conn "SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @t" @{ t = $t }
        if ([int]$exists.Rows[0].N -eq 0) {
            Write-Host ("  [skip] {0} … {1} が無い" -f $st.名前, $t) -ForegroundColor DarkGray
            $st.スキップ = $true
            continue
        }
        $cnt = Invoke-DbQuery $conn "SELECT COUNT(*) AS N FROM [$t]" $null
        Write-Host ("  {0,-40} {1} ({2} 行)" -f $st.名前, $t, [int]$cnt.Rows[0].N)
    }

    Write-Host ''
    Write-Host '  ※ 検査結果・判定・コース・料金・事業所名はそのまま残します。' -ForegroundColor DarkGray
    Write-Host '     取込や請求の試験はそのままできます。' -ForegroundColor DarkGray

    if (-not $Commit) {
        Write-Host ''
        Write-Host '※ 表示のみです。実際に置き換えるには -Commit を付けて再実行してください。' -ForegroundColor Yellow
        return
    }

    # ---- 実行前にサーバー名を打たせる (取り違え防止) ----
    Write-Host ''
    Write-Host ("これから {0} の {1} を書き換えます。元には戻せません。" -f $srv, $db) -ForegroundColor Yellow
    $ans = Read-Host "間違いなければ、サーバー名「$srv」をそのまま入力してください"
    if ((Normalize-Text $ans) -ne (Normalize-Text $srv)) {
        Write-Host '入力が一致しませんでした。中止します。' -ForegroundColor Red
        return
    }

    $tran = $conn.BeginTransaction()
    try {
        $done = 0
        foreach ($st in $steps) {
            if ($st.スキップ) { continue }
            $n = Invoke-DbExec $conn $tran $st.Sql @{}
            Write-Host ("  {0,-40} {1} 行" -f $st.名前, $n) -ForegroundColor Green
            $done++
        }
        $tran.Commit()
        Write-Host ''
        Write-Host ("[完了] {0} 項目を架空の値に置き換えました。" -f $done) -ForegroundColor Green
        Write-Host '氏名は「検査0001太郎」のような形になっています。' -ForegroundColor DarkGray
    }
    catch { $tran.Rollback(); throw }
}
finally { $conn.Close() }
