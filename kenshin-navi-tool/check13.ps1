<#
  テスト用の単独PCを作れるか調べる (check13.ps1)
  check13.bat をダブルクリックすると実行され、結果 r_check13.txt がメモ帳で開きます。
  DBは読むだけで、一切変更しません。

  やりたいこと: サーバーに繋がない1台のPCだけで健診ナビを動かし、
  本番に影響しない場所で取込や請求の試験をしたい。

  そのために必要な情報を集める:
    1. SQL Server の版とバージョン (同じ版を用意する必要がある)
    2. DBの大きさ (SQL Server Express は 10GB までという制限がある)
    3. バックアップの設定 (復元用の .bak が作れるか)
    4. 健診ナビ側のフォルダの大きさ (帳票などを丸ごとコピーする必要がある)
    5. 接続先の書き方 (テストPCではここを書き換える)
#>
$ErrorActionPreference = 'Continue'
$dir = $PSScriptRoot
$tool = Join-Path $dir 'db_tool.ps1'
$out = Join-Path $dir 'r_check13.txt'

function W($t) { $t | Out-File $out -Append -Encoding Default }
function Q($title, $sql, $max) {
    W ''
    W $title
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tool -Sql $sql -MaxRows $max *>&1 | Out-File $out -Append -Encoding Default
}

"=== テスト用の単独PCを作れるか 調査 $(Get-Date -Format 'yyyy/MM/dd HH:mm') ===" | Out-File $out -Encoding Default
if (-not (Test-Path $tool)) { W "db_tool.ps1 がありません: $dir"; notepad $out; return }

Q '--- 1. SQL Server の版とバージョン (テストPCにも同じか新しい版を入れる) ---' @"
SELECT SERVERPROPERTY('Edition') AS 版,
       SERVERPROPERTY('ProductVersion') AS バージョン,
       SERVERPROPERTY('ProductLevel') AS SP,
       SERVERPROPERTY('Collation') AS 照合順序,
       @@SERVERNAME AS サーバー名
"@ 5

Q '--- 2. DBの大きさ (Express は 1DB 10GB まで) ---' @"
SELECT DB_NAME() AS DB名,
       CAST(SUM(CASE WHEN type_desc = 'ROWS' THEN size END) * 8.0 / 1024 AS decimal(10,1)) AS データMB,
       CAST(SUM(CASE WHEN type_desc = 'LOG'  THEN size END) * 8.0 / 1024 AS decimal(10,1)) AS ログMB
FROM sys.database_files
"@ 5

Q '--- 3. 行数の多いテーブル (コピーする量の目安) ---' @"
SELECT TOP 12 t.name AS テーブル, SUM(p.rows) AS 行数
FROM sys.tables t JOIN sys.partitions p ON p.object_id = t.object_id AND p.index_id IN (0,1)
GROUP BY t.name ORDER BY SUM(p.rows) DESC
"@ 15

Q '--- 4. 復旧モデルと直近のバックアップ (.bak があればそれを復元するのが一番早い) ---' @"
SELECT d.name AS DB名, d.recovery_model_desc AS 復旧モデル,
       (SELECT MAX(backup_finish_date) FROM msdb.dbo.backupset b
         WHERE b.database_name = d.name AND b.type = 'D') AS 最後のフルバックアップ
FROM sys.databases d WHERE d.name = DB_NAME()
"@ 5

Q '--- 5. バックアップの保存先 (この .bak をテストPCへ持っていく) ---' @"
SELECT TOP 5 b.backup_finish_date AS 日時,
       CAST(b.backup_size / 1024 / 1024 AS decimal(10,1)) AS サイズMB,
       m.physical_device_name AS 保存先
FROM msdb.dbo.backupset b
JOIN msdb.dbo.backupmediafamily m ON m.media_set_id = b.media_set_id
WHERE b.database_name = DB_NAME() AND b.type = 'D'
ORDER BY b.backup_finish_date DESC
"@ 10

# ---- 健診ナビ側のファイル ----
W ''
W '--- 6. 接続先の設定ファイル (テストPCではここを書き換える) ---'
$cf = '\\KNSV\KenshinNavi\SQLSV\SQLServerConnect.txt'
if (Test-Path $cf) {
    W "場所: $cf"
    foreach ($ln in (Get-Content $cf)) {
        # パスワードらしい行は伏せる
        if ($ln -match '(?i)(password|pwd)\s*[=:]') { W '  (パスワードの行は伏せました)' }
        else { W "  $ln" }
    }
} else {
    W "見つかりません: $cf"
}

W ''
W '--- 7. 健診ナビのフォルダの大きさ (丸ごとコピーする分) ---'
foreach ($p in @('\\KNSV\KenshinNavi')) {
    if (-not (Test-Path $p)) { W "  $p : 見つかりません"; continue }
    try {
        $items = Get-ChildItem -Path $p -Recurse -File -ErrorAction SilentlyContinue
        $mb = [math]::Round((($items | Measure-Object Length -Sum).Sum / 1MB), 1)
        W ("  {0} : {1} ファイル / {2} MB" -f $p, $items.Count, $mb)
        W '  直下のフォルダ:'
        foreach ($d in (Get-ChildItem -Path $p -Directory -ErrorAction SilentlyContinue)) { W "    $($d.Name)" }
    } catch { W ("  読めませんでした: {0}" -f $_.Exception.Message) }
}

W ''
W '=== 完了。この内容をチャットに貼り付けてください ==='
notepad $out
