<#
  form フォルダの中身が正しいか調べる (form確認.ps1)
  form確認.bat をダブルクリックするだけ。DBには繋ぎません。

  対応表を差し替えたつもりが古いままだった、
  ダウンロードで (1) が付いた別名で保存されていた、
  といったことをその場で見分けるための道具です。
#>
$ErrorActionPreference = 'Continue'
$dir    = $PSScriptRoot
$MapDir = Join-Path $dir 'form'

function Line { Write-Host ('-' * 70) -ForegroundColor DarkGray }

Write-Host ''
Write-Host ('=' * 70) -ForegroundColor Cyan
Write-Host ' form フォルダの中身の確認' -ForegroundColor Cyan
Write-Host ('=' * 70) -ForegroundColor Cyan
Write-Host "  場所: $MapDir"

if (-not (Test-Path $MapDir)) {
    Write-Host '  ★ form フォルダがありません。' -ForegroundColor Red
    return
}

# ---- 1. 中身の一覧 ----
Write-Host ''
Write-Host '1. form フォルダにある csv' -ForegroundColor White
Line
foreach ($f in (Get-ChildItem -Path $MapDir -Filter '*.csv' -File | Sort-Object Name)) {
    $mark = ''
    if ($f.Name -match '\(\d+\)') { $mark = '  ★名前に (1) が付いています。差し替え失敗の可能性' }
    $color = if ($mark) { 'Red' } else { 'Gray' }
    Write-Host ('  {0,-34} {1,8:N0} バイト  {2}{3}' -f $f.Name, $f.Length, $f.LastWriteTime.ToString('MM/dd HH:mm'), $mark) -ForegroundColor $color
}

# ---- 2. 対応表の中身を見る ----
Write-Host ''
Write-Host '2. 対応表の中身' -ForegroundColor White
Line
Write-Host ('  {0,-30} {1,5} {2,7} {3,7} {4}' -f '対応表', '行数', '氏名漢字', '氏名カナ', '判定')
Line
foreach ($f in (Get-ChildItem -Path $MapDir -Filter 'mapping*.csv' -File | Sort-Object Name)) {
    $rows = @()
    try { $rows = @(Import-Csv -Path $f.FullName -Encoding UTF8 | Where-Object { ($_.Col -as [string]).Trim() -notlike '#*' }) } catch { }
    $kanji = @($rows | Where-Object { ($_.Kind -as [string]).Trim().ToUpper() -eq 'NAMEKANJI' }).Count
    $kana  = @($rows | Where-Object { ($_.Kind -as [string]).Trim().ToUpper() -eq 'NAMEKANA' }).Count
    $kenno = @($rows | Where-Object { ($_.Kind -as [string]).Trim().ToUpper() -eq 'KENNO' }).Count
    if ($rows.Count -eq 0)                      { $v = '★読めません'; $c = 'Red' }
    elseif ($kenno -gt 0)                       { $v = 'OK (受付番号あり)'; $c = 'Green' }
    elseif ($kanji -gt 0 -or $kana -gt 0)       { $v = 'OK (氏名で照合)'; $c = 'Green' }
    else                                        { $v = '★受付番号も氏名も無い'; $c = 'Red' }
    Write-Host ('  {0,-30} {1,5} {2,7} {3,7} {4}' -f $f.Name, $rows.Count, $kanji, $kana, $v) -ForegroundColor $c
}

# ---- 3. 東振協250列の対応表を詳しく ----
Write-Host ''
Write-Host '3. 今回使う対応表 (mapping_tosinkyo250.csv)' -ForegroundColor White
Line
$p = Join-Path $MapDir 'mapping_tosinkyo250.csv'
if (-not (Test-Path $p)) {
    Write-Host '  ★ ありません。form フォルダに入れてください。' -ForegroundColor Red
}
else {
    $fi = Get-Item $p
    Write-Host ('  サイズ  : {0:N0} バイト' -f $fi.Length)
    Write-Host ('  更新日時: {0}' -f $fi.LastWriteTime)
    $rows = @()
    try { $rows = @(Import-Csv -Path $p -Encoding UTF8 | Where-Object { ($_.Col -as [string]).Trim() -notlike '#*' }) } catch { Write-Host "  ★読み込みに失敗: $_" -ForegroundColor Red }
    Write-Host ('  行数    : {0}' -f $rows.Count)
    # 見出し行(Col,Col2,...)がファイルの1行目にあるか。
    # 8行目などにあると Windows PowerShell 5.1 の Import-Csv が
    # コメント行を見出しと誤読して、Kind列が丸ごと消える。
    $head1 = (Get-Content $p -TotalCount 1 -Encoding UTF8) -replace "^\uFEFF", ''
    if ($head1 -notlike 'Col,Col2,Label,Kind*') {
        Write-Host '  ★NG  見出し行(Col,Col2,Label,Kind...)が1行目にありません' -ForegroundColor Red
        Write-Host ('        1行目: {0}' -f $head1) -ForegroundColor Red
    } else {
        Write-Host '  OK   見出し行が1行目にあります' -ForegroundColor Green
    }
    $kinds = @($rows | ForEach-Object { ($_.Kind -as [string]).Trim().ToUpper() } |
               Where-Object { $_ -ne '' } | Group-Object | Sort-Object Name |
               ForEach-Object { '{0}×{1}' -f $_.Name, $_.Count })
    Write-Host ('  種類    : {0}' -f ($kinds -join ' / '))
    Write-Host ''
    $need = @{ 'NAMEKANJI' = 4; 'NAMEKANA' = 5; 'KENYMD' = 1 }
    $ng = 0
    foreach ($k in ($need.Keys | Sort-Object)) {
        $hit = @($rows | Where-Object { ($_.Kind -as [string]).Trim().ToUpper() -eq $k })
        if ($hit.Count -eq 1 -and ([int]($hit[0].Col)) -eq $need[$k]) {
            Write-Host ('  OK   {0,-10} {1} 列目' -f $k, $need[$k]) -ForegroundColor Green
        } else {
            $ng++
            Write-Host ('  ★NG  {0,-10} {1} 列目にあるはずですが見つかりません' -f $k, $need[$k]) -ForegroundColor Red
        }
    }
    Write-Host ''
    if ($ng -eq 0 -and $head1 -like 'Col,Col2,Label,Kind*') {
        Write-Host '  ★ この対応表は正しいものです。取込に進めます。' -ForegroundColor Green
    } else {
        Write-Host '  ★ 対応表が古いか壊れています。送り直したものに差し替えてください。' -ForegroundColor Red
        Write-Host '     ダウンロードした場所に mapping_tosinkyo250 (1).csv のような'
        Write-Host '     別名のファイルが無いか確認してください。'
    }
}

Write-Host ''
Write-Host '何かキーを押すと閉じます...'
