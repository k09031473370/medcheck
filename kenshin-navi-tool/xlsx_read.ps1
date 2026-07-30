<#
  xlsx_read.ps1 — .xlsx / .xlsm を Excel を使わずに読むための共通部品

  form_import.ps1 と yoyaku_export.ps1 から読み込んで使う。
  Excelを起動しないので、OneDrive上・USB上のファイルでも
  「保護ビュー」等のダイアログで固まることがない。
#>
# ---------------------------------------------------------------------------
# .xlsx / .xlsm を Excel を使わずに読む
#   Excel COM は OneDrive上・USB上のファイルで「保護ビュー」等のダイアログを
#   裏側(非表示)で出してしまい、そのまま固まることがある。
#   xlsx は中身がZIP+XMLなので、直接読めばExcelが無くても・止まらずに読める。
# ---------------------------------------------------------------------------
function Read-Xlsx([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $NSP = 'http://schemas.openxmlformats.org/package/2006/relationships'

    $zip = $null
    try { $zip = [System.IO.Compression.ZipFile]::OpenRead($path) }
    catch { throw ("Excelファイルを開けません: {0}`n({1})" -f $path, $_.Exception.Message) }

    try {
        # XmlDocument も XmlNamespaceManager も IEnumerable なので、
        # 「, 」を付けて返さないとPowerShellが中身を展開してしまう。
        $getXml = {
            param($name)
            $ent = $null
            foreach ($e in $zip.Entries) { if ($e.FullName -eq $name) { $ent = $e; break } }
            if (-not $ent) { return $null }
            $st = $ent.Open()
            $sr = New-Object System.IO.StreamReader($st, [System.Text.Encoding]::UTF8)
            $txt = $sr.ReadToEnd()
            $sr.Close(); $st.Close()
            $doc = New-Object System.Xml.XmlDocument
            $doc.LoadXml($txt)
            return ,$doc
        }
        $nsMgr = {
            param($doc, $prefix, $uri)
            $m = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
            $m.AddNamespace($prefix, $uri)
            return ,$m
        }

        # --- 共有文字列 ---
        $shared = @()
        $xs = & $getXml 'xl/sharedStrings.xml'
        if ($xs) {
            $m = & $nsMgr $xs 'm' $NS
            foreach ($si in $xs.SelectNodes('/m:sst/m:si', $m)) {
                # m:t (単純な文字列) と m:r/m:t (書式付きの断片) だけを連結する。
                # m:rPh/m:t は「ふりがな」なので値ではない (健保外 + ケンポガイ になってしまう)
                $sb = New-Object System.Text.StringBuilder
                foreach ($t in $si.SelectNodes('m:t | m:r/m:t', $m)) { [void]$sb.Append($t.InnerText) }
                $shared += $sb.ToString()
            }
        }

        # --- 書式(日付かどうか) ---
        $dateXf = @{}
        $xst = & $getXml 'xl/styles.xml'
        if ($xst) {
            $m = & $nsMgr $xst 'm' $NS
            $custom = @{}
            foreach ($f in $xst.SelectNodes('/m:styleSheet/m:numFmts/m:numFmt', $m)) {
                $fid = 0
                if ([int]::TryParse($f.GetAttribute('numFmtId'), [ref]$fid)) { $custom[$fid] = $f.GetAttribute('formatCode') }
            }
            $i = 0
            foreach ($xf in $xst.SelectNodes('/m:styleSheet/m:cellXfs/m:xf', $m)) {
                $id = 0
                [void][int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$id)
                $isDate = $false
                if (($id -ge 14 -and $id -le 22) -or ($id -ge 45 -and $id -le 47)) { $isDate = $true }
                elseif ($custom.ContainsKey($id)) {
                    # 色指定[Red]やリテラル"年"を除いてから y / d を探す
                    $code = $custom[$id] -replace '\[[^\]]*\]', '' -replace '"[^"]*"', ''
                    if ($code -match '[yYdD]') { $isDate = $true }
                }
                $dateXf[$i] = $isDate
                $i++
            }
        }

        # --- 先頭シートのパス ---
        $sheetPath = 'xl/worksheets/sheet1.xml'
        $xw = & $getXml 'xl/workbook.xml'
        $xr = & $getXml 'xl/_rels/workbook.xml.rels'
        if ($xw -and $xr) {
            $mw = & $nsMgr $xw 'm' $NS
            $first = $xw.SelectSingleNode('/m:workbook/m:sheets/m:sheet', $mw)
            if ($first) {
                $rid = $first.GetAttribute('id', $NSR)
                $mr = & $nsMgr $xr 'p' $NSP
                foreach ($rel in $xr.SelectNodes('/p:Relationships/p:Relationship', $mr)) {
                    if ($rel.GetAttribute('Id') -ne $rid) { continue }
                    $t = $rel.GetAttribute('Target') -replace '^/', ''
                    if ($t -notlike 'xl/*') { $t = 'xl/' + $t }
                    $sheetPath = $t
                }
            }
        }

        # --- セルを読む ---
        $xsh = & $getXml $sheetPath
        if (-not $xsh) { throw "Excelの中にシートが見つかりません: $path" }
        $m = & $nsMgr $xsh 'm' $NS

        $lines = @()
        $maxCol = 0
        foreach ($row in $xsh.SelectNodes('/m:worksheet/m:sheetData/m:row', $m)) {
            $h = @{}
            foreach ($c in $row.SelectNodes('m:c', $m)) {
                # 列番号 (A→1, AB→28)
                $col = 0
                foreach ($ch in $c.GetAttribute('r').ToCharArray()) {
                    $n = [int][char]$ch
                    if     ($n -ge 65 -and $n -le 90) { $col = $col * 26 + ($n - 64) }
                    elseif ($n -ge 97 -and $n -le 122) { $col = $col * 26 + ($n - 96) }
                    else { break }
                }
                if ($col -le 0) { continue }

                $t = $c.GetAttribute('t')
                $vNode = $c.SelectSingleNode('m:v', $m)
                $val = ''
                if ($t -eq 'inlineStr') {
                    $isN = $c.SelectSingleNode('m:is', $m)
                    if ($isN) { foreach ($tn in $isN.SelectNodes('.//m:t', $m)) { $val += $tn.InnerText } }
                }
                elseif ($t -eq 's') {
                    if ($vNode) {
                        $ix = -1
                        if ([int]::TryParse($vNode.InnerText, [ref]$ix) -and $ix -ge 0 -and $ix -lt $shared.Count) { $val = $shared[$ix] }
                    }
                }
                elseif ($t -eq 'b') { if ($vNode) { $val = if ($vNode.InnerText -eq '1') { 'TRUE' } else { 'FALSE' } } }
                elseif ($t -eq 'str' -or $t -eq 'e') { if ($vNode) { $val = $vNode.InnerText } }
                elseif ($vNode) {
                    $raw = $vNode.InnerText
                    $d = 0.0
                    if ([double]::TryParse($raw, [System.Globalization.NumberStyles]::Float,
                            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
                        $sIdx = -1
                        $sAttr = $c.GetAttribute('s')
                        if ($sAttr -ne '') { $tmp = 0; if ([int]::TryParse($sAttr, [ref]$tmp)) { $sIdx = $tmp } }
                        if ($sIdx -ge 0 -and $dateXf.ContainsKey($sIdx) -and $dateXf[$sIdx] -and $d -gt 0) {
                            $val = ([datetime]::FromOADate($d)).ToString('yyyy/MM/dd')
                        }
                        elseif ($d -eq [math]::Floor($d) -and [math]::Abs($d) -lt 1e15) { $val = [string][long]$d }
                        else { $val = $d.ToString('0.##########', [System.Globalization.CultureInfo]::InvariantCulture) }
                    }
                    else { $val = $raw }
                }

                if ($val -ne '') {
                    $h[$col] = $val
                    if ($col -gt $maxCol) { $maxCol = $col }
                }
            }
            $lines += ,$h
        }

        if ($maxCol -le 0) { throw "Excelにデータがありません: $path" }
        $out = @()
        foreach ($h in $lines) {
            $arr = New-Object 'string[]' $maxCol
            for ($i = 0; $i -lt $maxCol; $i++) { $arr[$i] = '' }
            foreach ($k in $h.Keys) { if ($k -le $maxCol) { $arr[$k - 1] = [string]$h[$k] } }
            $out += ,$arr
        }
        return ,$out
    }
    finally { if ($zip) { $zip.Dispose() } }
}
