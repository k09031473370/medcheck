import zipfile, os, shutil, re, sys
src, vba, out = sys.argv[1], sys.argv[2], sys.argv[3]
work = os.path.join(os.path.dirname(out), 'panel_unz')
if os.path.exists(work): shutil.rmtree(work)
with zipfile.ZipFile(src) as z: names = z.namelist(); z.extractall(work)
shutil.copy(vba, f'{work}/xl/vbaProject.bin')
p = f'{work}/xl/worksheets/sheet2.xml'; s = open(p, encoding='utf-8').read()
# zoom back to normal
s, n = re.subn(r' zoomScale="140" zoomScaleNormal="140"', '', s); assert n == 1
# columns I..M (9-13) wide for the panel
old = '<col min="7" max="26" width="8.58203125" style="4"/>'
assert s.count(old) == 1
s = s.replace(old, '<col min="7" max="8" width="8.58203125" style="4"/><col min="9" max="13" width="21" style="4" customWidth="1"/><col min="14" max="26" width="8.58203125" style="4"/>')
# rows 34..200 tall enough for 16pt panel text
extra = ''.join(f'<row r="{r}" ht="28" customHeight="1"/>' for r in range(34, 201))
assert '<row r="34"' not in s
s = s.replace('</sheetData>', extra + '</sheetData>')
s = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="B1:AB200"/>', s)
open(p, 'w', encoding='utf-8').write(s)
# move the 受付No. explanation text box under the basic-info block (B35:D50)
d = f'{work}/xl/drawings/drawing2.xml'; ds = open(d, encoding='utf-8').read()
old = ('<xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
       '<xdr:to><xdr:col>18</xdr:col><xdr:colOff>44450</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>69850</xdr:rowOff></xdr:to>')
assert ds.count(old) == 1
ds = ds.replace(old, '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>34</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                     '<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>50</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>')
open(d, 'w', encoding='utf-8').write(ds)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names: z.write(f'{work}/{n}', n)
print('written', out)
