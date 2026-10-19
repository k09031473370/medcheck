"""入力シート: hide rows 19-33 (便潜血/オプション/備考) and enlarge 基本情報 text."""
import re, os, sys, zipfile, shutil
src, out = sys.argv[1], sys.argv[2]
work = os.path.join(os.path.dirname(out), 'restyle_unz')
if os.path.exists(work): shutil.rmtree(work)
with zipfile.ZipFile(src) as z: names = z.namelist(); z.extractall(work)
rd = lambda p: open(os.path.join(work, p), encoding='utf-8').read()
wr = lambda p, s: open(os.path.join(work, p), 'w', encoding='utf-8').write(s)

# ---- styles: clone xf + font with bigger size ----
st = rd('xl/styles.xml')
fonts_m = re.search(r'<fonts count="(\d+)"([^>]*)>(.*?)</fonts>', st, flags=re.S)
font_list = re.findall(r'<font>.*?</font>|<font/>', fonts_m.group(3), flags=re.S)
xfs_m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', st, flags=re.S)
xf_list = re.findall(r'<xf [^>]*/>|<xf [^>]*>.*?</xf>', xfs_m.group(2), flags=re.S)
assert len(font_list) == int(fonts_m.group(1)) and len(xf_list) == int(xfs_m.group(1))
new_fonts, new_xfs, mapping = [], [], {}
def bigger(xf_idx, size):
    xf = xf_list[xf_idx]; fid = int(re.search(r'fontId="(\d+)"', xf).group(1))
    f = font_list[fid]
    f = re.sub(r'<sz val="[^"]*"/>', f'<sz val="{size}"/>', f) if '<sz ' in f else f.replace('<font>', f'<font><sz val="{size}"/>')
    new_fonts.append(f); new_fid = len(font_list) + len(new_fonts) - 1
    new_xfs.append(re.sub(r'fontId="\d+"', f'fontId="{new_fid}"', xf)); mapping[xf_idx] = len(xf_list) + len(new_xfs) - 1
for idx, size in ((25, 14), (26, 16), (33, 16), (34, 16), (81, 16), (36, 18)):
    bigger(idx, size)
st = st.replace(fonts_m.group(0), f'<fonts count="{len(font_list)+len(new_fonts)}"{fonts_m.group(2)}>{fonts_m.group(3)}{"".join(new_fonts)}</fonts>')
st = st.replace(xfs_m.group(0), f'<cellXfs count="{len(xf_list)+len(new_xfs)}">{xfs_m.group(2)}{"".join(new_xfs)}</cellXfs>')
wr('xl/styles.xml', st)

# ---- sheet2 ----
s = rd('xl/worksheets/sheet2.xml')
def restyle_row(m):
    r = int(m.group(1)); row = m.group(0)
    if 5 <= r <= 17:
        row = re.sub(r'(<c r="[BC]%d" s=")(\d+)(")' % r, lambda c: c.group(1) + str(mapping.get(int(c.group(2)), int(c.group(2)))) + c.group(3), row)
        ht = {5: 28}.get(r, 34 if r in (6, 7) else 28)
        if r != 8: row = re.sub(r'ht="[^"]*"', f'ht="{ht}"', row, count=1)
    if 19 <= r <= 33:
        row = re.sub(r'^<row ([^>]*?)(/?)>', lambda h: '<row ' + h.group(1) + ' hidden="1"' + h.group(2) + '>', row)
    return row
s, n = re.subn(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', restyle_row, s, flags=re.S)
assert s.count('hidden="1"') == 15, s.count('hidden="1"')
s = s.replace('<col min="2" max="2" width="19.33203125" style="4" customWidth="1"/>', '<col min="2" max="2" width="24" style="4" customWidth="1"/>')
s = s.replace('<col min="3" max="3" width="28.58203125" style="4" customWidth="1"/>', '<col min="3" max="3" width="42" style="4" customWidth="1"/>')
assert 'width="24"' in s and 'width="42"' in s
# blank display while 会社名/氏名 empty or master cell empty (instead of 0 / 1899/12/30)
def blank_guard(m):
    col = m.group(1); inner = f'INDEX(受診者マスタ[{col}],MATCH(会社名&amp;"_"&amp;氏名,受診者マスタ[検索キー],0))'
    return f'<f>IF(OR(会社名="",氏名=""),"",IFERROR(IF({inner}="","",{inner}),""))</f>'
s, k = re.subn(r'<f>IFERROR\(INDEX\(受診者マスタ\[([^\]]+)\],MATCH\(会社名&amp;"_"&amp;氏名,受診者マスタ\[検索キー\],0\)\),""\)</f>', blank_guard, s)
assert k == 9, k
wr('xl/worksheets/sheet2.xml', s)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names: z.write(os.path.join(work, n), n)
print('written', out, 'new xfs:', mapping)
