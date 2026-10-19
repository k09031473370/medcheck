import re, os, sys, zipfile, shutil
src = sys.argv[1]; out = sys.argv[2]
work = os.path.join(os.path.dirname(out), 'tpl_unz')
if os.path.exists(work): shutil.rmtree(work)
with zipfile.ZipFile(src) as z:
    names = z.namelist(); z.extractall(work)

def rd(p): return open(os.path.join(work,p),encoding='utf-8').read()
def wr(p,s): open(os.path.join(work,p),'w',encoding='utf-8').write(s)
def sub1(s, pat, rep, flags=0, count_expected=None):
    new, n = re.subn(pat, rep, s, flags=flags)
    if count_expected is not None: assert n == count_expected, (pat, n)
    return new

# 1. workbook.xml: drop 20251128 sheet, its sheet-scoped names, reset active tab
wb = rd('xl/workbook.xml')
wb = sub1(wb, r'<sheet name="20251128"[^>]*/>', '', count_expected=1)
wb = sub1(wb, r'<definedName [^>]*localSheetId="7"[^>]*>[^<]*</definedName>', '', count_expected=3)
wb = sub1(wb, r'firstSheet="2" activeTab="7"', 'firstSheet="0" activeTab="0"', count_expected=1)
wr('xl/workbook.xml', wb)

# 2. workbook rels: drop sheet8 + calcChain
rels = rd('xl/_rels/workbook.xml.rels')
rels = sub1(rels, r'<Relationship Id="rId8" [^>]*worksheets/sheet8.xml"/>', '', count_expected=1)
rels = sub1(rels, r'<Relationship Id="rId16" [^>]*calcChain.xml"/>', '', count_expected=1)
wr('xl/_rels/workbook.xml.rels', rels)

# 3. content types
ct = rd('[Content_Types].xml')
for part in ('/xl/worksheets/sheet8.xml','/xl/drawings/drawing4.xml','/xl/tables/table5.xml','/xl/calcChain.xml'):
    ct = sub1(ct, r'<Override PartName="%s"[^>]*/>' % re.escape(part), '', count_expected=1)
wr('[Content_Types].xml', ct)

# 4. delete orphaned parts
for p in ('xl/worksheets/sheet8.xml','xl/worksheets/_rels/sheet8.xml.rels','xl/drawings/drawing4.xml',
          'xl/drawings/_rels/drawing4.xml.rels','xl/tables/table5.xml','xl/calcChain.xml'):
    os.remove(os.path.join(work,p))

# 5. app.xml titles
app = rd('docProps/app.xml')
app = sub1(app, r'<vt:lpstr>20251128</vt:lpstr>', '', count_expected=1)
app = sub1(app, r"<vt:lpstr>'20251128'![^<]*</vt:lpstr>", '', count_expected=3)
app = sub1(app, r'<vt:i4>8</vt:i4>', '<vt:i4>7</vt:i4>', count_expected=1)
app = sub1(app, r'<vt:i4>27</vt:i4>', '<vt:i4>24</vt:i4>', count_expected=1)
app = sub1(app, r'<vt:vector size="35" baseType="lpstr">', '<vt:vector size="31" baseType="lpstr">', count_expected=1)
wr('docProps/app.xml', app)

# 6. 受診者マスタ (sheet3): clear data rows 4..142, keep C formula
s3 = rd('xl/worksheets/sheet3.xml')
def clear_row(m):
    row = m.group(0)
    def clear_cell(cm):
        c = cm.group(0)
        col = re.match(r'<c r="([A-Z]+)', c).group(1)
        if col == 'C':
            return re.sub(r'<v>[^<]*</v>', '<v>_</v>', c)
        attrs = re.match(r'<c ([^>]*?)/?>', c).group(1)
        attrs = re.sub(r'\s*t="[^"]*"', '', attrs)
        return '<c %s/>' % attrs
    return re.sub(r'<c r="[A-Z]+\d+"[^>]*?(?:/>|>.*?</c>)', clear_cell, row, flags=re.S)
s3, n = re.subn(r'<row r="(?:[4-9]|[1-9]\d|1[0-3]\d|14[0-2])"[^>]*>.*?</row>', clear_row, s3, flags=re.S)
assert n == 139, n
s3 = sub1(s3, r'<selection activeCell="O24" sqref="O24"/>', '<selection activeCell="D4" sqref="D4"/>', count_expected=1)
wr('xl/worksheets/sheet3.xml', s3)

# 7. 会社マスタ (sheet4): clear B4, B5
s4 = rd('xl/worksheets/sheet4.xml')
s4 = sub1(s4, r'<c r="B4" s="4" t="s"><v>134</v></c>', '<c r="B4" s="4"/>', count_expected=1)
s4 = sub1(s4, r'<c r="B5" s="4" t="s"><v>173</v></c>', '<c r="B5" s="4"/>', count_expected=1)
wr('xl/worksheets/sheet4.xml', s4)

# 8. 入力シート (sheet2): 受付No start = 1
s2 = rd('xl/worksheets/sheet2.xml')
s2 = sub1(s2, r'<c r="C3" s="24"><v>20</v></c>', '<c r="C3" s="24"><v>1</v></c>', count_expected=1)
wr('xl/worksheets/sheet2.xml', s2)

# 9. re-zip in original order
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names:
        p = os.path.join(work, n)
        if os.path.exists(p): z.write(p, n)
print("written", out)
