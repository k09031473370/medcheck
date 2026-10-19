"""Build an Excel 2016/2019 compatible variant of the blank template:
replace XLOOKUP/FILTER/SORTBY/LET/UNIQUE with INDEX/MATCH + INDIRECT."""
import re, os, sys, zipfile, shutil
src, out = sys.argv[1], sys.argv[2]
work = os.path.join(os.path.dirname(out), 'compat_unz')
if os.path.exists(work): shutil.rmtree(work)
with zipfile.ZipFile(src) as z:
    names = z.namelist(); z.extractall(work)
rd = lambda p: open(os.path.join(work, p), encoding='utf-8').read()
wr = lambda p, s: open(os.path.join(work, p), 'w', encoding='utf-8').write(s)
def subn(s, pat, rep, n, flags=0):
    new, k = re.subn(pat, rep, s, flags=flags); assert k == n, (pat[:60], k, n); return new

# --- 入力シート ---
s2 = rd('xl/worksheets/sheet2.xml')
s2 = subn(s2, r'_xlfn\.XLOOKUP\(会社名&amp;"_"&amp;氏名,受診者マスタ\[検索キー\],受診者マスタ\[([^\]]+)\],""\)',
          r'IFERROR(INDEX(受診者マスタ[\1],MATCH(会社名&amp;"_"&amp;氏名,受診者マスタ[検索キー],0)),"")', 9)
s2 = subn(s2, r'<formula1>_xlfn\.ANCHORARRAY\(氏名リスト元\)</formula1>', '<formula1>INDIRECT("受診者マスタ[氏名]")</formula1>', 1)
s2 = subn(s2, r'<c r="AA1" s="80" t="str" cm="1"><f t="array" ref="AA1">.*?</f><v/></c>', '<c r="AA1" s="80"/>', 1, re.S)
wr('xl/worksheets/sheet2.xml', s2)

# --- 【原本】受付名簿 + its table definition ---
pat = r'_xlfn\.XLOOKUP\((tblDay1\[\[#This Row\],\[会社名\]\]&amp;"_"&amp;tblDay1\[\[#This Row\],\[氏名\]\]),受診者マスタ\[検索キー\],受診者マスタ\[([^\]]+)\]\)'
rep = r'IFERROR(INDEX(受診者マスタ[\2],MATCH(\1,受診者マスタ[検索キー],0)),"")'
s7 = rd('xl/worksheets/sheet7.xml')
s7 = subn(s7, pat, rep, 70)
s7 = subn(s7, r'<c r="C28" s="1" t="e" cm="1" vm="1"><f t="array" ref="C28">.*?</f>.*?</c>', '<c r="C28" s="1"/>', 1, re.S)
s7 = subn(s7, r'<c r="D28" s="1" cm="1"><f t="array" ref="D28">.*?</f>.*?</c>', '<c r="D28" s="1"/>', 1, re.S)
for ref in ('C26', 'C27', 'D27'):   # フリー1集計 labels (feature not available without dynamic arrays)
    s7 = subn(s7, r'<c r="%s" (s="\d+") t="s"><v>\d+</v></c>' % ref, r'<c r="%s" \1/>' % ref, 1)
wr('xl/worksheets/sheet7.xml', s7)
t4 = rd('xl/tables/table4.xml')
t4 = subn(t4, pat, rep, 7)
wr('xl/tables/table4.xml', t4)

# --- 受診者マスタ table: 正式コース名 calculated column ---
t1 = rd('xl/tables/table1.xml')
t1 = subn(t1, r'_xlfn\.XLOOKUP\(受診者マスタ\[\[#This Row\],\[検診コース\]\],tbl_コース\[コース\],tbl_コース\[正式コース名\]\)',
          'IFERROR(INDEX(tbl_コース[正式コース名],MATCH(受診者マスタ[[#This Row],[検診コース]],tbl_コース[コース],0)),"")', 1)
wr('xl/tables/table1.xml', t1)

# --- force full recalculation on open ---
wb = rd('xl/workbook.xml')
wb = subn(wb, r'<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>', 1)
wr('xl/workbook.xml', wb)

with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names:
        p = os.path.join(work, n)
        if os.path.exists(p): z.write(p, n)
print('written', out)
