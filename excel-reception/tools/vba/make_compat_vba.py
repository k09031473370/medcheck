"""Final step: put the rebuilt vbaProject.bin into the compat template and restore the
受診者マスタ 年齢 / 正式コース名 formulas in rows 4-142."""
import re, os, sys, zipfile, shutil
src, vba_bin, out = sys.argv[1], sys.argv[2], sys.argv[3]
work = os.path.join(os.path.dirname(out), 'final_unz')
if os.path.exists(work): shutil.rmtree(work)
with zipfile.ZipFile(src) as z: names = z.namelist(); z.extractall(work)
shutil.copy(vba_bin, os.path.join(work, 'xl/vbaProject.bin'))
p = os.path.join(work, 'xl/worksheets/sheet3.xml'); s = open(p, encoding='utf-8').read()
AGE = 'IF(受診者マスタ[[#This Row],[生年月日]]="","",DATEDIF(受診者マスタ[[#This Row],[生年月日]],IF(MONTH(TODAY())&gt;=4,DATE(YEAR(TODAY())+1,3,31),DATE(YEAR(TODAY()),3,31)),"Y"))'
COURSE = 'IFERROR(INDEX(tbl_コース[正式コース名],MATCH(受診者マスタ[[#This Row],[検診コース]],tbl_コース[コース],0)),"")'
s, nK = re.subn(r'<c r="K(\d+)" s="4"/>', lambda m: f'<c r="K{m.group(1)}" s="4" t="str"><f>{COURSE}</f><v></v></c>' if int(m.group(1)) >= 4 else m.group(0), s)
s, nL = re.subn(r'<c r="L(\d+)" s="4"/>', lambda m: f'<c r="L{m.group(1)}" s="4" t="str"><f>{AGE}</f><v></v></c>' if int(m.group(1)) >= 4 else m.group(0), s)
assert s.count("<f>" + COURSE) == 139 and s.count("<f>" + AGE) == 139, (s.count("<f>" + COURSE), s.count("<f>" + AGE))
open(p, 'w', encoding='utf-8').write(s)
t = os.path.join(work, 'xl/tables/table1.xml'); ts = open(t, encoding='utf-8').read()
ts, n = re.subn(r'(<tableColumn [^>]*name="年齢"[^>]*>)<calculatedColumnFormula>[^<]*</calculatedColumnFormula>', lambda m: m.group(1) + f'<calculatedColumnFormula>{AGE}</calculatedColumnFormula>', ts); assert n == 1
open(t, 'w', encoding='utf-8').write(ts)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in names: z.write(os.path.join(work, n), n)
print('written', out)
