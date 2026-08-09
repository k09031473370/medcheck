# -*- coding: utf-8 -*-
"""
配布用 excel_tools.zip を作る。

Windowsのエクスプローラーは、UTF-8フラグ付きのZIPの日本語ファイル名を
正しく読めないことがある(文字化け)。日本語Windowsの既定である cp932 で
ファイル名を書き、UTF-8フラグを立てないようにする。
"""
import os
import sys
import zipfile

# --- ファイル名を cp932 で書く (UTF-8フラグを立てない) ---
_orig = zipfile.ZipInfo._encodeFilenameFlags


def _encode_cp932(self):
    try:
        return self.filename.encode('ascii'), self.flag_bits
    except UnicodeEncodeError:
        try:
            return self.filename.encode('cp932'), self.flag_bits
        except UnicodeEncodeError:
            return _orig(self)


zipfile.ZipInfo._encodeFilenameFlags = _encode_cp932

HERE = os.path.dirname(os.path.abspath(__file__))

ENTRIES = [
    'xlsx_read.ps1',
    'form_import.ps1',
    'form_import_gui.ps1',
    'ecg_import.ps1',
    'yoyaku_export.ps1',
    'db_tool.ps1',
    'restore_backup.ps1',
    'check.ps1', 'check.bat',
    'check2.ps1', 'check2.bat',
    'check3.ps1', 'check3.bat',
    'check4.ps1', 'check4.bat',
    'check5.ps1', 'check5.bat',
    'check6.ps1', 'check6.bat',
    'check7.ps1', 'check7.bat',
    'check8.ps1', 'check8.bat',
    'check9.ps1', 'check9.bat',
    'check10.ps1', 'check10.bat',
    'check11.ps1', 'check11.bat',
    'check12.ps1', 'check12.bat',
    'start.bat',
    '心電図取込.bat',
    '元に戻す.bat',
    'USBで使うときは.txt',
    'form',
    '受付番号リスト_ひな形.xlsx',
    '取込マニュアル.docx',
    '取込マニュアル.md',
]


def add(zf, rel):
    full = os.path.join(HERE, rel)
    if os.path.isdir(full):
        for name in sorted(os.listdir(full)):
            add(zf, os.path.join(rel, name))
        return
    if not os.path.exists(full):
        print('  [skip] %s (ありません)' % rel)
        return
    zf.write(full, rel.replace(os.sep, '/'))
    print('  %s' % rel)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'excel_tools.zip')
    if os.path.exists(out):
        os.remove(out)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zf:
        for e in ENTRIES:
            add(zf, e)
    print('作成: %s (%d bytes)' % (out, os.path.getsize(out)))


if __name__ == '__main__':
    main()
