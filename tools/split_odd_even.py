#!/usr/bin/env python3
"""複数PDFを1ページずつのPDFに分割し、奇数ページと偶数ページを別フォルダに振り分ける。
使い方: python3 split_odd_even.py 入力1.pdf 入力2.pdf ... [-o 出力先ディレクトリ]
出力:
  出力先/odd/foo_p001.pdf, foo_p003.pdf, ...   (奇数ページ)
  出力先/even/foo_p002.pdf, foo_p004.pdf, ...  (偶数ページ)
"""
import argparse, pathlib
from pypdf import PdfReader, PdfWriter

def split(src: pathlib.Path, outdir: pathlib.Path) -> None:
    reader = PdfReader(str(src))
    n_odd = n_even = 0
    for i, page in enumerate(reader.pages, start=1):   # i は1始まりのページ番号
        sub = "odd" if i % 2 == 1 else "even"
        out = outdir / sub / f"{src.stem}_p{i:03d}.pdf"
        w = PdfWriter()
        w.add_page(page)
        w.write(str(out))
        if sub == "odd":
            n_odd += 1
        else:
            n_even += 1
    print(f"{src.name}: 全{len(reader.pages)}ページ -> odd {n_odd}件 / even {n_even}件")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pdfs", nargs="+", type=pathlib.Path)
    ap.add_argument("-o", "--outdir", type=pathlib.Path, default=pathlib.Path("."))
    a = ap.parse_args()
    (a.outdir / "odd").mkdir(parents=True, exist_ok=True)
    (a.outdir / "even").mkdir(parents=True, exist_ok=True)
    for p in a.pdfs:
        split(p, a.outdir)
