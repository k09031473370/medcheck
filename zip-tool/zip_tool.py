"""パスワード付きフォルダ圧縮ツール (AES-256 ZIP)

使い方:
    python zip_tool.py          # GUI を起動
    python zip_tool.py compress <フォルダ> <出力.zip> [-p パスワード]
    python zip_tool.py extract  <入力.zip> <展開先>  [-p パスワード]

exe 版では「自己解凍形式」で保存でき、受け取った人は exe をダブルクリックして
パスワードを入れるだけで展開できる（Windows のみ・追加ソフト不要）。
"""
import argparse
import getpass
import os
import shutil
import sys
import tempfile
import threading

import pyzipper


SFX_MARK = b"KANTAN-SFX"  # 自己解凍 exe に埋め込んだ ZIP の目印（ZIP コメント）


def compress_folder(folder, zip_path, password=None, progress=None, comment=b""):
    """フォルダを ZIP に圧縮する。password があれば AES-256 で暗号化。"""
    folder = os.path.abspath(folder)
    base = os.path.dirname(folder)
    files = []
    for root, dirs, names in os.walk(folder):
        if not names and not dirs:
            files.append((root, True))  # 空フォルダも保持
        for n in names:
            files.append((os.path.join(root, n), False))

    if password:
        zf = pyzipper.AESZipFile(zip_path, "w", compression=pyzipper.ZIP_DEFLATED,
                                 encryption=pyzipper.WZ_AES)
        zf.setpassword(password.encode("utf-8"))
    else:
        zf = pyzipper.ZipFile(zip_path, "w", compression=pyzipper.ZIP_DEFLATED)

    out = os.path.abspath(zip_path)
    with zf:
        zf.comment = comment
        for i, (path, is_dir) in enumerate(files, 1):
            if os.path.abspath(path) == out:
                continue  # 出力先がフォルダ内にある場合は自分自身を含めない
            arcname = os.path.relpath(path, base)
            if is_dir:
                zf.writestr(arcname.replace(os.sep, "/") + "/", b"")
            else:
                zf.write(path, arcname)
            if progress:
                progress(i, len(files))
    return len(files)


def extract_zip(zip_path, dest, password=None, progress=None):
    """ZIP を展開する。パスワードが違う場合は RuntimeError。"""
    dest = os.path.abspath(dest)
    with pyzipper.AESZipFile(zip_path) as zf:
        if password:
            zf.setpassword(password.encode("utf-8"))
        members = zf.infolist()
        for i, m in enumerate(members, 1):
            target = os.path.abspath(os.path.join(dest, m.filename))
            if not (target == dest or target.startswith(dest + os.sep)):
                raise ValueError(f"不正なパスを含むため中止しました: {m.filename}")
            zf.extract(m, dest)
            if progress:
                progress(i, len(members))
    return len(members)


def make_sfx(folder, exe_path, password=None, progress=None, stub=None):
    """自己解凍 exe を作る。stub（このツールの exe）の後ろに ZIP を連結する。"""
    stub = stub or sys.executable
    fd, tmp = tempfile.mkstemp(suffix=".zip")
    os.close(fd)
    try:
        n = compress_folder(folder, tmp, password, progress, comment=SFX_MARK)
        with open(exe_path, "wb") as out:
            for src in (stub, tmp):
                with open(src, "rb") as f:
                    shutil.copyfileobj(f, out)
    finally:
        os.remove(tmp)
    return n


def sfx_payload_path():
    """自分自身が自己解凍 exe なら exe のパスを返す。"""
    if not getattr(sys, "frozen", False):
        return None
    try:
        with pyzipper.AESZipFile(sys.executable) as zf:
            return sys.executable if zf.comment == SFX_MARK else None
    except Exception:
        return None


def run_sfx(exe):
    """自己解凍: パスワードを聞いて exe と同じ場所に展開し、フォルダを開く。"""
    import tkinter as tk
    from tkinter import messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    dest = os.path.dirname(exe)
    with pyzipper.AESZipFile(exe) as zf:
        encrypted = any(m.flag_bits & 0x1 for m in zf.infolist())
        top = zf.namelist()[0].split("/")[0] if zf.namelist() else ""
    while True:
        pw = None
        if encrypted:
            pw = simpledialog.askstring("パスワード", "パスワードを入力してください:",
                                        show="●", parent=root)
            if pw is None:
                return  # キャンセル
        try:
            extract_zip(exe, dest, pw)
            break
        except RuntimeError:
            if not messagebox.askretrycancel("エラー", "パスワードが違います。"):
                return
        except Exception as e:
            return messagebox.showerror("エラー", str(e))
    opened = os.path.join(dest, top)
    messagebox.showinfo("完了", f"展開しました:\n{opened}")
    if hasattr(os, "startfile"):
        os.startfile(opened)


# ---------------- GUI ----------------

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("かんたん圧縮ツール")
    root.resizable(False, False)

    nb = ttk.Notebook(root)
    nb.pack(padx=10, pady=10, fill="both")

    def make_tab(title):
        f = ttk.Frame(nb, padding=10)
        nb.add(f, text=title)
        return f

    def row(parent, r, label, var, browse, show=None):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=var, width=45, show=show).grid(row=r, column=1, pady=3)
        if browse:
            ttk.Button(parent, text="参照...", command=browse).grid(row=r, column=2, padx=5)

    progress = ttk.Progressbar(root, length=400, mode="determinate")
    progress.pack(padx=10, fill="x")
    status = tk.StringVar(value="準備OK")
    ttk.Label(root, textvariable=status).pack(padx=10, pady=(2, 10), anchor="w")

    def set_progress(i, n):
        root.after(0, lambda: (progress.configure(maximum=n, value=i),
                               status.set(f"処理中... {i}/{n}")))

    def run_task(fn, done_msg, buttons):
        for b in buttons:
            b.configure(state="disabled")

        def worker():
            try:
                n = fn()
                msg = done_msg.format(n=n)
                root.after(0, lambda: (status.set(msg), messagebox.showinfo("完了", msg)))
            except RuntimeError:
                root.after(0, lambda: (status.set("失敗"),
                                       messagebox.showerror("エラー", "パスワードが違うか、未入力です。")))
            except Exception as e:
                err = str(e)
                root.after(0, lambda: (status.set("失敗"), messagebox.showerror("エラー", err)))
            finally:
                root.after(0, lambda: [b.configure(state="normal") for b in buttons])

        threading.Thread(target=worker, daemon=True).start()

    # --- 圧縮タブ ---
    c = make_tab("圧縮")
    c_src, c_out = tk.StringVar(), tk.StringVar()
    c_pw, c_pw2 = tk.StringVar(), tk.StringVar()

    def pick_src():
        d = filedialog.askdirectory(title="圧縮するフォルダ")
        if d:
            c_src.set(d)
            if not c_out.get():
                c_out.set(d.rstrip("/\\") + (".exe" if c_sfx.get() else ".zip"))

    def pick_out():
        p = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP", "*.zip")])
        if p:
            c_out.set(os.path.splitext(p)[0] + (".exe" if c_sfx.get() else ".zip"))

    row(c, 0, "フォルダ:", c_src, pick_src)
    row(c, 1, "保存先:", c_out, pick_out)
    row(c, 2, "パスワード:", c_pw, None, show="●")
    row(c, 3, "確認用:", c_pw2, None, show="●")
    ttk.Label(c, text="※パスワード空欄なら暗号化なし", foreground="gray").grid(
        row=4, column=1, sticky="w")

    c_sfx = tk.BooleanVar()
    frozen = getattr(sys, "frozen", False)

    def on_sfx():
        out = c_out.get()
        if out:
            stem = os.path.splitext(out)[0]
            c_out.set(stem + (".exe" if c_sfx.get() else ".zip"))

    ttk.Checkbutton(
        c, variable=c_sfx, command=on_sfx,
        text="自己解凍形式（相手はダブルクリック→パスワードで開ける）"
             + ("" if frozen else " ※exe版のみ"),
        state="normal" if frozen else "disabled",
    ).grid(row=5, column=1, sticky="w")

    def do_compress():
        if not os.path.isdir(c_src.get()):
            return messagebox.showwarning("確認", "フォルダを選んでください。")
        if not c_out.get():
            return messagebox.showwarning("確認", "保存先を指定してください。")
        if c_pw.get() != c_pw2.get():
            return messagebox.showwarning("確認", "パスワードが一致しません。")
        if os.path.exists(c_out.get()) and not messagebox.askyesno("確認", "上書きしますか？"):
            return
        src, out, pw = c_src.get(), c_out.get(), c_pw.get() or None
        fn = make_sfx if c_sfx.get() else compress_folder
        run_task(lambda: fn(src, out, pw, set_progress),
                 "圧縮しました（{n} 件）", [c_btn, e_btn])

    c_btn = ttk.Button(c, text="圧縮する", command=do_compress)
    c_btn.grid(row=6, column=1, pady=8)

    # --- 展開タブ ---
    e = make_tab("展開")
    e_src, e_out, e_pw = tk.StringVar(), tk.StringVar(), tk.StringVar()

    def pick_zip():
        p = filedialog.askopenfilename(filetypes=[("ZIP", "*.zip")])
        if p:
            e_src.set(p)
            if not e_out.get():
                e_out.set(os.path.dirname(p))

    def pick_dest():
        d = filedialog.askdirectory(title="展開先")
        if d:
            e_out.set(d)

    row(e, 0, "ZIPファイル:", e_src, pick_zip)
    row(e, 1, "展開先:", e_out, pick_dest)
    row(e, 2, "パスワード:", e_pw, None, show="●")

    def do_extract():
        if not os.path.isfile(e_src.get()):
            return messagebox.showwarning("確認", "ZIPファイルを選んでください。")
        if not e_out.get():
            return messagebox.showwarning("確認", "展開先を指定してください。")
        src, out, pw = e_src.get(), e_out.get(), e_pw.get() or None
        run_task(lambda: extract_zip(src, out, pw, set_progress),
                 "展開しました（{n} 件）", [c_btn, e_btn])

    e_btn = ttk.Button(e, text="展開する", command=do_extract)
    e_btn.grid(row=3, column=1, pady=8)

    root.mainloop()


# ---------------- CLI ----------------

def main(argv=None):
    if argv is None and len(sys.argv) == 1:
        exe = sfx_payload_path()
        if exe:
            return run_sfx(exe)

    p = argparse.ArgumentParser(description="パスワード付きフォルダ圧縮ツール")
    sub = p.add_subparsers(dest="cmd")
    for name, a, b in [("compress", "folder", "zip"), ("extract", "zip", "dest")]:
        s = sub.add_parser(name)
        s.add_argument(a)
        s.add_argument(b)
        s.add_argument("-p", "--password", help="省略時は入力を求めます（空 Enter で暗号化なし）")
        if name == "compress":
            s.add_argument("--sfx", action="store_true", help="自己解凍 exe で保存（exe 版のみ）")
    args = p.parse_args(argv)

    if args.cmd is None:
        return run_gui()

    pw = args.password
    if pw is None:
        pw = getpass.getpass("パスワード（なしなら Enter）: ")
    pw = pw or None
    try:
        if args.cmd == "compress" and args.sfx:
            if not getattr(sys, "frozen", False):
                sys.exit("エラー: 自己解凍形式は exe 版でのみ作れます。")
            n = make_sfx(args.folder, args.zip, pw)
        elif args.cmd == "compress":
            n = compress_folder(args.folder, args.zip, pw)
        else:
            n = extract_zip(args.zip, args.dest, pw)
    except RuntimeError:
        sys.exit("エラー: パスワードが違うか、未入力です。")
    print(f"完了: {n} 件")


if __name__ == "__main__":
    main()
