# -*- coding: utf-8 -*-
# =============================================
# Author: John
# Date: 2026-09-03
# Description: One-click launcher that picks a raw PDIV workbook and builds the Excel and HTML reports
# =============================================
# --- Version History ---
# v1.0.0 (2026-09-03): initial release, Tk file picker plus CLI mode, auto-discovers newest generators,
#                      validates the source layout before generating anything
# =============================================

VERSION = "v1.0.0"

import contextlib
import importlib.util
import io
import os
import re
import subprocess
import sys
import traceback
from pathlib import Path

HERE = Path(__file__).resolve().parent

# 生成器脚本前缀 -> (输出类型, 中文基名, 英文基名)
GENERATORS = {
    "pdiv_nok_report": ("xlsx", "PDIV检测数据统计_A站B站", "PDIV_Inspection_Statistics_A_vs_B"),
    "pdiv_nok_html_report": ("html", "PDIV检测数据趋势", "PDIV_Reject_Rate_Trend"),
}


# ---------------- 生成器发现与加载 ----------------
def parse_version(name):
    """从 'pdiv_nok_report_v1.2.1.py' 里取出 (1, 2, 1)，取不到就返回 (0,)。"""
    m = re.search(r"_v(\d+(?:\.\d+)*)\.py$", name)
    return tuple(int(x) for x in m.group(1).split(".")) if m else (0,)


def find_generator(prefix, search_dir=HERE):
    """找同目录下版本号最高的生成器脚本；html 版前缀更长，注意别被短前缀抢走。"""
    best, best_ver = None, None
    for path in Path(search_dir).glob(f"{prefix}_v*.py"):
        # pdiv_nok_report 的通配也会匹配到 pdiv_nok_html_report，用完整前缀再确认一次
        if not path.name.startswith(prefix + "_v"):
            continue
        ver = parse_version(path.name)
        if best_ver is None or ver > best_ver:
            best, best_ver = path, ver
    return best


def load_module(path):
    """脚本名里带点（v1.2.1），不能直接 import，用 importlib 按路径加载。"""
    mod_name = "gen_" + re.sub(r"\W", "_", path.stem)
    spec = importlib.util.spec_from_file_location(mod_name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = module
    spec.loader.exec_module(module)
    return module


# ---------------- 输出文件命名 ----------------
def next_output(out_dir, base, ext):
    """在输出目录里找没被占用的最小版本号，返回 <base>_v<N>.<ext>，绝不覆盖已有文件。"""
    used = set()
    for path in Path(out_dir).glob(f"{base}_v*.{ext}"):
        m = re.search(rf"_v(\d+)\.{ext}$", path.name)
        if m:
            used.add(int(m.group(1)))
    n = 1
    while n in used:
        n += 1
    return Path(out_dir) / f"{base}_v{n}.{ext}"


# ---------------- 核心流程（无界面，GUI 和命令行共用） ----------------
def run_batch(src, out_dir, langs, formats, log=print):
    """按选定的语言和格式跑一遍生成器，返回生成的文件路径列表。"""
    src, out_dir = Path(src), Path(out_dir)
    if not src.is_file():
        raise FileNotFoundError(f"找不到源数据文件：{src}")
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        import openpyxl  # noqa: F401
    except ImportError:
        raise RuntimeError("缺少 openpyxl。请先在命令行运行：pip install openpyxl")

    # 先用 xlsx 生成器的解析函数读一遍做校验，通过了才动手生成
    checker = find_generator("pdiv_nok_report")
    if checker is None:
        raise FileNotFoundError("同目录下找不到生成器脚本 pdiv_nok_report_v*.py")
    n_days, d0, d1 = validate_rows(load_module(checker).read_raw(str(src)))
    log(f"校验通过：{n_days} 个生产日（{d0} ~ {d1}）\n")

    made = []
    for prefix, (kind, base_zh, base_en) in GENERATORS.items():
        if kind not in formats:
            continue
        script = find_generator(prefix)
        if script is None:
            raise FileNotFoundError(f"同目录下找不到生成器脚本 {prefix}_v*.py")
        module = load_module(script)
        log(f"生成器：{script.name}（{module.VERSION}）")

        for lang in langs:
            base = base_en if lang == "en" else base_zh
            out = next_output(out_dir, base, kind)
            # 生成器自己往 stdout 打进度，接过来转进日志框，图形界面下才看得到
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                rc = module.main(str(src), str(out), lang)
            for line in buf.getvalue().splitlines():
                if line.strip():
                    log("    " + line.strip())
            if rc != 0:
                raise RuntimeError(f"{script.name} 生成 {out.name} 失败（返回码 {rc}）")
            log(f"  ✓ {out.name}")
            made.append(out)
    return made


DATE_RE = re.compile(r"^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}$")

LAYOUT_HINT = (
    "源文件格式不对。期望的布局是：第 3 行表头，第 4 行起是数据，A 列日期（如 2026.09.02），"
    "B~E 列 A 站（产量、PDIV 1/2/3 次 NOK），F~I 列 B 站。\n"
    "常见原因：选成了本工具生成的报表，或者选错了工作簿。"
)


def validate_rows(rows):
    """挡住选错文件的情况：日期列要像日期，产量要有数，否则宁可报错也不出垃圾报表。"""
    if not rows:
        raise ValueError("这个工作簿第 4 行起没读到任何数据行。\n" + LAYOUT_HINT)
    dated = sum(1 for r in rows if DATE_RE.match(str(r[0]).strip()))
    if dated < max(1, int(len(rows) * 0.8)):
        raise ValueError(f"A 列里只有 {dated}/{len(rows)} 行像日期。\n" + LAYOUT_HINT)
    if sum(r[1] for r in rows) <= 0 or sum(r[5] for r in rows) <= 0:
        raise ValueError("A 站或 B 站的生产数量合计为 0。\n" + LAYOUT_HINT)
    return len(rows), rows[0][0], rows[-1][0]


def open_folder(path):
    """生成完直接把输出目录打开，省得再去找。"""
    try:
        if os.name == "nt":
            os.startfile(str(path))                        # noqa: S606  (Windows)
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


# ---------------- 命令行模式 ----------------
def main_cli(argv):
    args = [a for a in argv if not a.startswith("--")]
    langs = []
    if "--zh" in argv:
        langs.append("zh")
    if "--en" in argv:
        langs.append("en")
    if not langs:
        langs = ["zh"]
    formats = [f for f in ("xlsx", "html") if f"--no-{f}" not in argv]
    if not args:
        print("用法：python pdiv_report_gui_v{v}.py <源数据.xlsx> [输出目录] "
              "[--zh] [--en] [--no-xlsx] [--no-html]".format(v=VERSION.lstrip("v")))
        return 2
    src = Path(args[0])
    out_dir = Path(args[1]) if len(args) > 1 else src.parent
    try:
        made = run_batch(src, out_dir, langs, formats)
    except Exception as exc:
        print(f"\n生成失败：{exc}", file=sys.stderr)
        return 1
    print(f"\n完成，共 {len(made)} 个文件，输出目录：{out_dir}")
    return 0


# ---------------- 图形界面模式 ----------------
def main_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(f"PDIV 报表生成工具 {VERSION}")
    root.geometry("680x460")
    root.minsize(620, 420)

    src_var = tk.StringVar()
    out_var = tk.StringVar()
    zh_var, en_var = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
    xlsx_var, html_var = tk.BooleanVar(value=True), tk.BooleanVar(value=True)

    pad = {"padx": 10, "pady": 6}
    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)

    ttk.Label(frm, text="源数据文件").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=src_var).grid(row=0, column=1, sticky="ew", **pad)

    def pick_src():
        path = filedialog.askopenfilename(
            title="选择原始数据表", filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if path:
            src_var.set(path)
            if not out_var.get():
                out_var.set(str(Path(path).parent))

    ttk.Button(frm, text="选择…", command=pick_src, width=10).grid(row=0, column=2, **pad)

    ttk.Label(frm, text="输出目录").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=out_var).grid(row=1, column=1, sticky="ew", **pad)
    ttk.Button(frm, text="选择…", width=10,
               command=lambda: out_var.set(filedialog.askdirectory(title="选择输出目录") or out_var.get())
               ).grid(row=1, column=2, **pad)

    opt = ttk.LabelFrame(frm, text="生成选项", padding=8)
    opt.grid(row=2, column=0, columnspan=3, sticky="ew", **pad)
    ttk.Checkbutton(opt, text="中文版", variable=zh_var).grid(row=0, column=0, padx=8)
    ttk.Checkbutton(opt, text="English", variable=en_var).grid(row=0, column=1, padx=8)
    ttk.Separator(opt, orient="vertical").grid(row=0, column=2, sticky="ns", padx=14)
    ttk.Checkbutton(opt, text="Excel (.xlsx)", variable=xlsx_var).grid(row=0, column=3, padx=8)
    ttk.Checkbutton(opt, text="网页 (.html)", variable=html_var).grid(row=0, column=4, padx=8)

    log_box = tk.Text(frm, height=12, wrap="word", state="disabled",
                      font=("Consolas", 9) if os.name == "nt" else None)
    log_box.grid(row=4, column=0, columnspan=3, sticky="nsew", **pad)
    frm.rowconfigure(4, weight=1)

    def log(msg):
        log_box.configure(state="normal")
        log_box.insert("end", str(msg) + "\n")
        log_box.see("end")
        log_box.configure(state="disabled")
        root.update_idletasks()

    def go():
        langs = [l for l, v in (("zh", zh_var), ("en", en_var)) if v.get()]
        formats = [f for f, v in (("xlsx", xlsx_var), ("html", html_var)) if v.get()]
        if not src_var.get():
            messagebox.showwarning("还差一步", "请先选择源数据文件。")
            return
        if not langs or not formats:
            messagebox.showwarning("还差一步", "语言和输出格式各至少要选一项。")
            return
        btn.configure(state="disabled")
        log_box.configure(state="normal"); log_box.delete("1.0", "end"); log_box.configure(state="disabled")
        out_dir = out_var.get() or str(Path(src_var.get()).parent)
        try:
            log(f"源数据：{Path(src_var.get()).name}")
            log(f"输出到：{out_dir}\n")
            made = run_batch(src_var.get(), out_dir, langs, formats, log=log)
            log(f"\n完成，共 {len(made)} 个文件。")
            open_folder(out_dir)
        except Exception as exc:
            log(f"\n出错了：{exc}")
            log(traceback.format_exc())
            messagebox.showerror("生成失败", str(exc))
        finally:
            btn.configure(state="normal")

    btn = ttk.Button(frm, text="开始生成", command=go)
    btn.grid(row=3, column=0, columnspan=3, sticky="ew", **pad)

    log_box.configure(state="normal")
    log_box.insert("end", "使用说明\n"
                          "1. 选择原始数据表（表头在第 3 行，数据从第 4 行开始，A~I 列）\n"
                          "2. 选输出目录，默认与源文件同目录\n"
                          "3. 勾选要生成的语言和格式，点「开始生成」\n"
                          "文件名版本号自动递增，不会覆盖已有报表。\n")
    log_box.configure(state="disabled")
    root.mainloop()
    return 0


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(main_cli(sys.argv[1:]))
    try:
        sys.exit(main_gui())
    except ImportError:
        print("这台机器的 Python 没装 tkinter，无法打开图形界面。\n"
              "可以改用命令行：python {} <源数据.xlsx> [输出目录] --zh --en"
              .format(Path(__file__).name))
        sys.exit(1)
