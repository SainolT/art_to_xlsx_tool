# -*- coding: utf-8 -*-
"""美术资源配置工具 - 界面（tkinter）。逻辑见 core.py。"""
import tkinter as tk
from pathlib import Path
from tkinter import ttk, scrolledtext, messagebox, filedialog

from core import (
    BACKUP_DIR, FMT_NAME, FMT_PATH, FMT_NAMES,
    app_dir, load_config, save_config, get_config_dir, save_config_dir,
    scan_resources, exts_of, filter_index, scope_index, parse_patterns,
    load_table_sheets, plan_match, preview_changes, find_in_use,
    split_dup_groups, run_match, check_existence, pinyin,
)

ROOT_LABEL = "（资源库根目录）"


class ResolveDialog(tk.Toplevel):
    """冲突解决/手动指定窗口：先选资源，再选行，把资源指定给选中的行。
    self.resolutions = {行号: 资源相对路径}（含带入的已有指定）。"""

    def __init__(self, master, items, all_rels, resolutions, title="解决冲突 / 手动指定"):
        """items: [(行号, 名称, [候选相对路径...], 旧值)]，候选为空表示未匹配。
        resolutions: {行号: 资源相对路径} 已有指定。"""
        super().__init__(master)
        self.title(title)
        self.geometry("1060x560")
        self.resolutions = dict(resolutions)
        self._all = sorted(all_rels)
        self._items = items
        self._by_row = {r: (name, cands, old) for r, name, cands, old in items}

        ttk.Label(self, padding=(8, 8, 8, 0),
                  text="左侧搜索并选中资源（选中行后可只看该行的候选），右侧搜索并勾选要写入的行（可多选），点「← 指定给选中的行」。",
                  foreground="#888888").pack(anchor=tk.W)

        pane = ttk.Frame(self, padding=8)
        pane.pack(fill=tk.BOTH, expand=True)
        pane.columnconfigure(0, weight=1)
        pane.columnconfigure(1, weight=1)
        pane.rowconfigure(0, weight=1)

        # 左：资源
        f1 = ttk.LabelFrame(pane, text="1. 选资源", padding=4)
        f1.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.rel_filter = tk.StringVar()
        self.rel_filter.trace_add("write", lambda *a: self._refresh_rels())
        ttk.Entry(f1, textvariable=self.rel_filter).pack(fill=tk.X, pady=(0, 2))
        self.cand_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f1, text="只看选中行的候选", variable=self.cand_var,
                        command=self._refresh_rels).pack(anchor=tk.W)
        self.rel_list = tk.Listbox(f1, exportselection=False)
        self.rel_list.pack(fill=tk.BOTH, expand=True)

        # 右：行
        f2 = ttk.LabelFrame(pane, text="2. 选行（可多选；显示 行号 | 名称 | 旧值）", padding=4)
        f2.grid(row=0, column=1, sticky="nsew")
        self.row_filter = tk.StringVar()
        self.row_filter.trace_add("write", lambda *a: self._refresh_rows())
        ttk.Entry(f2, textvariable=self.row_filter).pack(fill=tk.X, pady=(0, 2))
        self.row_list = tk.Listbox(f2, selectmode=tk.EXTENDED, exportselection=False)
        self.row_list.pack(fill=tk.BOTH, expand=True)
        self.row_list.bind("<<ListboxSelect>>", lambda e: self._refresh_rels())

        bar = ttk.Frame(self, padding=8)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="← 指定给选中的行", command=self._assign).pack(side=tk.LEFT)
        ttk.Button(bar, text="清除选中行的指定", command=self._clear).pack(
            side=tk.LEFT, padx=8)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT, padx=8)
        ttk.Button(bar, text="完成", command=self.destroy).pack(side=tk.RIGHT)

        self._refresh_rows()
        self._refresh_rels()
        self.transient(master)

    def _selected_rows(self):
        return [self._shown_rows[i] for i in self.row_list.curselection()]

    def _row_label(self, r):
        name, cands, old = self._by_row[r]
        done = "✔ " if r in self.resolutions else ""
        state = f"冲突{len(cands)}候选" if cands else "未匹配"
        return f"{done}行{r} | {name} | {state} | 旧值 {old or '（空）'}"

    def _refresh_rows(self):
        kw = self.row_filter.get().strip().lower()
        self._shown_rows = []
        self.row_list.delete(0, tk.END)
        for r, name, cands, old in self._items:
            label = self._row_label(r)
            if not kw or kw in label.lower():
                self._shown_rows.append(r)
                self.row_list.insert(tk.END, label)

    def _rel_pool(self):
        """当前资源池：勾选「只看候选」且有选中行时，为选中行候选的并集；否则全部。"""
        if self.cand_var.get():
            rows = self._selected_rows()
            pool = sorted({c for r in rows for c in self._by_row[r][1]})
            if pool:
                return pool
        return self._all

    def _refresh_rels(self):
        kw = self.rel_filter.get().strip().lower()
        self.rel_list.delete(0, tk.END)
        for rel in self._rel_pool():
            if not kw or kw in rel.lower():
                self.rel_list.insert(tk.END, rel)

    def _assign(self):
        rows = self._selected_rows()
        sel = self.rel_list.curselection()
        if not rows or not sel:
            messagebox.showwarning("缺少选择", "请先选中资源和要写入的行", parent=self)
            return
        rel = self.rel_list.get(sel[0])
        for r in rows:
            self.resolutions[r] = rel
        self._refresh_rows()
        self._refresh_status()

    def _clear(self):
        for r in self._selected_rows():
            self.resolutions.pop(r, None)
        self._refresh_rows()
        self._refresh_status()

    def _refresh_status(self):
        done = sum(1 for r, _, _, _ in self._items if r in self.resolutions)
        self.status.config(text=f"已指定 {done} / {len(self._items)} 行")


class DupDialog(tk.Toplevel):
    """重复资源决定窗口：同一资源被多个不同名称命中时，选定只写入哪个名称（或都不写）。
    结果写入 self.decisions = {资源相对路径: 名称 或 ""}。"""

    def __init__(self, master, groups):
        """groups: [(资源相对路径, [名称...]), ...]"""
        super().__init__(master)
        self.title("解决重复资源")
        self.geometry("640x520")
        self.decisions = dict(getattr(master, "_dup_decisions", {}))
        self._vars = {}

        ttk.Label(self, padding=8,
                  text="同一资源被多个不同名称命中（疑似首字母碰撞），请选定每个资源只写入哪个名称的行；"
                       "选「都不写」则该资源相关的行全部跳过。未选择的组按「都不写」处理。",
                  foreground="#888888", wraplength=610, justify=tk.LEFT).pack(
            anchor=tk.W, fill=tk.X)

        # 底部按钮栏先 pack 到底部，避免被滚动区挤走
        bar = ttk.Frame(self, padding=8)
        bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT)
        ttk.Button(bar, text="完成", command=self._done).pack(side=tk.RIGHT)

        # 可滚动区域
        canvas = tk.Canvas(self, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        inner = ttk.Frame(canvas, padding=8)
        win = canvas.create_window((0, 0), window=inner, anchor=tk.NW)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        for rel, names in groups:
            g = ttk.LabelFrame(inner, text=rel, padding=4)
            g.pack(fill=tk.X, pady=3)
            var = tk.StringVar(value=self.decisions.get(rel) or "__skip__")
            self._vars[rel] = var
            for name in names:
                ttk.Radiobutton(g, text=f"只写入「{name}」的行", variable=var,
                                value=name).pack(anchor=tk.W)
            ttk.Radiobutton(g, text="都不写", variable=var,
                            value="__skip__").pack(anchor=tk.W)

        self._refresh_status()
        self.transient(master)

    def _done(self):
        for rel, var in self._vars.items():
            v = var.get()
            self.decisions[rel] = "" if v == "__skip__" else v
        self.destroy()

    def _refresh_status(self):
        n = sum(1 for v in self._vars.values() if v.get() != "__skip__")
        self.status.config(text=f"共 {len(self._vars)} 组重复资源，已选保留 {n} 组")


class ChangesDialog(tk.Toplevel):
    """全部改动窗口：列出所有待写入改动，可双击/按钮排除部分行再执行。
    self.excluded = {行号}，关闭后由调用方读取。"""

    def __init__(self, master, changes, excluded, auto_excluded=(), included=(),
                 occupy=None):
        """changes: [(行号, 名称, 旧值, 新值, 资源相对路径)]；excluded: {行号} 已排除的。
        auto_excluded: {行号} 新值被其他行占用而默认排除的；included: {行号} 已人工恢复的占用行。
        occupy: {行号: 占用方描述}。"""
        super().__init__(master)
        self.title("全部改动")
        self.geometry("1080x540")
        self.excluded = set(excluded)
        self.included = set(included)
        self._auto = set(auto_excluded)
        self._occupy = occupy or {}
        self._data = {r: (name, old, val, rel) for r, name, old, val, rel in changes}

        ttk.Label(self, padding=(8, 8, 8, 0),
                  text="双击（或选中后按空格）切换 写入/排除；「占用」行的资源已被其他名称的行使用，"
                       "默认排除，可手动恢复。",
                  foreground="#888888").pack(anchor=tk.W)

        cols = ("state", "row", "name", "old", "value", "rel", "note")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for cid, text, w in [("state", "写入", 56), ("row", "行号", 56),
                             ("name", "名称", 120), ("old", "旧值", 170),
                             ("value", "新值", 190), ("rel", "资源", 280),
                             ("note", "备注", 160)]:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=tk.W)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.tree.bind("<Double-Button-1>", lambda e: self._toggle_selected())
        self.tree.bind("<space>", lambda e: self._toggle_selected())

        for r, name, old, val, rel in changes:
            self.tree.insert("", "end", iid=str(r), values=self._values(r))

        bar = ttk.Frame(self, padding=8)
        bar.pack(fill=tk.X)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT)
        ttk.Button(bar, text="排除选中", command=lambda: self._set_selected(False)).pack(
            side=tk.LEFT, padx=8)
        ttk.Button(bar, text="恢复选中", command=lambda: self._set_selected(True)).pack(
            side=tk.LEFT)
        ttk.Button(bar, text="全部恢复", command=self._restore_all).pack(side=tk.LEFT, padx=8)
        ttk.Button(bar, text="完成", command=self.destroy).pack(side=tk.RIGHT)
        self._refresh_status()
        self.transient(master)

    def _checked(self, r):
        """行当前是否勾选写入：人工排除 > 占用默认排除（可人工恢复）。"""
        return r not in self.excluded and (r not in self._auto or r in self.included)

    def _values(self, r):
        name, old, val, rel = self._data[r]
        if self._checked(r):
            state = "✓"
        elif r in self._auto:
            state = "—占用"
        else:
            state = "—"
        return (state, r, name, old, val, rel, self._occupy.get(r, ""))

    def _set_row(self, r, write):
        """勾选/取消勾选一行（处理占用行的恢复逻辑）。"""
        if write:
            if r in self._auto:
                self.included.add(r)
            else:
                self.excluded.discard(r)
        else:
            self.included.discard(r)
            if r not in self._auto:
                self.excluded.add(r)

    def _toggle_selected(self):
        for iid in self.tree.selection():
            r = int(iid)
            self._set_row(r, not self._checked(r))
            self.tree.item(iid, values=self._values(r))
        self._refresh_status()

    def _set_selected(self, write):
        for iid in self.tree.selection():
            r = int(iid)
            self._set_row(r, write)
            self.tree.item(iid, values=self._values(r))
        self._refresh_status()

    def _restore_all(self):
        self.excluded.clear()
        self.included = set(self._auto)
        for iid in self.tree.get_children():
            self.tree.item(iid, values=self._values(int(iid)))
        self._refresh_status()

    def _refresh_status(self):
        n_write = sum(1 for r in self._data if self._checked(r))
        self.status.config(
            text=f"共 {len(self._data)} 处改动，写入 {n_write} 行，"
                 f"排除 {len(self._data) - n_write} 行（含占用默认排除 {len(self._auto - self.included)} 行）")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("美术资源配置工具")
        self.geometry("1060x720")
        self.app_dir = app_dir()
        self.res_root = get_config_dir(self.app_dir, "res_root") \
            or get_config_dir(self.app_dir, "icon_root")  # 兼容旧配置
        self.backup_dir = get_config_dir(self.app_dir, "backup_dir") \
            or (self.app_dir / BACKUP_DIR)
        self._index = {}        # 资源库索引 {规范化词干: [相对路径...]}
        self._resolutions = {}  # {(文件路径, 子表, 资源字段, 行号): 资源相对路径}（会话内保持）
        self._dup_decisions = {}  # {资源相对路径: 保留的名称 或 ""}（会话内保持）
        self._excluded = set()  # {(文件路径, 子表, 资源字段, 行号)}（会话内保持）
        self._included = set()  # 人工恢复的占用行，同上维度（会话内保持）
        self._start_dirs = []   # 路径起始点候选（资源库根目录及其各层父目录）
        # 两个页签各自独立的配置表选择状态
        for prefix in ("m", "c"):
            setattr(self, f"{prefix}_file", None)    # 已选配置表路径
            setattr(self, f"{prefix}_sheets", [])    # [(子表名, [表头...])]
        self._build()
        self._rescan(log=False)
        self._log("[启动] 美术资源配置工具已就绪"
                  + ("（未安装 pypinyin，中文名只按原名匹配）" if not pinyin else ""))

    # ---------- 界面搭建 ----------

    def _build(self):
        # 顶部路径设置栏
        bar = ttk.LabelFrame(self, text="目录设置", padding=6)
        bar.pack(fill=tk.X, padx=6, pady=(6, 0))
        self._dir_row(bar, 0, "资源库根目录:", "res_root",
                      lambda: self._pick_dir("res_root"))
        ttk.Button(bar, text="重新扫描", command=lambda: self._rescan(log=True)).grid(
            row=1, column=1, sticky=tk.W, pady=(4, 0))
        self.count_var = tk.StringVar(value="资源库未扫描")
        ttk.Label(bar, textvariable=self.count_var, foreground="#555555").grid(
            row=1, column=2, sticky=tk.W, padx=8)
        ttk.Button(bar, text="设置备份目录", command=self._set_backup_dir).grid(
            row=1, column=3, sticky=tk.E, padx=4)

        ttk.Label(bar, text="资源类型（多选，默认全部）:").grid(
            row=0, column=4, sticky=tk.W, padx=(16, 4))
        self.ext_list = tk.Listbox(bar, selectmode=tk.EXTENDED, height=3,
                                   width=14, exportselection=False)
        self.ext_list.grid(row=1, column=4, sticky=tk.NW, padx=(16, 4), pady=(4, 0))
        ttk.Button(bar, text="全选", command=self._select_all_exts).grid(
            row=1, column=5, sticky=tk.NW, pady=(4, 0))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._build_match_tab()
        self._build_check_tab()

        self.log_area = scrolledtext.ScrolledText(self, height=16, font=("Consolas", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    def _dir_row(self, parent, row, label, attr, cmd):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W)
        var = tk.StringVar(value=str(getattr(self, attr) or "（未设置）"))
        setattr(self, f"{attr}_var", var)
        ttk.Label(parent, textvariable=var, width=70,
                  foreground="#333333").grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Button(parent, text="选择…", command=cmd).grid(row=row, column=3, sticky=tk.E)

    def _table_row(self, parent, prefix):
        """配置表选择行：直接弹文件框选一个 xlsx。"""
        ttk.Label(parent, text="配置表:").grid(row=0, column=0, sticky=tk.W)
        var = tk.StringVar(value="（未选择）")
        setattr(self, f"{prefix}_path_var", var)
        ttk.Label(parent, textvariable=var, width=50,
                  foreground="#333333").grid(row=0, column=1, columnspan=3, sticky=tk.W, padx=4)
        ttk.Button(parent, text="选择配置表…",
                   command=lambda: self._pick_table(prefix)).grid(
            row=0, column=4, columnspan=2, sticky=tk.W)

    def _base_row(self, parent, row, prefix):
        """路径起始点下拉（资源库根目录及其父目录，仅 相对路径+文件名 时生效）。"""
        ttk.Label(parent, text="路径起始点:").grid(row=row, column=0, sticky=tk.W)
        var = tk.StringVar(value=ROOT_LABEL)
        box = ttk.Combobox(parent, textvariable=var, state="readonly", width=34)
        box.grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(parent, text="（写相对路径时从哪一级目录开始算，可选资源库的父目录）",
                  foreground="#888888").grid(row=row, column=3, columnspan=3, sticky=tk.W)
        setattr(self, f"{prefix}_base_var", var)
        setattr(self, f"{prefix}_base_box", box)

    def _selected_prefix(self, prefix):
        """路径起始点 → 写入值的路径前缀（起始点到资源库根目录的相对路径）。"""
        pos = getattr(self, f"{prefix}_base_box").current()
        if pos <= 0 or not self.res_root:  # 0=资源库根目录，无前缀
            return ""
        start = self._start_dirs[pos]
        try:
            return self.res_root.relative_to(start).as_posix()
        except ValueError:
            return ""

    def _refresh_base_boxes(self):
        """路径起始点候选：资源库根目录 + 各层父目录（到盘符为止）。"""
        self._start_dirs = []
        if self.res_root:
            d = self.res_root.resolve()
            self._start_dirs.append(d)
            self._start_dirs.extend(d.parents)
        vals = [ROOT_LABEL] + [str(d) for d in self._start_dirs[1:]]
        for prefix in ("m", "c"):
            box = getattr(self, f"{prefix}_base_box", None)
            var = getattr(self, f"{prefix}_base_var", None)
            if box is None:
                continue
            old = var.get()
            box["values"] = vals
            var.set(old if old in vals else ROOT_LABEL)

    def _build_match_tab(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  批量匹配  ")

        sel = ttk.LabelFrame(tab, text="① 选择配置表与字段", padding=6)
        sel.pack(fill=tk.X, padx=8, pady=(6, 0))

        self._table_row(sel, "m")

        ttk.Label(sel, text="子表:").grid(row=1, column=0, sticky=tk.W)
        self.m_sheet_var = tk.StringVar()
        self.m_sheet_box = ttk.Combobox(sel, textvariable=self.m_sheet_var,
                                        state="readonly", width=24)
        self.m_sheet_box.grid(row=1, column=1, sticky=tk.W, padx=4)
        self.m_sheet_box.bind("<<ComboboxSelected>>", lambda e: self._m_on_sheet())

        ttk.Label(sel, text="匹配字段（名称）:").grid(row=1, column=2, sticky=tk.W)
        self.m_name_var = tk.StringVar()
        self.m_name_box = ttk.Combobox(sel, textvariable=self.m_name_var,
                                       state="readonly", width=18)
        self.m_name_box.grid(row=1, column=3, sticky=tk.W, padx=4)

        ttk.Label(sel, text="要填的字段（资源）:").grid(row=1, column=4, sticky=tk.W)
        self.m_value_var = tk.StringVar()
        self.m_value_box = ttk.Combobox(sel, textvariable=self.m_value_var,
                                        state="readonly", width=18)
        self.m_value_box.grid(row=1, column=5, sticky=tk.W, padx=4)

        opt = ttk.LabelFrame(tab, text="② 匹配设置", padding=6)
        opt.pack(fill=tk.X, padx=8, pady=4)

        ttk.Label(opt, text="写入格式:").grid(row=0, column=0, sticky=tk.W)
        self.m_fmt_var = tk.StringVar(value=FMT_NAME)
        ttk.Radiobutton(opt, text="文件名", variable=self.m_fmt_var,
                        value=FMT_NAME).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(opt, text="相对路径+文件名", variable=self.m_fmt_var,
                        value=FMT_PATH).grid(row=0, column=2, sticky=tk.W)
        self.m_keep_ext = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="保留扩展名", variable=self.m_keep_ext).grid(
            row=0, column=3, sticky=tk.W, padx=8)

        self._base_row(opt, 1, "m")

        cfg = load_config(self.app_dir)
        ttk.Label(opt, text="限定前缀:").grid(row=2, column=0, sticky=tk.W)
        self.m_prefix_var = tk.StringVar(value=cfg.get("scope_prefix", ""))
        ttk.Entry(opt, textvariable=self.m_prefix_var, width=34).grid(
            row=2, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(opt, text="限定后缀:").grid(row=2, column=3, sticky=tk.W)
        self.m_suffix_var = tk.StringVar(value=cfg.get("scope_suffix", ""))
        ttk.Entry(opt, textvariable=self.m_suffix_var, width=24).grid(
            row=2, column=4, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(opt, text="（多个用逗号分隔；只有文件名符合限定前后缀的资源才参与匹配，如 V1_icon_zw_）",
                  foreground="#888888").grid(row=3, column=0, columnspan=6, sticky=tk.W)

        act = ttk.Frame(tab, padding=6)
        act.pack(fill=tk.X)
        ttk.Button(act, text="预览", command=lambda: self._run_match(dry=True)).pack(side=tk.LEFT)
        ttk.Button(act, text="解决冲突", command=lambda: self._open_resolve("conflict")).pack(
            side=tk.LEFT, padx=(0, 8))
        ttk.Button(act, text="手动指定", command=lambda: self._open_resolve("missing")).pack(
            side=tk.LEFT)
        ttk.Button(act, text="解决重复资源", command=self._open_dups).pack(side=tk.LEFT, padx=8)
        ttk.Button(act, text="全部改动…", command=self._open_changes).pack(
            side=tk.LEFT, padx=8)
        ttk.Button(act, text="开始写入", command=lambda: self._run_match(dry=False)).pack(
            side=tk.RIGHT)

    def _build_check_tab(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  存在性校验  ")

        sel = ttk.LabelFrame(tab, text="① 选择配置表与资源字段（字段可多选）", padding=6)
        sel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(6, 0))

        self._table_row(sel, "c")

        ttk.Label(sel, text="子表:").grid(row=1, column=0, sticky=tk.W)
        self.c_sheet_var = tk.StringVar()
        self.c_sheet_box = ttk.Combobox(sel, textvariable=self.c_sheet_var,
                                        state="readonly", width=24)
        self.c_sheet_box.grid(row=1, column=1, sticky=tk.W, padx=4)
        self.c_sheet_box.bind("<<ComboboxSelected>>", lambda e: self._c_on_sheet())

        ttk.Label(sel, text="资源字段:").grid(row=1, column=2, sticky=tk.NW)
        self.c_field_list = tk.Listbox(sel, selectmode=tk.EXTENDED, height=6,
                                       exportselection=False)
        self.c_field_list.grid(row=1, column=3, sticky=tk.NSEW, padx=4)

        opt = ttk.Frame(tab, padding=6)
        opt.pack(fill=tk.X)
        ttk.Label(opt, text="② 校验方式:").grid(row=0, column=0, sticky=tk.W)
        self.c_fmt_var = tk.StringVar(value=FMT_NAME)
        ttk.Radiobutton(opt, text="按文件名", variable=self.c_fmt_var,
                        value=FMT_NAME).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(opt, text="按相对路径", variable=self.c_fmt_var,
                        value=FMT_PATH).grid(row=0, column=2, sticky=tk.W)
        self._base_row(opt, 1, "c")
        ttk.Button(opt, text="开始校验（只读，不写表）",
                   command=self._run_check).grid(row=0, column=3, sticky=tk.E, padx=8)

    # ---------- 数据装载 ----------

    def _pick_dir(self, attr):
        d = filedialog.askdirectory(parent=self, title="选择目录",
                                    initialdir=str(getattr(self, attr) or self.app_dir))
        if not d:
            return
        setattr(self, attr, Path(d))
        getattr(self, f"{attr}_var").set(d)
        save_config_dir(self.app_dir, attr, d)
        if attr == "res_root":
            self._rescan(log=True)

    def _set_backup_dir(self):
        d = filedialog.askdirectory(parent=self, initialdir=str(self.backup_dir),
                                    title="选择备份目录")
        if not d:
            return
        self.backup_dir = Path(d)
        save_config_dir(self.app_dir, "backup_dir", d)
        self._log(f"[备份目录] 已设置为: {self.backup_dir}")

    def _rescan(self, log=True):
        self._index = scan_resources(self.res_root, cache_dir=self.app_dir) \
            if self.res_root else {}
        self._refresh_base_boxes()
        # 扩展名列表：默认全选（= 不过滤）
        self.ext_list.delete(0, tk.END)
        for e in exts_of(self._index):
            self.ext_list.insert(tk.END, e or "（无扩展名）")
        self._select_all_exts()
        n = sum(len(v) for v in self._index.values())
        self.count_var.set(f"资源库 {n} 个文件")
        if log:
            self._log(f"[扫描] 资源库 {n} 个文件")

    def _select_all_exts(self):
        self.ext_list.select_set(0, tk.END)

    def _cur_index(self):
        """按勾选的资源类型（扩展名）过滤后的资源索引。全选/未选 = 不过滤。"""
        total = self.ext_list.size()
        sel = self.ext_list.curselection()
        if not sel or len(sel) == total:
            return self._index
        exts = set()
        for i in sel:
            e = self.ext_list.get(i)
            exts.add("" if e == "（无扩展名）" else e)
        return filter_index(self._index, exts)

    # ---------- 配置表选择 ----------

    def _pick_table(self, prefix):
        """直接弹文件框选一个配置表 xlsx，加载其子表。"""
        init = get_config_dir(self.app_dir, "table_dir") or self.app_dir
        f = filedialog.askopenfilename(
            parent=self, title="选择配置表", initialdir=str(init),
            filetypes=[("Excel 表格", "*.xlsx")])
        if not f:
            return
        p = Path(f)
        try:
            sheets = load_table_sheets(p)
        except Exception as ex:
            messagebox.showerror("读取失败", f"{p.name}: {ex}")
            return
        if not sheets:
            messagebox.showwarning("没有可用子表", f"{p.name} 里没有表头合规的子表")
            return
        save_config_dir(self.app_dir, "table_dir", p.parent)  # 记住目录，下次直接打开这里
        setattr(self, f"{prefix}_file", p)
        setattr(self, f"{prefix}_sheets", sheets)
        getattr(self, f"{prefix}_path_var").set(str(p))
        sheet_box = getattr(self, f"{prefix}_sheet_box")
        sheet_box["values"] = [s for s, _ in sheets]
        getattr(self, f"{prefix}_sheet_var").set("")
        if prefix == "m":
            self.m_name_box["values"] = []
            self.m_value_box["values"] = []
        else:
            self.c_field_list.delete(0, tk.END)
        self._log(f"[配置表] 已选择 {p}（{len(sheets)} 个可用子表）")

    def _cur_sheet_hdr(self, prefix):
        """当前选中的 (子表名, 表头)；未选返回 None。"""
        sheets = getattr(self, f"{prefix}_sheets")
        pos = getattr(self, f"{prefix}_sheet_box").current()
        return sheets[pos] if 0 <= pos < len(sheets) else None

    # ---------- 页签1 事件 ----------

    def _m_on_sheet(self):
        cur = self._cur_sheet_hdr("m")
        if not cur:
            return
        self.m_name_box["values"] = cur[1]
        self.m_value_box["values"] = cur[1]

    def _m_selection(self):
        """校验并返回 (文件路径, 子表, 名称字段, 资源字段, 写入格式, 前缀, 后缀, 路径前缀)，不完整返回 None。"""
        if not self.m_file:
            messagebox.showwarning("缺少选择", "请先选择配置表")
            return None
        cur = self._cur_sheet_hdr("m")
        if not cur:
            messagebox.showwarning("缺少选择", "请先选择子表")
            return None
        name_col, value_col = self.m_name_var.get(), self.m_value_var.get()
        if not name_col or not value_col:
            messagebox.showwarning("缺少选择", "请选择匹配字段和要填的字段")
            return None
        if name_col == value_col:
            messagebox.showwarning("选择有误", "匹配字段和要填的字段不能相同")
            return None
        if not self.res_root:
            messagebox.showwarning("缺少设置", "请先在顶部设置资源库根目录")
            return None
        fmt = self.m_fmt_var.get()
        scope_prefixes = parse_patterns(self.m_prefix_var.get())
        scope_suffixes = parse_patterns(self.m_suffix_var.get())
        path_prefix = self._selected_prefix("m")
        save_config(self.app_dir,
                    scope_prefix=self.m_prefix_var.get(),
                    scope_suffix=self.m_suffix_var.get())
        return (self.m_file, cur[0], name_col, value_col,
                fmt, scope_prefixes, scope_suffixes, path_prefix)

    def _run_match(self, dry):
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, prefixes, suffixes, path_prefix = sel
        exclude_rows = self._excluded_rows(path, sheet, value_col)
        include_rows = self._included_rows(path, sheet, value_col)
        res_rows = self._res_rows(path, sheet, value_col)
        if not dry:
            prefix_desc = f"\n路径起始点: {self.m_base_var.get()}" if fmt == FMT_PATH else ""
            if not messagebox.askyesno(
                "确认写入",
                f"文件: {path.name}\n子表: {sheet}\n"
                f"匹配字段「{name_col}」→ 要填的字段「{value_col}」\n"
                f"写入格式: {FMT_NAMES[fmt]}{prefix_desc}\n"
                f"已手动指定 {len(res_rows)} 行，"
                f"已决定重复资源 {len(self._dup_decisions)} 组，"
                f"已排除 {len(exclude_rows)} 行\n\n"
                "只会改写数据区，表头不动。确定执行？"):
                return
        self._log("=" * 60)
        try:
            run_match(path, sheet, name_col, value_col, self._cur_index(), fmt,
                      prefixes, suffixes, self.m_keep_ext.get(), path_prefix,
                      resolutions=res_rows, dup_decisions=self._dup_decisions,
                      exclude_rows=exclude_rows, include_rows=include_rows,
                      backup_dir=self.backup_dir, dry=dry, log=self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")

    def _open_resolve(self, mode):
        """mode='conflict' 只列冲突行；mode='missing' 只列未匹配行。"""
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, prefixes, suffixes, path_prefix = sel
        cur_index = self._cur_index()
        try:
            plan = plan_match(path, sheet, name_col, value_col, cur_index, fmt,
                              prefixes, suffixes, self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("读取失败", str(ex))
            return
        if mode == "conflict":
            items = [(r, name, cands, plan["old_values"].get(r, ""))
                     for r, name, cands in plan["conflict"]]
            title, empty_msg = "解决冲突", "当前没有冲突的行"
        else:
            items = [(r, name, [], plan["old_values"].get(r, ""))
                     for r, name in plan["missing"]]
            title, empty_msg = "手动指定", "当前没有未匹配的行"
        if not items:
            messagebox.showinfo("无需指定", empty_msg)
            return
        all_rels = sorted(rel for rels in
                          scope_index(cur_index, prefixes, suffixes).values()
                          for rel in rels)
        res_rows = self._res_rows(path, sheet, value_col)
        dlg = ResolveDialog(self, sorted(items), all_rels, res_rows, title=title)
        self.wait_window(dlg)
        self._save_res_rows(path, sheet, value_col, dlg.resolutions)
        self._log(f"[{title}] 当前共 {len(dlg.resolutions)} 行已指定资源")
        self._run_match(dry=True)  # 关闭后自动刷新预览

    def _dup_groups(self, plan, res_rows, fmt, keep_ext, prefix):
        """从匹配计划算出待人工决定的重复资源组: {资源相对路径: [不同名称...]}。
        已被组外名称的行填入的资源交给占用检测，不出现在这里；
        已被组内某名称的行填入的，默认保留该名称（写入 _dup_decisions）。
        同一名称多行共享资源暂不拦截。"""
        changes = [(r, n, None, rel) for r, n, v, rel in plan["ok"]]
        for r, name, cands in plan["conflict"]:
            pick = res_rows.get(r)
            if pick:
                changes.append((r, name, None, pick))
        for r, name in plan["missing"]:
            pick = res_rows.get(r)
            if pick:
                changes.append((r, name, None, pick))
        groups, preselect, _ = split_dup_groups(changes, plan, fmt, keep_ext, prefix)
        for rel, name in preselect.items():
            self._dup_decisions.setdefault(rel, name)
        return groups

    def _open_dups(self):
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, scope_prefixes, scope_suffixes, path_prefix = sel
        try:
            plan = plan_match(path, sheet, name_col, value_col, self._cur_index(), fmt,
                              scope_prefixes, scope_suffixes,
                              self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("读取失败", str(ex))
            return
        groups = self._dup_groups(plan, self._res_rows(path, sheet, value_col),
                                  fmt, self.m_keep_ext.get(), path_prefix)
        if not groups:
            messagebox.showinfo("无重复资源", "当前没有同一资源被多个不同名称命中的情况")
            return
        dlg = DupDialog(self, sorted(groups.items()))
        self.wait_window(dlg)
        self._dup_decisions.update(dlg.decisions)
        self._log(f"[重复资源] 当前共 {len(self._dup_decisions)} 组已决定")
        self._run_match(dry=True)  # 关闭后自动刷新预览

    def _excl_key(self, path, sheet, value_col):
        return (str(path), sheet, value_col)

    def _excluded_rows(self, path, sheet, value_col):
        key = self._excl_key(path, sheet, value_col)
        return {r for f, s, c, r in self._excluded if (f, s, c) == key}

    def _included_rows(self, path, sheet, value_col):
        key = self._excl_key(path, sheet, value_col)
        return {r for f, s, c, r in self._included if (f, s, c) == key}

    def _res_rows(self, path, sheet, value_col):
        """当前 文件+子表+字段 的人工指定：{行号: 资源相对路径}。"""
        key = self._excl_key(path, sheet, value_col)
        return {r: rel for (f, s, c, r), rel in self._resolutions.items()
                if (f, s, c) == key}

    def _save_res_rows(self, path, sheet, value_col, res_rows):
        """把对话框的行级指定写回全局 _resolutions（覆盖该上下文的旧值）。"""
        key = self._excl_key(path, sheet, value_col)
        self._resolutions = {k: v for k, v in self._resolutions.items()
                             if k[:3] != key}
        self._resolutions.update({key + (r,): rel for r, rel in res_rows.items()})

    def _open_changes(self):
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, prefixes, suffixes, path_prefix = sel
        try:
            plan = plan_match(path, sheet, name_col, value_col, self._cur_index(), fmt,
                              prefixes, suffixes, self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("读取失败", str(ex))
            return
        try:
            changes = preview_changes(plan, self._res_rows(path, sheet, value_col),
                                      self._dup_decisions,
                                      fmt, self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("计算改动失败", str(ex))
            return
        if not changes:
            messagebox.showinfo("没有改动", "当前没有待写入的改动")
            return
        key = self._excl_key(path, sheet, value_col)
        excl = self._excluded_rows(path, sheet, value_col)
        try:
            in_use = find_in_use(path, sheet, name_col, value_col, changes)
        except Exception as ex:
            messagebox.showerror("占用检测失败", str(ex))
            return
        auto = {r for r, *_ in in_use}
        occupy = {r: "占用: " + "、".join(holders)
                  for r, _, _, _, holders in in_use}
        incl = self._included_rows(path, sheet, value_col)
        dlg = ChangesDialog(self, changes, excl, auto, incl, occupy)
        self.wait_window(dlg)
        self._excluded = {k for k in self._excluded if k[:3] != key}
        self._excluded |= {(key[0], key[1], key[2], r) for r in dlg.excluded}
        self._included = {k for k in self._included if k[:3] != key}
        self._included |= {(key[0], key[1], key[2], r) for r in dlg.included}
        n_write = sum(1 for r, *_ in changes if dlg._checked(r))
        self._log(f"[全部改动] 共 {len(changes)} 处改动，写入 {n_write} 行，"
                  f"排除 {len(changes) - n_write} 行")
        self._run_match(dry=True)  # 关闭后自动刷新预览

    # ---------- 页签2 事件 ----------

    def _c_on_sheet(self):
        cur = self._cur_sheet_hdr("c")
        if not cur:
            return
        self.c_field_list.delete(0, tk.END)
        for h in cur[1]:
            self.c_field_list.insert(tk.END, h)

    def _run_check(self):
        if not self.c_file:
            messagebox.showwarning("缺少选择", "请先选择配置表")
            return
        cur = self._cur_sheet_hdr("c")
        if not cur:
            messagebox.showwarning("缺少选择", "请先选择子表")
            return
        cols = [self.c_field_list.get(i) for i in self.c_field_list.curselection()]
        if not cols:
            messagebox.showwarning("缺少选择", "请勾选要校验的资源字段")
            return
        if not self.res_root:
            messagebox.showwarning("缺少设置", "请先在顶部设置资源库根目录")
            return
        self._log("=" * 60)
        try:
            check_existence(self.c_file, cur[0], cols,
                            self._cur_index(), self.c_fmt_var.get(),
                            prefix=self._selected_prefix("c"), log=self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")

    def _log(self, msg):
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.update_idletasks()
