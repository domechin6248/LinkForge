#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LinkForge v1.5.0 ─ Word ハイパーリンク自動挿入ツール
Mac/Windows 対応・tkinterdnd2によるドラッグ&ドロップ対応
複数リンク資料フォルダ・何階層でも対応
"""

import os
import sys
import shutil
import threading
import platform
from pathlib import Path
from copy import deepcopy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── 依存パッケージの確認 ──────────────────────────────────────────
try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("エラー",
        "python-docx がインストールされていません。\n\n"
        "ターミナルで以下を実行してください:\n\n"
        "【Mac】  python3.14 -m pip install python-docx\n"
        "【Win】  python -m pip install python-docx")
    sys.exit(1)

# ── ドラッグ&ドロップ ─────────────────────────────────────────────
_DND_BACKEND = None
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_BACKEND = "tkdnd"
except ImportError:
    DND_FILES = "DND_Files"

# ── 定数 ─────────────────────────────────────────────────────────
HYPERLINK_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/hyperlink"
)
APP_VERSION = "1.5.0"

LINK_EXTENSIONS = {
    ".pdf", ".docx", ".doc", ".xlsx", ".xls",
    ".pptx", ".ppt", ".csv", ".txt",
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff",
    ".zip", ".rtf", ".odt", ".ods", ".odp",
}

C = {
    "bg":       "#1A1A2E",
    "surface":  "#16213E",
    "accent":   "#0F3460",
    "primary":  "#E94560",
    "text":     "#EAEAEA",
    "sub":      "#8899AA",
    "ok":       "#00D9A5",
    "warn":     "#FFC857",
    "err":      "#FF6B6B",
    "info":     "#48C9F7",
    "input_bg": "#0D1B2A",
    "border":   "#2A3A5E",
    "drop_hi":  "#1E3A5F",
}


# ════════════════════════════════════════════════════════════════
#  リンク処理コア
# ════════════════════════════════════════════════════════════════

import re as _re
import unicodedata as _ud

def _nfc(s: str) -> str:
    """Unicode NFC正規化（macOSのNFDファイル名をWordのNFCに統一）"""
    return _ud.normalize("NFC", s)

def _strip_number(name: str) -> str:
    """ファイル名先頭の数字・記号を除去する"""
    name = name.translate(str.maketrans('０１２３４５６７８９', '0123456789'))
    name = _re.sub(r'^[\d\s\.\-_、。．・()（）【】\[\]「」『』〔〕]+', '', name)
    return name.strip()

def get_file_map(link_dirs):
    """
    複数のリンク資料フォルダを再帰的に走査してファイルマップを生成。
    macOSのNFDファイル名をNFCに正規化してWordの文字列と統一する。
    先頭数字を除いたキーでもマッチできるようにする。
    """
    fm = {}
    for link_dir in link_dirs:
        base = Path(link_dir)
        for f in sorted(base.rglob("*")):
            if f.is_file() and f.suffix.lower() in LINK_EXTENSIONS:
                rel_str = str(f.relative_to(base.parent)).replace("\\", "/")
                # NFC正規化したステム・ファイル名で登録
                stem_nfc = _nfc(f.stem)
                name_nfc = _nfc(f.name)
                fm[stem_nfc] = rel_str
                fm[name_nfc] = rel_str
                # 先頭数字を除いたキーも登録
                stripped = _strip_number(stem_nfc)
                if stripped:
                    fm[stripped] = rel_str
                    fm[stripped + f.suffix] = rel_str
    return fm

def _nfc_find_matches(full: str, fm: dict) -> list:
    """NFC正規化してからマッチングする"""
    return _find_matches(_nfc(full), fm)

def copy_link_trees(link_dirs, dst_parent):
    """複数のリンク資料フォルダを出力先にコピー（元のフォルダ名のまま）"""
    for link_dir in link_dirs:
        src = Path(link_dir)
        dst = Path(dst_parent) / src.name
        if dst.exists():
            shutil.rmtree(str(dst))
        shutil.copytree(str(src), str(dst))

def iter_all_paragraphs(doc):
    def _tables(tables):
        for t in tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from cell.paragraphs
                    yield from _tables(cell.tables)
    yield from doc.paragraphs
    yield from _tables(doc.tables)

def _make_run(text, rpr=None):
    r = OxmlElement("w:r")
    if rpr is not None: r.append(deepcopy(rpr))
    t = OxmlElement("w:t")
    t.text = text
    if text != text.strip(): t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r

def _make_hyperlink(rid, text, rpr=None):
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), rid)
    hl.set(qn("w:history"), "1")
    r = OxmlElement("w:r")
    rp = OxmlElement("w:rPr")
    rs = OxmlElement("w:rStyle"); rs.set(qn("w:val"), "Hyperlink"); rp.append(rs)
    co = OxmlElement("w:color");  co.set(qn("w:val"), "0563C1");    rp.append(co)
    u  = OxmlElement("w:u");      u.set(qn("w:val"), "single");     rp.append(u)
    skip = {"rStyle", "color", "u"}
    if rpr is not None:
        for c in rpr:
            loc = c.tag.split("}")[-1] if "}" in c.tag else c.tag
            if loc not in skip: rp.append(deepcopy(c))
    r.append(rp)
    t = OxmlElement("w:t"); t.text = text
    if text != text.strip(): t.set(qn("xml:space"), "preserve")
    r.append(t); hl.append(r)
    return hl

def _find_matches(full, fm):
    keys = sorted(fm.keys(), key=len, reverse=True)
    matches, used = [], set()
    for k in keys:
        pos = 0
        while True:
            idx = full.find(k, pos)
            if idx == -1: break
            end = idx + len(k)
            span = set(range(idx, end))
            if not span & used:
                matches.append((idx, end, fm[k]))
                used |= span
            pos = idx + 1
    return sorted(matches, key=lambda x: x[0])

def process_paragraph(para, fm, part):
    p = para._p
    runs = []
    pos = 0
    for ch in list(p):
        if ch.tag == qn("w:r"):
            te = ch.find(qn("w:t"))
            tx = (te.text or "") if te is not None else ""
            runs.append(dict(elem=ch, text=tx, start=pos,
                             end=pos+len(tx), rpr=ch.find(qn("w:rPr"))))
            pos += len(tx)
    if not runs: return 0
    full = "".join(r["text"] for r in runs)
    ms = _nfc_find_matches(full, fm)
    if not ms: return 0
    crpr = [None] * len(full)
    for rd in runs:
        for i in range(rd["start"], rd["end"]): crpr[i] = rd["rpr"]
    for rd in runs: p.remove(rd["elem"])
    ppr = p.find(qn("w:pPr"))
    ins = (list(p).index(ppr)+1) if ppr is not None else 0
    elems, prev = [], 0
    for (s, e, rp) in ms:
        if prev < s:
            elems.append(_make_run(full[prev:s], crpr[prev] if prev < len(crpr) else None))
        rid = part.relate_to(rp, HYPERLINK_TYPE, is_external=True)
        elems.append(_make_hyperlink(rid, full[s:e], crpr[s] if s < len(crpr) else None))
        prev = e
    if prev < len(full):
        elems.append(_make_run(full[prev:], crpr[prev] if prev < len(crpr) else None))
    for i, el in enumerate(elems): p.insert(ins+i, el)
    return len(ms)


# ════════════════════════════════════════════════════════════════
#  DropZone ウィジェット
# ════════════════════════════════════════════════════════════════

class DropZone(tk.Frame):

    def __init__(self, parent, label_text, hint_text,
                 select_mode="file", file_types=None,
                 allow_multiple=False, **kwargs):
        super().__init__(parent, bg=C["surface"], **kwargs)
        self.select_mode = select_mode
        self.file_types = file_types or []
        self.allow_multiple = allow_multiple
        self.selected_paths = []
        self.on_change = None

        self.configure(
            highlightbackground=C["border"],
            highlightthickness=2,
            highlightcolor=C["primary"]
        )

        self.title_lbl = tk.Label(
            self, text="▸ " + label_text,
            font=("Helvetica", 12, "bold"),
            bg=C["surface"], fg=C["text"]
        )
        self.title_lbl.pack(pady=(12, 2), padx=14, anchor="w")

        self.hint_lbl = tk.Label(
            self, text=hint_text,
            font=("Helvetica", 9),
            bg=C["surface"], fg=C["sub"]
        )
        self.hint_lbl.pack(padx=14, anchor="w")

        pf = tk.Frame(self, bg=C["surface"])
        pf.pack(fill=tk.X, padx=14, pady=(6, 2))

        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(
            pf, textvariable=self.path_var,
            font=("Helvetica", 9),
            bg=C["input_bg"], fg=C["sub"],
            insertbackground=C["primary"],
            selectbackground=C["primary"],
            relief=tk.FLAT, bd=1
        )
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.path_entry.insert(0, "ここにパスをペースト...")
        self.path_entry.bind("<FocusIn>",  self._entry_in)
        self.path_entry.bind("<FocusOut>", self._entry_out)
        self.path_entry.bind("<Return>",   self._entry_submit)

        tk.Button(
            pf, text="読込", command=self._entry_submit,
            font=("Helvetica", 9),
            bg=C["accent"], fg=C["text"],
            activebackground=C["primary"], activeforeground="white",
            relief=tk.FLAT, bd=0, padx=10, pady=3, cursor="hand2"
        ).pack(side=tk.LEFT)

        self.info_lbl = tk.Label(
            self, text="", font=("Helvetica", 9),
            bg=C["surface"], fg=C["info"],
            justify=tk.LEFT, anchor="w", wraplength=440
        )
        self.info_lbl.pack(padx=14, anchor="w", pady=(2, 0))

        bf = tk.Frame(self, bg=C["surface"])
        bf.pack(padx=14, pady=(6, 12), anchor="w")

        label = "ファイル選択" if select_mode == "file" else "フォルダ選択"
        self.sel_btn = tk.Button(
            bf, text=label, command=self._on_click,
            font=("Helvetica", 9),
            bg=C["accent"], fg=C["text"],
            activebackground=C["primary"], activeforeground="white",
            relief=tk.FLAT, bd=0, padx=12, pady=4, cursor="hand2"
        )
        self.sel_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.clr_btn = None

        for w in [self, self.title_lbl, self.hint_lbl]:
            w.bind("<Button-1>", lambda e: self._on_click())

        if _DND_BACKEND == "tkdnd":
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>",      self._on_drop)
                self.dnd_bind("<<DragEnter>>", self._on_enter)
                self.dnd_bind("<<DragLeave>>", self._on_leave)
            except Exception:
                pass

    def _on_enter(self, event):
        self.configure(highlightbackground=C["primary"], highlightthickness=3,
                       bg=C["drop_hi"])
        self.title_lbl.configure(bg=C["drop_hi"])
        self.hint_lbl.configure(bg=C["drop_hi"])
        self.info_lbl.configure(bg=C["drop_hi"])
        return event.action

    def _on_leave(self, event):
        bg = C["surface"]
        self.configure(bg=bg, highlightthickness=2,
                       highlightbackground=C["ok"] if self.selected_paths else C["border"])
        self.title_lbl.configure(bg=bg)
        self.hint_lbl.configure(bg=bg)
        self.info_lbl.configure(bg=bg)
        return event.action

    def _on_drop(self, event):
        self._on_leave(event)
        paths = self._parse_paths(event.data)
        if paths:
            self._set_paths(paths)
        return event.action

    @staticmethod
    def _parse_paths(raw):
        paths = []
        i = 0
        while i < len(raw):
            if raw[i] == '{':
                end = raw.index('}', i)
                paths.append(raw[i+1:end])
                i = end + 2
            elif raw[i] == ' ':
                i += 1
            else:
                end = raw.find(' ', i)
                if end == -1: end = len(raw)
                paths.append(raw[i:end])
                i = end + 1
        return [p for p in paths if p.strip()]

    def _entry_in(self, e=None):
        if self.path_var.get() == "ここにパスをペースト...":
            self.path_entry.delete(0, tk.END)
            self.path_entry.configure(fg=C["text"])

    def _entry_out(self, e=None):
        if not self.path_var.get().strip():
            self.path_entry.configure(fg=C["sub"])
            self.path_entry.insert(0, "ここにパスをペースト...")

    def _entry_submit(self, e=None):
        raw = self.path_var.get().strip()
        if not raw or raw == "ここにパスをペースト...": return
        path = raw.strip("'\"").strip()
        if os.path.exists(path):
            self._set_paths([path])
            self.path_entry.delete(0, tk.END)
            self._entry_out()
        else:
            self.path_entry.configure(fg=C["err"])

    def _on_click(self):
        if self.select_mode == "folder":
            p = filedialog.askdirectory(title="フォルダを選択")
            if p: self._set_paths([p])
        elif self.allow_multiple:
            ps = filedialog.askopenfilenames(
                title="ファイルを選択",
                filetypes=self.file_types or [("All", "*.*")])
            if ps: self._set_paths(list(ps))
        else:
            p = filedialog.askopenfilename(
                title="ファイルを選択",
                filetypes=self.file_types or [("All", "*.*")])
            if p: self._set_paths([p])

    def _set_paths(self, paths):
        if self.select_mode == "folder" and not self.allow_multiple:
            self.selected_paths = [paths[0]]
        elif self.allow_multiple:
            ex = set(self.selected_paths)
            for p in paths:
                if p not in ex:
                    self.selected_paths.append(p); ex.add(p)
        else:
            self.selected_paths = [paths[0]]
        self._update_display()
        if self.on_change: self.on_change()

    def _update_display(self):
        if not self.selected_paths:
            self.info_lbl.config(text="")
            self.configure(highlightbackground=C["border"])
            if self.clr_btn: self.clr_btn.destroy(); self.clr_btn = None
            return

        self.configure(highlightbackground=C["ok"], highlightthickness=2)

        names = [Path(p).name for p in self.selected_paths]
        if len(names) == 1 and self.select_mode == "folder":
            folder = Path(self.selected_paths[0])
            files = [f for f in folder.rglob("*")
                     if f.is_file() and f.suffix.lower() in LINK_EXTENSIONS]
            text = f"✓  {folder.name}/  （{len(files)} ファイル）"
        elif len(names) <= 4:
            text = "\n".join(f"✓  {n}" for n in names)
        else:
            text = "\n".join(f"✓  {n}" for n in names[:3])
            text += f"\n  … 他 {len(names)-3} 件"

        self.info_lbl.config(text=text, fg=C["ok"])

        if not self.clr_btn:
            self.clr_btn = tk.Button(
                self.sel_btn.master, text="クリア", command=self._clear,
                font=("Helvetica", 8),
                bg="#3A1010", fg=C["err"],
                activebackground=C["err"], activeforeground="white",
                relief=tk.FLAT, bd=0, padx=8, pady=3, cursor="hand2"
            )
            self.clr_btn.pack(side=tk.LEFT)

    def _clear(self):
        self.selected_paths = []
        self.info_lbl.config(fg=C["info"])
        self._update_display()
        if self.on_change: self.on_change()


# ════════════════════════════════════════════════════════════════
#  メインアプリ
# ════════════════════════════════════════════════════════════════

class LinkForgeApp:

    def __init__(self):
        if _DND_BACKEND == "tkdnd":
            try:
                self.root = TkinterDnD.Tk()
            except Exception:
                self.root = tk.Tk()
        else:
            self.root = tk.Tk()

        self.root.title(f"LinkForge v{APP_VERSION}")
        self.root.geometry("620x920")
        self.root.minsize(520, 780)
        self.root.configure(bg=C["bg"])

        if platform.system() == "Darwin":
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(100, lambda: self.root.attributes("-topmost", False))
            try:
                self.root.createcommand("::tk::mac::OpenDocument", self._on_mac_open)
            except Exception:
                pass

        self.folder_name_entries = []
        self._build_ui()

        if _DND_BACKEND == "tkdnd":
            self._log("ドラッグ＆ドロップ: 有効 ✓", "success")
            self.dnd_lbl.configure(text="D&D ✓", fg=C["ok"])
        else:
            self._log("ドラッグ＆ドロップ: ダイアログ / パスペーストを利用", "info")

    def _build_ui(self):
        hdr = tk.Frame(self.root, bg=C["accent"], height=56)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)

        tk.Label(hdr, text="⛓  LinkForge",
                 font=("Helvetica", 18, "bold"),
                 bg=C["accent"], fg=C["primary"]
                 ).pack(side=tk.LEFT, padx=20, pady=10)
        tk.Label(hdr, text="Word ハイパーリンク自動挿入",
                 font=("Helvetica", 10),
                 bg=C["accent"], fg=C["sub"]
                 ).pack(side=tk.LEFT, pady=10)

        self.dnd_lbl = tk.Label(hdr, text="---",
                                 font=("Helvetica", 9),
                                 bg=C["accent"], fg=C["sub"])
        self.dnd_lbl.pack(side=tk.RIGHT, padx=6, pady=10)
        tk.Label(hdr, text=f"v{APP_VERSION}",
                 font=("Helvetica", 8),
                 bg=C["accent"], fg=C["sub"]
                 ).pack(side=tk.RIGHT, padx=(0, 2), pady=10)

        tk.Frame(self.root, bg=C["primary"], height=3).pack(fill=tk.X)

        outer = tk.Frame(self.root, bg=C["bg"])
        outer.pack(fill=tk.BOTH, expand=True)

        cv = tk.Canvas(outer, bg=C["bg"], highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        main = tk.Frame(cv, bg=C["bg"])
        fid = cv.create_window((0, 0), window=main, anchor="nw")
        main.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(fid, width=e.width))

        pad = dict(padx=16, pady=(0, 10))

        # ① 計画書（複数可）
        self.word_zone = DropZone(
            main,
            label_text="計画書（Word ファイル）",
            hint_text="ドラッグ＆ドロップ、またはクリックして選択（複数可）",
            select_mode="file",
            file_types=[("Word 文書", "*.docx")],
            allow_multiple=True
        )
        self.word_zone.pack(fill=tk.X, **pad)
        self.word_zone.on_change = self._on_word_changed

        # ② リンク資料フォルダ（複数可）
        self.link_zone = DropZone(
            main,
            label_text="リンク資料フォルダ",
            hint_text="ドラッグ＆ドロップで複数追加可能（何階層でも対応）",
            select_mode="folder",
            allow_multiple=True
        )
        self.link_zone.pack(fill=tk.X, **pad)
        self.link_zone.on_change = self._check_ready

        # ③ 出力先フォルダ
        self.output_zone = DropZone(
            main,
            label_text="出力先フォルダ",
            hint_text="ドラッグ＆ドロップ、またはクリックして選択",
            select_mode="folder",
            allow_multiple=False
        )
        self.output_zone.pack(fill=tk.X, **pad)
        self.output_zone.on_change = self._check_ready

        # ④ 出力フォルダ名
        self.fname_outer = tk.Frame(main, bg=C["bg"])

        # 出力構成プレビュー
        self.structure_lbl = tk.Label(
            main, text="",
            font=("Helvetica", 9),
            bg=C["bg"], fg=C["sub"],
            justify=tk.LEFT, anchor="w"
        )
        self.structure_lbl.pack(padx=16, anchor="w")

        tk.Frame(main, bg=C["border"], height=1).pack(fill=tk.X, padx=16, pady=(8, 10))

        # 実行ボタン
        self.run_btn = tk.Button(
            main,
            text="▶  リンクを作成",
            command=self._run,
            font=("Helvetica", 14, "bold"),
            bg=C["accent"], fg=C["sub"],
            activebackground=C["primary"], activeforeground="white",
            relief=tk.FLAT, bd=0, padx=30, pady=12,
            cursor="arrow", state=tk.DISABLED
        )
        self.run_btn.pack(pady=(0, 10))

        # ログ
        log_wrap = tk.Frame(main, bg=C["accent"], padx=1, pady=1)
        log_wrap.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))
        tk.Label(log_wrap, text=" 処理ログ",
                 font=("Helvetica", 9),
                 bg=C["accent"], fg=C["sub"]
                 ).pack(anchor="w")
        log_inner = tk.Frame(log_wrap, bg=C["input_bg"])
        log_inner.pack(fill=tk.BOTH, expand=True)
        mono = "Menlo" if platform.system() == "Darwin" else "Consolas"
        self.log_text = tk.Text(
            log_inner, height=8, font=(mono, 9),
            bg=C["input_bg"], fg=C["sub"],
            insertbackground=C["primary"],
            relief=tk.FLAT, bd=0, padx=10, pady=8,
            state=tk.DISABLED, wrap=tk.WORD
        )
        lsb = ttk.Scrollbar(log_inner, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        lsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_configure("success", foreground=C["ok"])
        self.log_text.tag_configure("error",   foreground=C["err"])
        self.log_text.tag_configure("info",    foreground=C["info"])
        self.log_text.tag_configure("warning", foreground=C["warn"])

    def _log(self, msg, tag=""):
        self.log_text.configure(state=tk.NORMAL)
        if tag:
            self.log_text.insert(tk.END, msg + "\n", tag)
        else:
            self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _on_word_changed(self):
        self._rebuild_folder_names()
        self._check_ready()

    def _rebuild_folder_names(self):
        self.fname_outer.pack_forget()
        for w in self.fname_outer.winfo_children(): w.destroy()
        self.folder_name_entries = []

        word_paths = self.word_zone.selected_paths
        if not word_paths:
            self.structure_lbl.configure(text="")
            return

        tk.Label(self.fname_outer,
                 text="出力フォルダ名（変更可）",
                 font=("Helvetica", 10, "bold"),
                 bg=C["bg"], fg=C["sub"]
                 ).pack(anchor="w", padx=16, pady=(0, 4))

        for wp in word_paths:
            p = Path(wp)
            if p.name.startswith("~$"): continue
            row = tk.Frame(self.fname_outer, bg=C["bg"])
            row.pack(fill=tk.X, padx=16, pady=2)
            tk.Label(row, text=f"{p.name}  →",
                     font=("Helvetica", 9),
                     bg=C["bg"], fg=C["sub"],
                     anchor="e", width=30
                     ).pack(side=tk.LEFT, padx=(0, 6))
            var = tk.StringVar(value=p.stem)
            tk.Entry(row, textvariable=var,
                     font=("Helvetica", 10),
                     bg=C["input_bg"], fg=C["text"],
                     insertbackground=C["primary"],
                     relief=tk.FLAT, bd=1
                     ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
            def _reset(v=var, d=p.stem): v.set(d)
            tk.Button(row, text="戻す", command=_reset,
                      font=("Helvetica", 8),
                      bg=C["accent"], fg=C["sub"],
                      relief=tk.FLAT, bd=0, padx=6, pady=2, cursor="hand2"
                      ).pack(side=tk.LEFT)
            self.folder_name_entries.append((wp, var))

        self.fname_outer.pack(fill=tk.X, pady=(0, 4), before=self.structure_lbl)
        self._update_structure_label()

    def _update_structure_label(self):
        if not self.folder_name_entries: return
        lines = ["📁 出力フォルダ構成プレビュー:"]
        for wp, var in self.folder_name_entries:
            folder_name = var.get() or Path(wp).stem
            lines.append(f"  出力先/{folder_name}/")
            lines.append(f"    ├─ {Path(wp).name}")
            for lp in self.link_zone.selected_paths:
                lines.append(f"    ├─ {Path(lp).name}/")
        self.structure_lbl.configure(text="\n".join(lines))

    def _check_ready(self):
        self._update_structure_label()
        ok = (self.word_zone.selected_paths and
              self.link_zone.selected_paths and
              self.output_zone.selected_paths)
        if ok:
            self.run_btn.configure(state=tk.NORMAL, bg=C["primary"],
                                   fg="white", cursor="hand2")
        else:
            self.run_btn.configure(state=tk.DISABLED, bg=C["accent"],
                                   fg=C["sub"], cursor="arrow")

    def _run(self):
        self.run_btn.configure(state=tk.DISABLED,
                               text="⏳  処理中...", bg=C["accent"], fg=C["warn"])
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            link_dirs  = self.link_zone.selected_paths
            output_dir = Path(self.output_zone.selected_paths[0])
            fn_map     = {wp: var.get().strip() or Path(wp).stem
                          for wp, var in self.folder_name_entries}

            # 全リンク資料フォルダからファイルマップ生成
            fm = get_file_map(link_dirs)
            if not fm:
                self.root.after(0, lambda: self._log(
                    "[警告] リンク資料フォルダに対象ファイルがありません。", "warning"))
                self.root.after(0, self._reset_btn); return

            self.root.after(0, lambda: self._log("リンク対象ファイル:", "info"))
            seen = set()
            for v in fm.values():
                if v not in seen:
                    vv = v
                    self.root.after(0, lambda x=vv: self._log(f"  {x}"))
                    seen.add(v)
            self.root.after(0, lambda: self._log(""))

            entries = [(wp, fn) for wp, fn in fn_map.items()
                       if not Path(wp).name.startswith("~$")]
            self.root.after(0, lambda: self._log(
                f"{len(entries)} 件の Word を処理...", "info"))

            total = 0
            for wp, cname in entries:
                wpath = Path(wp)
                self.root.after(0, lambda n=wpath.name, f=cname:
                                self._log(f"  {n}  →  {f}/"))
                try:
                    doc = Document(wpath)
                    out = output_dir / cname
                    out.mkdir(parents=True, exist_ok=True)

                    # リンク資料フォルダを出力先にコピー
                    copy_link_trees(link_dirs, out)
                    n_dirs = len(link_dirs)
                    self.root.after(0, lambda n=n_dirs: self._log(
                        f"    資料フォルダ {n} 個をコピー完了"))

                    # ハイパーリンク挿入
                    cnt = sum(process_paragraph(p, fm, doc.part)
                              for p in iter_all_paragraphs(doc))

                    # 計画書を出力先に保存
                    doc.save(out / wpath.name)

                    if cnt > 0:
                        m = f"    {cnt} 箇所リンク挿入"
                        self.root.after(0, lambda x=m: self._log(x, "success"))
                    else:
                        self.root.after(0, lambda: self._log(
                            "    対象テキストなし", "warning"))
                    total += cnt
                except Exception as e:
                    er = str(e)
                    self.root.after(0, lambda x=er: self._log(f"    エラー: {x}", "error"))

            self.root.after(0, lambda: self._log(""))
            self.root.after(0, lambda: self._log(
                f"完了！  合計 {total} 箇所のリンクを挿入しました。", "success"))
            self.root.after(0, lambda: self._log(f"出力先: {output_dir}", "info"))
            self.root.after(0, lambda: messagebox.showinfo(
                "完了",
                f"処理が完了しました！\n\n合計 {total} 箇所のリンクを挿入\n出力先: {output_dir}"
            ))
        except Exception as e:
            self.root.after(0, lambda: self._log(f"\n予期しないエラー: {e}", "error"))
        finally:
            self.root.after(0, self._reset_btn)

    def _reset_btn(self):
        self.run_btn.configure(text="▶  リンクを作成")
        self._check_ready()

    def _on_mac_open(self, *args):
        paths = [str(a) for a in args if os.path.exists(str(a))]
        for p in paths:
            if p.lower().endswith(".docx"):
                self.word_zone._set_paths([p])
            elif os.path.isdir(p):
                self.link_zone._set_paths([p])

    def run(self):
        self.root.mainloop()


# ════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    LinkForgeApp().run()