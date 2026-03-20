#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LinkForge ─ Word ハイパーリンク自動挿入ツール
Mac/Windows 対応・古いTkでも動くCanvas描画ベースUI
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
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "エラー",
        "python-docx がインストールされていません。\n\n"
        "ターミナルで以下を実行してください:\n\n"
        "【Mac】  python3 -m pip install python-docx\n"
        "【Win】  python -m pip install python-docx"
    )
    sys.exit(1)

# ── DnD バックエンド ──────────────────────────────────────────────
_DND_BACKEND = None
DND_FILES = "DND_Files"

# ── OOXML 定数 ────────────────────────────────────────────────────
HYPERLINK_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/hyperlink"
)

APP_VERSION = "1.2.0"

# ── カラーテーマ ──────────────────────────────────────────────────
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

# ── リンク対象拡張子 ──────────────────────────────────────────────
LINK_EXTENSIONS = {
    ".pdf", ".docx", ".doc", ".xlsx", ".xls",
    ".pptx", ".ppt", ".csv", ".txt",
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff",
    ".zip", ".rtf", ".odt", ".ods", ".odp",
}


# ════════════════════════════════════════════════════════════════
#  リンク処理コア（変更なし）
# ════════════════════════════════════════════════════════════════

def get_file_map(link_dir: Path) -> dict:
    file_map: dict = {}
    for f in sorted(link_dir.rglob("*")):
        if f.is_file() and f.suffix.lower() in LINK_EXTENSIONS:
            rel = str(f.relative_to(link_dir))
            file_map[f.stem] = rel
            file_map[f.name] = rel
    return file_map


def copy_link_tree(link_dir: Path, dest_dir: Path):
    for item in link_dir.iterdir():
        dst = dest_dir / item.name
        if item.is_dir():
            shutil.copytree(str(item), str(dst), dirs_exist_ok=True)
        else:
            shutil.copy2(str(item), str(dst))


def iter_all_paragraphs(doc):
    def _iter_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from cell.paragraphs
                    yield from _iter_tables(cell.tables)
    yield from doc.paragraphs
    yield from _iter_tables(doc.tables)


def _make_run(text: str, rpr_elem=None):
    r = OxmlElement("w:r")
    if rpr_elem is not None:
        r.append(deepcopy(rpr_elem))
    t = OxmlElement("w:t")
    t.text = text
    if text != text.strip():
        t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def _make_hyperlink(r_id: str, display_text: str, rpr_elem=None):
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    hl.set(qn("w:history"), "1")
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rpr.append(rStyle)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(u)
    skip_tags = {"rStyle", "color", "u"}
    if rpr_elem is not None:
        for child in rpr_elem:
            local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if local not in skip_tags:
                rpr.append(deepcopy(child))
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = display_text
    if display_text != display_text.strip():
        t.set(qn("xml:space"), "preserve")
    r.append(t)
    hl.append(r)
    return hl


def _find_matches(full_text: str, file_map: dict) -> list:
    keys_sorted = sorted(file_map.keys(), key=len, reverse=True)
    matches: list = []
    used: set = set()
    for key in keys_sorted:
        pos = 0
        while True:
            idx = full_text.find(key, pos)
            if idx == -1:
                break
            end = idx + len(key)
            span = set(range(idx, end))
            if not span & used:
                matches.append((idx, end, file_map[key]))
                used |= span
            pos = idx + 1
    matches.sort(key=lambda x: x[0])
    return matches


def process_paragraph(paragraph, file_map: dict, doc_part) -> int:
    p_elem = paragraph._p
    runs_data = []
    pos = 0
    for child in list(p_elem):
        if child.tag == qn("w:r"):
            t_elem = child.find(qn("w:t"))
            text = (t_elem.text or "") if t_elem is not None else ""
            rpr = child.find(qn("w:rPr"))
            runs_data.append(
                dict(elem=child, text=text, start=pos, end=pos + len(text), rpr=rpr)
            )
            pos += len(text)
    if not runs_data:
        return 0
    full_text = "".join(r["text"] for r in runs_data)
    matches = _find_matches(full_text, file_map)
    if not matches:
        return 0
    char_rpr: list = [None] * len(full_text)
    for rd in runs_data:
        for i in range(rd["start"], rd["end"]):
            char_rpr[i] = rd["rpr"]
    for rd in runs_data:
        p_elem.remove(rd["elem"])
    ppr = p_elem.find(qn("w:pPr"))
    insert_at = (list(p_elem).index(ppr) + 1) if ppr is not None else 0
    new_elems = []
    prev_end = 0
    for (start, end, rel_path) in matches:
        if prev_end < start:
            seg_text = full_text[prev_end:start]
            rpr = char_rpr[prev_end] if prev_end < len(char_rpr) else None
            new_elems.append(_make_run(seg_text, rpr))
        match_text = full_text[start:end]
        rpr = char_rpr[start] if start < len(char_rpr) else None
        r_id = doc_part.relate_to(rel_path, HYPERLINK_TYPE, is_external=True)
        new_elems.append(_make_hyperlink(r_id, match_text, rpr))
        prev_end = end
    if prev_end < len(full_text):
        seg_text = full_text[prev_end:]
        rpr = char_rpr[prev_end] if prev_end < len(char_rpr) else None
        new_elems.append(_make_run(seg_text, rpr))
    for i, elem in enumerate(new_elems):
        p_elem.insert(insert_at + i, elem)
    return len(matches)


# ════════════════════════════════════════════════════════════════
#  GUI ── Canvas描画ベース（古いTkでも確実に動く）
# ════════════════════════════════════════════════════════════════

def _btn(parent, text, command, bg, fg="#EAEAEA", width=None, font_size=9):
    """シンプルなボタン生成ヘルパー"""
    kw = dict(
        text=text, command=command,
        bg=bg, fg=fg, activebackground=C["primary"], activeforeground="white",
        relief=tk.FLAT, bd=0, padx=10, pady=4,
        font=("Helvetica", font_size), cursor="hand2"
    )
    if width:
        kw["width"] = width
    return tk.Button(parent, **kw)


class DropZone(tk.Frame):
    """ファイル/フォルダ選択エリア（Canvas枠線で装飾）"""

    def __init__(self, parent, label_text, hint_text,
                 select_mode="file", file_types=None,
                 allow_multiple=False, **kwargs):
        super().__init__(parent, bg=C["bg"], **kwargs)
        self.select_mode = select_mode
        self.file_types = file_types or []
        self.allow_multiple = allow_multiple
        self.selected_paths: list = []
        self.on_change = None
        self._border_color = C["border"]

        # Canvas で枠線を描画（古いTkでも確実に機能する）
        self.canvas = tk.Canvas(
            self, height=4, bg=C["bg"],
            highlightthickness=0, bd=0
        )
        self.canvas.pack(fill=tk.X)

        # 内部フレーム
        inner = tk.Frame(self, bg=C["surface"])
        inner.pack(fill=tk.X, padx=2, pady=0)

        # タイトル行
        title_row = tk.Frame(inner, bg=C["surface"])
        title_row.pack(fill=tk.X, padx=12, pady=(10, 2))

        # 左側：アイコン＋ラベル
        tk.Label(
            title_row, text="▸ " + label_text,
            font=("Helvetica", 12, "bold"),
            bg=C["surface"], fg=C["text"]
        ).pack(side=tk.LEFT)

        # ヒント
        self.hint_label = tk.Label(
            inner, text=hint_text,
            font=("Helvetica", 9),
            bg=C["surface"], fg=C["sub"]
        )
        self.hint_label.pack(anchor="w", padx=14)

        # パス入力欄
        path_row = tk.Frame(inner, bg=C["surface"])
        path_row.pack(fill=tk.X, padx=12, pady=(4, 2))

        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(
            path_row, textvariable=self.path_var,
            font=("Helvetica", 9),
            bg=C["input_bg"], fg=C["sub"],
            insertbackground=C["primary"],
            selectbackground=C["primary"],
            relief=tk.FLAT, bd=1
        )
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.path_entry.insert(0, "ここにパスをペースト...")
        self.path_entry.bind("<FocusIn>", self._entry_focus_in)
        self.path_entry.bind("<FocusOut>", self._entry_focus_out)
        self.path_entry.bind("<Return>", self._entry_submit)

        _btn(path_row, "読込", self._entry_submit, C["accent"]).pack(side=tk.LEFT)

        # 選択ファイル表示
        self.info_label = tk.Label(
            inner, text="", font=("Helvetica", 9),
            bg=C["surface"], fg=C["info"],
            justify=tk.LEFT, anchor="w", wraplength=440
        )
        self.info_label.pack(anchor="w", padx=14, pady=(2, 0))

        # ボタン行
        btn_row = tk.Frame(inner, bg=C["surface"])
        btn_row.pack(anchor="w", padx=12, pady=(4, 10))

        label = "ファイル選択" if select_mode == "file" else "フォルダ選択"
        self.sel_btn = _btn(btn_row, label, self._on_click, C["accent"])
        self.sel_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.clr_btn = None

        # 下線（Canvas）
        self.canvas_bottom = tk.Canvas(
            self, height=2, bg=C["bg"],
            highlightthickness=0, bd=0
        )
        self.canvas_bottom.pack(fill=tk.X)

        self._draw_border(C["border"])

        # クリックでファイル選択
        for w in [inner, self.hint_label, self.info_label]:
            w.bind("<Button-1>", lambda e: self._on_click())

    def _draw_border(self, color):
        self._border_color = color
        for cv in [self.canvas, self.canvas_bottom]:
            cv.configure(bg=color)

    # ── パス入力欄 ──
    def _entry_focus_in(self, e=None):
        if self.path_var.get() == "ここにパスをペースト...":
            self.path_entry.delete(0, tk.END)
            self.path_entry.configure(fg=C["text"])

    def _entry_focus_out(self, e=None):
        if not self.path_var.get().strip():
            self.path_entry.configure(fg=C["sub"])
            self.path_entry.insert(0, "ここにパスをペースト...")

    def _entry_submit(self, e=None):
        raw = self.path_var.get().strip()
        if not raw or raw == "ここにパスをペースト...":
            return
        path = raw.strip("'\"").strip()
        if os.path.exists(path):
            self._set_paths([path])
            self.path_entry.delete(0, tk.END)
            self._entry_focus_out()
        else:
            self.path_entry.configure(fg=C["err"])

    def _on_click(self):
        if self.select_mode == "folder":
            p = filedialog.askdirectory(title="フォルダを選択")
            if p:
                self._set_paths([p])
        else:
            if self.allow_multiple:
                ps = filedialog.askopenfilenames(
                    title="ファイルを選択",
                    filetypes=self.file_types or [("All", "*.*")]
                )
                if ps:
                    self._set_paths(list(ps))
            else:
                p = filedialog.askopenfilename(
                    title="ファイルを選択",
                    filetypes=self.file_types or [("All", "*.*")]
                )
                if p:
                    self._set_paths([p])

    def _set_paths(self, paths: list):
        if self.select_mode == "folder":
            self.selected_paths = [paths[0]]
        elif self.allow_multiple:
            ex = set(self.selected_paths)
            for p in paths:
                if p not in ex:
                    self.selected_paths.append(p)
                    ex.add(p)
        else:
            self.selected_paths = [paths[0]]
        self._update_display()
        if self.on_change:
            self.on_change()

    def _update_display(self):
        if not self.selected_paths:
            self.info_label.config(text="")
            self._draw_border(C["border"])
            if self.clr_btn:
                self.clr_btn.destroy()
                self.clr_btn = None
            return

        self._draw_border(C["ok"])

        if self.select_mode == "folder":
            folder = Path(self.selected_paths[0])
            files = [f for f in folder.rglob("*")
                     if f.is_file() and f.suffix.lower() in LINK_EXTENSIONS]
            text = f"✓  {folder.name}/  （{len(files)} ファイル）"
        else:
            names = [Path(p).name for p in self.selected_paths]
            if len(names) <= 4:
                text = "  ".join(f"✓ {n}" for n in names)
            else:
                text = "  ".join(f"✓ {n}" for n in names[:3])
                text += f"  … 他 {len(names)-3} 件"

        self.info_label.config(text=text, fg=C["ok"])

        if not self.clr_btn:
            self.clr_btn = _btn(
                self.sel_btn.master, "クリア", self._clear,
                "#3A1010", fg=C["err"]
            )
            self.clr_btn.pack(side=tk.LEFT)

    def _clear(self):
        self.selected_paths = []
        self.info_label.config(fg=C["info"])
        self._update_display()
        if self.on_change:
            self.on_change()

    # tkdnd ネイティブDnD登録
    def _setup_tkdnd_native(self, root):
        try:
            widget_path = str(self)
            root.tk.eval(f"tkdnd::drop_target register {widget_path} DND_Files")
            self.tk.createcommand(f"lf_drop_{id(self)}", self._on_tkdnd_drop)
            self.tk.createcommand(f"lf_enter_{id(self)}", self._on_dnd_enter)
            self.tk.createcommand(f"lf_leave_{id(self)}", self._on_dnd_leave)
            root.tk.eval(f"""
                bind {widget_path} <<Drop>> {{ lf_drop_{id(self)} %D; break }}
                bind {widget_path} <<DragEnter>> {{ lf_enter_{id(self)} }}
                bind {widget_path} <<DragLeave>> {{ lf_leave_{id(self)} }}
            """)
        except Exception:
            pass

    def _on_tkdnd_drop(self, *args):
        self._on_dnd_leave()
        raw = " ".join(str(a) for a in args)
        paths = self._parse_dnd_paths(raw)
        if paths:
            self._set_paths(paths)

    def _on_dnd_enter(self, *args):
        self._draw_border(C["primary"])

    def _on_dnd_leave(self, *args):
        if self.selected_paths:
            self._draw_border(C["ok"])
        else:
            self._draw_border(C["border"])

    @staticmethod
    def _parse_dnd_paths(raw: str) -> list:
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
                if end == -1:
                    end = len(raw)
                paths.append(raw[i:end])
                i = end + 1
        return [p for p in paths if p.strip()]


# ════════════════════════════════════════════════════════════════
#  メインアプリ
# ════════════════════════════════════════════════════════════════

class LinkForgeApp:

    def __init__(self):
        global _DND_BACKEND
        self.root = tk.Tk()
        self.root.title(f"LinkForge v{APP_VERSION}")
        self.root.geometry("620x860")
        self.root.minsize(520, 720)
        self.root.configure(bg=C["bg"])

        if platform.system() == "Darwin":
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(100, lambda: self.root.attributes("-topmost", False))
            try:
                self.root.createcommand("::tk::mac::OpenDocument", self._on_mac_open_doc)
            except Exception:
                pass

        self.folder_name_entries: list = []
        self._build_ui()
        self.root.after(300, self._init_tkdnd)

    # ── UI構築 ──────────────────────────────────────────────────
    def _build_ui(self):
        # ヘッダー（Canvas描画）
        hdr = tk.Canvas(
            self.root, height=56, bg=C["accent"],
            highlightthickness=0, bd=0
        )
        hdr.pack(fill=tk.X)
        hdr.bind("<Configure>", self._redraw_header)
        self._hdr_canvas = hdr

        hdr.create_text(
            20, 28, anchor="w",
            text="LinkForge", fill=C["primary"],
            font=("Helvetica", 20, "bold"), tags="hdr"
        )
        hdr.create_text(
            155, 30, anchor="w",
            text="Word ハイパーリンク自動挿入",
            fill=C["sub"], font=("Helvetica", 10), tags="hdr"
        )
        self._dnd_tag_id = hdr.create_text(
            580, 28, anchor="e",
            text="---", fill=C["sub"],
            font=("Helvetica", 9), tags="hdr"
        )
        hdr.create_text(
            610, 28, anchor="e",
            text=f"v{APP_VERSION}", fill=C["sub"],
            font=("Helvetica", 8), tags="hdr"
        )

        # アクセントライン
        line = tk.Canvas(
            self.root, height=3, bg=C["primary"],
            highlightthickness=0, bd=0
        )
        line.pack(fill=tk.X)

        # スクロール可能なメインエリア
        outer = tk.Frame(self.root, bg=C["bg"])
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, bg=C["bg"], highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._scroll_canvas = canvas
        main = tk.Frame(canvas, bg=C["bg"])
        self._main_frame_id = canvas.create_window((0, 0), window=main, anchor="nw")
        main.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(
            self._main_frame_id, width=e.width))

        pad = {"padx": 16, "pady": (0, 10)}

        # ── ① 計画書 ──
        self.word_zone = DropZone(
            main,
            label_text="計画書（Word ファイル）",
            hint_text="クリックして選択、またはパスをペーストしてください",
            select_mode="file",
            file_types=[("Word 文書", "*.docx")],
            allow_multiple=True
        )
        self.word_zone.pack(fill=tk.X, **pad)
        self.word_zone.on_change = self._on_word_changed

        # ── ② リンク資料フォルダ ──
        self.link_zone = DropZone(
            main,
            label_text="リンク資料フォルダ",
            hint_text="リンクしたい資料が入ったフォルダを選択してください",
            select_mode="folder"
        )
        self.link_zone.pack(fill=tk.X, **pad)
        self.link_zone.on_change = self._check_ready

        # ── ③ 出力先フォルダ ──
        self.output_zone = DropZone(
            main,
            label_text="出力先フォルダ",
            hint_text="完成ファイルの保存先を選択してください（Dropboxなど）",
            select_mode="folder"
        )
        self.output_zone.pack(fill=tk.X, **pad)
        self.output_zone.on_change = self._check_ready

        # ── ④ 出力フォルダ名 ──
        self.fname_outer = tk.Frame(main, bg=C["bg"])

        # セパレーター
        sep = tk.Canvas(main, height=1, bg=C["border"],
                        highlightthickness=0, bd=0)
        sep.pack(fill=tk.X, padx=16, pady=(0, 10))

        # ── 実行ボタン ──
        self.run_btn = tk.Button(
            main,
            text="▶  リンクを作成",
            command=self._run_process,
            font=("Helvetica", 14, "bold"),
            bg=C["accent"], fg=C["sub"],
            activebackground=C["primary"], activeforeground="white",
            relief=tk.FLAT, bd=0, padx=30, pady=12,
            cursor="arrow", state=tk.DISABLED
        )
        self.run_btn.pack(pady=(0, 10))

        # ── ログエリア ──
        log_outer = tk.Frame(main, bg=C["accent"], padx=1, pady=1)
        log_outer.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))

        log_hdr = tk.Frame(log_outer, bg=C["accent"])
        log_hdr.pack(fill=tk.X)
        tk.Label(
            log_hdr, text=" 処理ログ",
            font=("Helvetica", 9), bg=C["accent"], fg=C["sub"]
        ).pack(side=tk.LEFT)

        log_inner = tk.Frame(log_outer, bg=C["input_bg"])
        log_inner.pack(fill=tk.BOTH, expand=True)

        mono = "Menlo" if platform.system() == "Darwin" else "Consolas"
        self.log_text = tk.Text(
            log_inner, height=8,
            font=(mono, 9),
            bg=C["input_bg"], fg=C["sub"],
            insertbackground=C["primary"],
            relief=tk.FLAT, bd=0,
            padx=10, pady=8,
            state=tk.DISABLED, wrap=tk.WORD
        )
        sb = ttk.Scrollbar(log_inner, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_configure("success", foreground=C["ok"])
        self.log_text.tag_configure("error",   foreground=C["err"])
        self.log_text.tag_configure("info",    foreground=C["info"])
        self.log_text.tag_configure("warning", foreground=C["warn"])

    def _redraw_header(self, event):
        w = event.width
        self._hdr_canvas.coords(self._dnd_tag_id, w - 60, 28)

    # ── tkdnd 初期化 ──────────────────────────────────────────
    def _init_tkdnd(self):
        global _DND_BACKEND
        try:
            self.root.tk.eval("package require tkdnd")
            _DND_BACKEND = "tkdnd"
            for zone in [self.word_zone, self.link_zone, self.output_zone]:
                zone._setup_tkdnd_native(self.root)
            self._log("ドラッグ＆ドロップ: 有効", "success")
            for zone in [self.word_zone, self.link_zone, self.output_zone]:
                old = zone.hint_label.cget("text")
                zone.hint_label.configure(
                    text=old.replace("パスをペースト", "ドラッグ＆ドロップ / パスをペースト")
                )
            self._hdr_canvas.itemconfig(self._dnd_tag_id, text="D&D", fill=C["ok"])
        except Exception:
            self._log("ドラッグ＆ドロップ: ダイアログ / パスペーストを利用", "info")

    # ── ログ ──────────────────────────────────────────────────
    def _log(self, msg: str, tag: str = ""):
        self.log_text.configure(state=tk.NORMAL)
        if tag:
            self.log_text.insert(tk.END, msg + "\n", tag)
        else:
            self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # ── フォルダ名セクション ──────────────────────────────────
    def _on_word_changed(self):
        self._rebuild_folder_names()
        self._check_ready()

    def _rebuild_folder_names(self):
        self.fname_outer.pack_forget()
        for w in self.fname_outer.winfo_children():
            w.destroy()
        self.folder_name_entries = []

        word_paths = self.word_zone.selected_paths
        if not word_paths:
            return

        tk.Label(
            self.fname_outer,
            text="出力フォルダ名（変更可）",
            font=("Helvetica", 10, "bold"),
            bg=C["bg"], fg=C["sub"]
        ).pack(anchor="w", padx=16, pady=(0, 4))

        for wp in word_paths:
            p = Path(wp)
            if p.name.startswith("~$"):
                continue
            row = tk.Frame(self.fname_outer, bg=C["bg"])
            row.pack(fill=tk.X, padx=16, pady=2)

            tk.Label(
                row, text=f"{p.name}  →",
                font=("Helvetica", 9), bg=C["bg"], fg=C["sub"],
                anchor="e", width=30
            ).pack(side=tk.LEFT, padx=(0, 6))

            var = tk.StringVar(value=p.stem)
            tk.Entry(
                row, textvariable=var,
                font=("Helvetica", 10),
                bg=C["input_bg"], fg=C["text"],
                insertbackground=C["primary"],
                relief=tk.FLAT, bd=1
            ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

            def _reset(v=var, d=p.stem):
                v.set(d)
            _btn(row, "戻す", _reset, C["accent"],
                 font_size=8).pack(side=tk.LEFT)

            self.folder_name_entries.append((wp, var))

        self.fname_outer.pack(fill=tk.X, pady=(0, 8), before=self.run_btn)

    # ── 実行可否チェック ──────────────────────────────────────
    def _check_ready(self):
        ok = (self.word_zone.selected_paths and
              self.link_zone.selected_paths and
              self.output_zone.selected_paths)
        if ok:
            self.run_btn.configure(
                state=tk.NORMAL, bg=C["primary"], fg="white", cursor="hand2"
            )
        else:
            self.run_btn.configure(
                state=tk.DISABLED, bg=C["accent"], fg=C["sub"], cursor="arrow"
            )

    # ── 実行 ──────────────────────────────────────────────────
    def _run_process(self):
        self.run_btn.configure(
            state=tk.DISABLED, text="⏳  処理中...", bg=C["accent"], fg=C["warn"]
        )
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            link_dir   = Path(self.link_zone.selected_paths[0])
            output_dir = Path(self.output_zone.selected_paths[0])

            folder_names = {}
            for wp, var in self.folder_name_entries:
                name = var.get().strip() or Path(wp).stem
                folder_names[wp] = name

            file_map = get_file_map(link_dir)
            if not file_map:
                self.root.after(0, lambda: self._log(
                    "[警告] リンク資料フォルダに対象ファイルがありません。", "warning"))
                self.root.after(0, self._reset_btn)
                return

            self.root.after(0, lambda: self._log("リンク対象ファイル:", "info"))
            seen = set()
            for v in file_map.values():
                if v not in seen:
                    vv = v
                    self.root.after(0, lambda x=vv: self._log(f"  {x}"))
                    seen.add(v)
            self.root.after(0, lambda: self._log(""))

            entries = [(wp, fn) for wp, fn in folder_names.items()
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
                    copy_link_tree(link_dir, out)
                    self.root.after(0, lambda: self._log("    資料コピー完了"))

                    cnt = 0
                    for para in iter_all_paragraphs(doc):
                        cnt += process_paragraph(para, file_map, doc.part)
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

    def _handle_drop(self, paths):
        if not paths:
            return
        docx = [p for p in paths if p.lower().endswith(".docx")]
        dirs = [p for p in paths if os.path.isdir(p)]
        if docx:
            self.word_zone._set_paths(docx)
            self._log(f"Word ファイルを追加: {len(docx)} 件", "info")
        elif dirs:
            if not self.link_zone.selected_paths:
                self.link_zone._set_paths(dirs[:1])
                self._log(f"リンク資料フォルダ: {Path(dirs[0]).name}", "info")
            elif not self.output_zone.selected_paths:
                self.output_zone._set_paths(dirs[:1])
                self._log(f"出力先フォルダ: {Path(dirs[0]).name}", "info")
            else:
                ch = messagebox.askyesnocancel(
                    "振り分け",
                    f"「{Path(dirs[0]).name}」をどこに？\n\n"
                    "はい → リンク資料フォルダ\nいいえ → 出力先フォルダ"
                )
                if ch is True:
                    self.link_zone._set_paths(dirs[:1])
                elif ch is False:
                    self.output_zone._set_paths(dirs[:1])
        else:
            self.word_zone._set_paths(paths)

    def _on_mac_open_doc(self, *args):
        paths = [str(a) for a in args if os.path.exists(str(a))]
        if paths:
            self._handle_drop(paths)

    def run(self):
        self.root.mainloop()


# ════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    LinkForgeApp().run()