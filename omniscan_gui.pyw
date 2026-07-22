#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OmniScan GUI — Tkinter front-end for the OmniScan engine (omniscan.py).

Pure stdlib (tkinter/ttk/threading/queue). The engine file is imported
unmodified; keep omniscan.py in the same directory.

LICENSE: MIT, Copyright (c) 2026 Stephen Berson — same terms as omniscan.py
(see the license block in that file).

Layout (MaINbox/SmartScan review style):
  toolbar   : Add Files / Scan / Export JSON / Copy Items / Clear
  options   : catalog DB, teach-rules JSON, OCR toggle, Ollama host,
              min confidence, include-attachment toggle, Teach-on-correct
  left pane : hierarchical results tree (files -> attachments/members)
  right pane: notebook —
      Review: SIDE-BY-SIDE comparison. Left half = ORIGINAL DOCUMENT text,
              right half = EXTRACTED OUTPUT (one formatted line per item).
              A colored rectangle is drawn around the source line AND the
              output line of the item the user is currently on, kept in
              sync while navigating (click either side, or Prev/Next).
              Below: correction fields + Apply / Delete / Add / Prev / Next.
      Warnings & Meta
  statusbar : progress + counts + teach-rule count

Teach: with "Teach on correct" enabled, Apply writes a set-rule, Delete
writes a suppress-rule, and Add From Selected Line writes a set-rule to the
teach JSON. The ENGINE consults these rules on every future scan, so
corrections stick permanently across sessions — same philosophy as
SmartScan's learned corrections / dismissed_candidates.json.

All scanning and catalog loading runs on a worker thread; the UI thread
only drains a queue via after() polling — no blocking on the main loop.
"""

APP_TITLE = "OmniScan GUI v0.5.1"  # v0.5.1: manual Qty corrections use the engine thousands-aware parser (typing 1,000 means one thousand, not 1.0) | v0.5.0: teach persistence fix — default teach file now lives next to the script (writability-probed, home-dir fallback) instead of the process CWD, which on double-clicked Windows apps is system32 where writes silently failed; every teach write is verified and a failure raises a loud error dialog instead of losing the correction | v0.4.1: preview pages rotate by the engine's per-page deskew angle so on-image overlay boxes stay aligned on slanted scans | v0.4.0: SmartScan-style document Preview tab — renders the actual PDF pages (pypdfium2) or image and draws the colored rectangle on the document itself via engine OCR line boxes, synced with the extracted-output box; tabs restructured to Preview / Extracted-Merged Text / Warnings & Meta; Description column moved right of Part; smoke test scans a rendered OCR table image end-to-end and asserts correct qtys + on-image overlay | v0.3.0: side-by-side review w/ synced rectangles + teach-on-correct persistence | v0.2.0: review loop (highlight/correct/delete/add) | v0.1.0: initial release

import json
import os
import queue
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional, Tuple

try:
    import omniscan
except ImportError as exc:
    sys.stderr.write("omniscan_gui: omniscan.py must be in the same directory "
                     f"or on PYTHONPATH ({exc})\n")
    raise

try:
    from PIL import Image as PILImage, ImageTk
    _PIL_OK = True
except ImportError:
    _PIL_OK = False

_MONO = ("Consolas", 9)
_COL_BOX = "#ffb84d"        # colored rectangle for the current pair
_COL_BOX_BG = "#fff1d6"
_COL_TAUGHT = "#e4f4e0"     # taught / corrected line tint
_COL_PICK = "#cfe6ff"       # user-picked source line (add candidate)
_PREVIEW_W = 640            # preview render width in px
_OUT_HEADER = ("  QTY      U/M  PART                "
               "DESCRIPTION                                 GAUGE     CONF")


class OmniScanGUI(tk.Tk):
    POLL_MS = 100

    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE.split(' # ')[0]}  |  engine: "
                   f"{omniscan.APP_TITLE.split(' # ')[0]}")
        self.geometry("1240x760")
        self.minsize(940, 580)

        self.paths: List[str] = []
        self.results: List[omniscan.ScanResult] = []
        self.node_map: Dict[str, omniscan.ScanResult] = {}
        self.res_to_iid: Dict[int, str] = {}
        self.msg_q: "queue.Queue" = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.scanning = False
        self.catalog: Optional["omniscan.Catalog"] = None
        self.teach: Optional["omniscan.TeachStore"] = None

        # review state
        self.pairs: List[Tuple[omniscan.ScanResult, omniscan.LineItem]] = []
        self.cur: Optional[int] = None           # index into pairs
        self.src_res: Optional[omniscan.ScanResult] = None  # doc shown on the left
        self.picked_line: Optional[int] = None   # 0-based picked source line

        self._build_ui()
        self._load_teach_store()
        self.after(self.POLL_MS, self._poll)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        bar = ttk.Frame(self, padding=(6, 6, 6, 2))
        bar.pack(fill="x")
        self.btn_add = ttk.Button(bar, text="Add Files…", command=self._on_add)
        self.btn_scan = ttk.Button(bar, text="Scan", command=self._start_scan)
        self.btn_export = ttk.Button(bar, text="Export JSON…", command=self._on_export)
        self.btn_copy = ttk.Button(bar, text="Copy Items (TSV)", command=self._on_copy_items)
        self.btn_clear = ttk.Button(bar, text="Clear", command=self._on_clear)
        for b in (self.btn_add, self.btn_scan, self.btn_export, self.btn_copy, self.btn_clear):
            b.pack(side="left", padx=(0, 6))
        ttk.Label(bar, text=self._deps_summary(), foreground="#555").pack(side="right")

        opts = ttk.Frame(self, padding=(6, 2, 6, 4))
        opts.pack(fill="x")
        ttk.Label(opts, text="Catalog DB:").pack(side="left")
        self.var_catalog = tk.StringVar()
        ttk.Entry(opts, textvariable=self.var_catalog, width=28).pack(side="left", padx=(4, 2))
        ttk.Button(opts, text="…", width=3, command=self._on_pick_catalog).pack(side="left", padx=(0, 10))
        ttk.Label(opts, text="Teach file:").pack(side="left")
        self.var_teach = tk.StringVar(value=self._default_teach_path())
        teach_entry = ttk.Entry(opts, textvariable=self.var_teach, width=28)
        teach_entry.pack(side="left", padx=(4, 2))
        teach_entry.bind("<FocusOut>", lambda _e: self._load_teach_store())
        ttk.Button(opts, text="…", width=3, command=self._on_pick_teach).pack(side="left", padx=(0, 10))
        self.var_ocr = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="OCR", variable=self.var_ocr).pack(side="left", padx=(0, 10))
        ttk.Label(opts, text="Ollama:").pack(side="left")
        self.var_ollama = tk.StringVar()
        ttk.Entry(opts, textvariable=self.var_ollama, width=18).pack(side="left", padx=(4, 10))
        ttk.Label(opts, text="Min conf:").pack(side="left")
        self.var_conf = tk.DoubleVar(value=0.45)
        ttk.Spinbox(opts, from_=0.0, to=1.0, increment=0.05, width=5,
                    textvariable=self.var_conf).pack(side="left", padx=(4, 10))
        self.var_aggregate = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Incl. attachments", variable=self.var_aggregate,
                        command=self._rebuild_review).pack(side="left", padx=(0, 10))
        self.var_teach_on = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Teach on correct",
                        variable=self.var_teach_on).pack(side="left")

        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=6, pady=(0, 2))

        # Results tree
        left = ttk.Frame(paned)
        self.tree = ttk.Treeview(left, columns=("fmt", "ok", "items", "warn"),
                                 selectmode="browse")
        self.tree.heading("#0", text="File / Attachment")
        self.tree.column("#0", width=230, stretch=True)
        for col, label, width in (("fmt", "Fmt", 55), ("ok", "OK", 36),
                                  ("items", "Items", 46), ("warn", "Warn", 46)):
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor="center", stretch=False)
        ysb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._rebuild_review())
        paned.add(left, weight=1)

        right = ttk.Frame(paned)
        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)

        # ---- Preview tab: document render | extracted output ----
        preview_tab = ttk.Frame(self.notebook)
        split = ttk.Panedwindow(preview_tab, orient="horizontal")
        split.pack(fill="both", expand=True)

        pv_holder = ttk.Frame(split)
        ttk.Label(pv_holder, text="ORIGINAL DOCUMENT", anchor="center",
                  font=("TkDefaultFont", 9, "bold")).pack(fill="x")
        # canvas (visual preview) and fallback text share this holder
        self.pv_frame = ttk.Frame(pv_holder)
        self.pv_canvas = tk.Canvas(self.pv_frame, background="#9a9a9a",
                                   highlightthickness=0)
        pvy = ttk.Scrollbar(self.pv_frame, orient="vertical",
                            command=self.pv_canvas.yview)
        pvx = ttk.Scrollbar(self.pv_frame, orient="horizontal",
                            command=self.pv_canvas.xview)
        self.pv_canvas.configure(yscrollcommand=pvy.set, xscrollcommand=pvx.set)
        pvy.pack(side="right", fill="y")
        pvx.pack(side="bottom", fill="x")
        self.pv_canvas.pack(fill="both", expand=True)
        self.pv_fallback_frame = ttk.Frame(pv_holder)
        self.pv_fallback = tk.Text(self.pv_fallback_frame, wrap="none", font=_MONO,
                                   cursor="arrow")
        pfy = ttk.Scrollbar(self.pv_fallback_frame, orient="vertical",
                            command=self.pv_fallback.yview)
        self.pv_fallback.configure(yscrollcommand=pfy.set, state="disabled")
        pfy.pack(side="right", fill="y")
        self.pv_fallback.pack(fill="both", expand=True)
        self.pv_frame.pack(fill="both", expand=True)      # canvas shown by default
        self.pv_images: List = []                          # PhotoImage refs
        self.pv_pages: List[Dict] = []                     # {y,w,h} per page
        self.pv_res: Optional[omniscan.ScanResult] = None
        split.add(pv_holder, weight=1)

        out_holder = ttk.Frame(split)
        ttk.Label(out_holder, text="EXTRACTED OUTPUT", anchor="center",
                  font=("TkDefaultFont", 9, "bold")).pack(fill="x")
        self.out = tk.Text(out_holder, wrap="none", font=_MONO, cursor="arrow")
        self._config_box_tags(self.out)
        self.out.tag_configure("taught", background=_COL_TAUGHT)
        self.out.tag_configure("hdr", foreground="#666")
        self.out.bind("<Button-1>", self._on_out_click)
        oysb = ttk.Scrollbar(out_holder, orient="vertical", command=self.out.yview)
        oxsb = ttk.Scrollbar(out_holder, orient="horizontal", command=self.out.xview)
        self.out.configure(yscrollcommand=oysb.set, xscrollcommand=oxsb.set,
                           state="disabled")
        oysb.pack(side="right", fill="y")
        oxsb.pack(side="bottom", fill="x")
        self.out.pack(fill="both", expand=True)
        split.add(out_holder, weight=1)

        # Correction / teach controls
        panel = ttk.Labelframe(preview_tab, text="Correct & Teach", padding=6)
        panel.pack(fill="x", padx=2, pady=(4, 2))
        labels = (("Qty", 7), ("U/M", 6), ("Part", 20), ("Gauge", 10))
        self.var_e_qty = tk.StringVar()
        self.var_e_uom = tk.StringVar()
        self.var_e_part = tk.StringVar()
        self.var_e_gauge = tk.StringVar()
        edit_vars = (self.var_e_qty, self.var_e_uom, self.var_e_part, self.var_e_gauge)
        for col, ((label, width), var) in enumerate(zip(labels, edit_vars)):
            ttk.Label(panel, text=label + ":").grid(row=0, column=col * 2, sticky="e")
            ttk.Entry(panel, textvariable=var, width=width).grid(
                row=0, column=col * 2 + 1, sticky="w", padx=(4, 10))
        ttk.Label(panel, text="Desc:").grid(row=1, column=0, sticky="e", pady=(4, 0))
        self.var_e_desc = tk.StringVar()
        ttk.Entry(panel, textvariable=self.var_e_desc).grid(
            row=1, column=1, columnspan=7, sticky="ew", padx=(4, 0), pady=(4, 0))
        btns = ttk.Frame(panel)
        btns.grid(row=2, column=0, columnspan=8, sticky="w", pady=(8, 0))
        ttk.Button(btns, text="◀ Prev", width=7,
                   command=lambda: self._step(-1)).pack(side="left", padx=(0, 4))
        ttk.Button(btns, text="Next ▶", width=7,
                   command=lambda: self._step(1)).pack(side="left", padx=(0, 12))
        ttk.Button(btns, text="Apply + Teach", command=self._on_apply).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Delete Item", command=self._on_delete_item).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Add From Selected Line",
                   command=self._on_add_from_line).pack(side="left", padx=(0, 8))
        panel.columnconfigure(7, weight=1)
        self.notebook.add(preview_tab, text="Preview")

        # ---- Extracted / Merged Text tab ----
        text_tab = ttk.Frame(self.notebook)
        self.src = tk.Text(text_tab, wrap="none", font=_MONO, cursor="arrow")
        self._config_box_tags(self.src)
        self.src.tag_configure("pick", background=_COL_PICK)
        self.src.bind("<Button-1>", self._on_src_click)
        sysb = ttk.Scrollbar(text_tab, orient="vertical", command=self.src.yview)
        sxsb = ttk.Scrollbar(text_tab, orient="horizontal", command=self.src.xview)
        self.src.configure(yscrollcommand=sysb.set, xscrollcommand=sxsb.set,
                           state="disabled")
        sysb.pack(side="right", fill="y")
        sxsb.pack(side="bottom", fill="x")
        self.src.pack(fill="both", expand=True)
        self.notebook.add(text_tab, text="Extracted / Merged Text")

        # ---- Warnings & Meta tab ----
        warn_frame = ttk.Frame(self.notebook)
        self.warn = tk.Text(warn_frame, wrap="word", font=_MONO)
        wysb = ttk.Scrollbar(warn_frame, orient="vertical", command=self.warn.yview)
        self.warn.configure(yscrollcommand=wysb.set, state="disabled")
        self.warn.pack(side="left", fill="both", expand=True)
        wysb.pack(side="right", fill="y")
        self.notebook.add(warn_frame, text="Warnings & Meta")

        paned.add(right, weight=3)

        status = ttk.Frame(self, padding=(6, 2, 6, 4))
        status.pack(fill="x")
        self.status_var = tk.StringVar(value="Ready. Add files, then Scan.")
        ttk.Label(status, textvariable=self.status_var).pack(side="left")
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=150)
        self.progress.pack(side="right")

    @staticmethod
    def _config_box_tags(widget: tk.Text):
        """The colored rectangle drawn around the current line in each pane."""
        try:
            widget.tag_configure("box", background=_COL_BOX_BG,
                                 borderwidth=2, relief="solid")
        except tk.TclError:
            widget.tag_configure("box", background=_COL_BOX_BG,
                                 borderwidth=2, relief="ridge")
        # colored edge accents so the box reads as a colored rectangle
        widget.tag_configure("boxedge", background=_COL_BOX)

    @staticmethod
    def _deps_summary() -> str:
        d = omniscan.DEPS
        return ("PDF:" + ("yes" if (d.pypdf or d.pdfplumber) else "NO")
                + "  OCR:" + ("yes" if (d.pytesseract and d.PIL) else "no")
                + "  PDF-OCR:" + ("yes" if (d.pypdfium2 and d.pytesseract and d.PIL) else "no")
                + "  MSG:native")

    # ----------------------------------------------------------- commands
    def _on_add(self):
        picked = filedialog.askopenfilenames(
            title="Add files to scan",
            filetypes=[("All supported", "*.pdf *.docx *.xlsx *.csv *.tsv *.txt "
                        "*.htm *.html *.eml *.msg *.zip *.png *.jpg *.jpeg *.gif "
                        "*.bmp *.tif *.tiff"),
                       ("All files", "*.*")])
        if picked:
            self._add_paths(list(picked))

    def _add_paths(self, paths: List[str]):
        added = sum(1 for p in paths if p not in self.paths and not self.paths.append(p))
        self.status_var.set(f"{len(self.paths)} file(s) queued ({added} added). Press Scan.")

    def _on_pick_catalog(self):
        picked = filedialog.askopenfilename(
            title="Select catalog SQLite DB",
            filetypes=[("SQLite DB", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")])
        if picked:
            self.var_catalog.set(picked)

    def _on_pick_teach(self):
        picked = filedialog.asksaveasfilename(
            title="Teach-rules JSON (created if missing)", defaultextension=".json",
            filetypes=[("JSON", "*.json")], confirmoverwrite=False)
        if picked:
            self.var_teach.set(picked)
            self._load_teach_store()

    @staticmethod
    def _default_teach_path() -> str:
        """Teach file lives next to the script when writable, else in the
        user's home dir. NEVER the process CWD — a double-clicked app on
        Windows starts in system32, where writes fail."""
        candidates = [os.path.dirname(os.path.abspath(__file__)),
                      os.path.expanduser("~")]
        for d in candidates:
            p = os.path.join(d, "omniscan_teach.json")
            try:
                with open(p, "a", encoding="utf-8"):
                    pass
                return p
            except OSError:
                continue
        return os.path.join(os.path.expanduser("~"), "omniscan_teach.json")

    def _load_teach_store(self):
        path = self.var_teach.get().strip()
        self.teach = omniscan.TeachStore(path) if path else None
        if self.teach is None:
            self.status_var.set("No teach file set — corrections will not persist.")
            return
        self.teach.save()                      # writability probe
        n = len(self.teach.rules)
        if self.teach.last_error:
            self.status_var.set(f"TEACH FILE NOT WRITABLE: {path}")
            messagebox.showerror(
                APP_TITLE.split(" # ")[0],
                f"The teach file cannot be written:\n{path}\n\n"
                f"{self.teach.last_error}\n\n"
                "Corrections will NOT persist until you pick a writable "
                "location with the … button.")
        else:
            self.status_var.set(f"Teach file: {path} — {n} learned rule(s).")

    def _teach_write_ok(self) -> bool:
        """Called after every teach write; a failed save is data loss and
        must be loud, not silent."""
        if self.teach and self.teach.last_error:
            messagebox.showerror(
                APP_TITLE.split(" # ")[0],
                f"TEACH SAVE FAILED — this correction was NOT persisted:\n"
                f"{self.teach.path}\n\n{self.teach.last_error}\n\n"
                "Pick a writable teach file location with the … button, "
                "then Apply again.")
            return False
        return True

    def _on_clear(self):
        if self.scanning:
            messagebox.showinfo(APP_TITLE, "Wait for the current scan to finish.")
            return
        self.paths.clear()
        self.results.clear()
        self.node_map.clear()
        self.res_to_iid.clear()
        self.tree.delete(*self.tree.get_children())
        self._clear_review()
        self.status_var.set("Cleared. Add files, then Scan.")

    def _on_export(self):
        if not self.results:
            messagebox.showinfo(APP_TITLE, "Nothing to export yet — run a scan first.")
            return
        path = filedialog.asksaveasfilename(
            title="Export results JSON", defaultextension=".json",
            filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump([r.to_dict() for r in self.results], fh, indent=2)
            self.status_var.set(f"Exported: {path}")
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Export failed: {exc}")

    def _on_copy_items(self):
        if not self.pairs:
            messagebox.showinfo(APP_TITLE, "No line items in the current view.")
            return
        header = "Qty\tU/M\tPart\tGauge\tConf\tCatalog\tDescription\tSourceFile"
        rows = []
        for owner, li in self.pairs:
            cat = (f"{li.catalog_match.get('part', '')} ({li.catalog_match.get('match', '')})"
                   if li.catalog_match else "")
            rows.append("\t".join([self._fmt_qty(li.qty), li.uom or "", li.part or "",
                                   li.gauge or "", f"{li.confidence:.2f}", cat,
                                   li.description or li.raw, owner.name]))
        self.clipboard_clear()
        self.clipboard_append(header + "\n" + "\n".join(rows))
        self.status_var.set(f"Copied {len(rows)} line item(s) to clipboard as TSV.")

    # ---------------------------------------------------------- scanning
    def _current_cfg(self) -> "omniscan.Config":
        host = self.var_ollama.get().strip() or None
        try:
            conf = max(0.0, min(1.0, float(self.var_conf.get())))
        except (tk.TclError, ValueError):
            conf = 0.45
        return omniscan.Config(ocr=bool(self.var_ocr.get()),
                               ollama_host=host,
                               catalog_path=self.var_catalog.get().strip() or None,
                               min_item_confidence=conf,
                               teach_path=self.var_teach.get().strip() or None)

    def _start_scan(self):
        if self.scanning:
            return
        if not self.paths:
            messagebox.showinfo(APP_TITLE, "Add files first.")
            return
        self.scanning = True
        for b in (self.btn_add, self.btn_scan, self.btn_clear):
            b.state(["disabled"])
        self.tree.delete(*self.tree.get_children())
        self.node_map.clear()
        self.res_to_iid.clear()
        self.results.clear()
        self._clear_review()
        self.progress.start(12)
        self.status_var.set("Scanning…")
        cfg = self._current_cfg()
        self.worker = threading.Thread(target=self._worker,
                                       args=(list(self.paths), cfg), daemon=True)
        self.worker.start()

    def _worker(self, paths: List[str], cfg: "omniscan.Config"):
        catalog = None
        try:
            if cfg.catalog_path:
                self.msg_q.put(("status", "Loading catalog…"))
                catalog = omniscan.Catalog(cfg.catalog_path)
                for warning in catalog.warnings:
                    self.msg_q.put(("catalog_warn", warning))
                self.msg_q.put(("catalog", catalog))
                self.msg_q.put(("status", f"Catalog: {len(catalog.index)} parts. Scanning…"))
            for i, path in enumerate(paths, 1):
                self.msg_q.put(("status", f"Scanning {i}/{len(paths)}: {os.path.basename(path)}"))
                res = omniscan.scan_path(path, cfg)
                if catalog:
                    self._ground_tree(catalog, res)
                self.msg_q.put(("result", res))
            self.msg_q.put(("done", len(paths)))
        except Exception as exc:
            self.msg_q.put(("error", f"{type(exc).__name__}: {exc}"))

    @staticmethod
    def _ground_tree(catalog: "omniscan.Catalog", res: "omniscan.ScanResult"):
        catalog.ground(res.line_items)
        for child in res.children:
            OmniScanGUI._ground_tree(catalog, child)

    def _poll(self):
        try:
            while True:
                kind, payload = self.msg_q.get_nowait()
                if kind == "result":
                    self.results.append(payload)
                    self._insert_node("", payload)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "catalog":
                    self.catalog = payload
                elif kind == "catalog_warn":
                    messagebox.showwarning(APP_TITLE, payload)
                elif kind == "done":
                    self._finish_scan(f"Done — {payload} file(s), "
                                      f"{sum(len(r.all_line_items()) for r in self.results)} "
                                      f"item(s). Review left-vs-right; corrections teach the parser.")
                elif kind == "error":
                    self._finish_scan(f"Scan error: {payload}")
                    messagebox.showerror(APP_TITLE, payload)
        except queue.Empty:
            pass
        self._poll_id = self.after(self.POLL_MS, self._poll)

    def destroy(self):
        poll_id = getattr(self, "_poll_id", None)
        if poll_id is not None:
            try:
                self.after_cancel(poll_id)
            except tk.TclError:
                pass
        super().destroy()

    def _finish_scan(self, status: str):
        self.scanning = False
        self.progress.stop()
        for b in (self.btn_add, self.btn_scan, self.btn_clear):
            b.state(["!disabled"])
        self.status_var.set(status)
        first = self.tree.get_children()
        if first:
            self.tree.selection_set(first[0])
            self.tree.focus(first[0])

    # ------------------------------------------------------- review pane
    def _insert_node(self, parent: str, res: "omniscan.ScanResult"):
        iid = self.tree.insert(parent, "end", text=res.name, open=True,
                               values=(res.fmt, "OK" if res.ok else "ERR",
                                       len(res.line_items), len(res.warnings)))
        self.node_map[iid] = res
        self.res_to_iid[id(res)] = iid
        for child in res.children:
            self._insert_node(iid, child)

    def _selected_result(self) -> Optional["omniscan.ScanResult"]:
        sel = self.tree.selection()
        return self.node_map.get(sel[0]) if sel else None

    @staticmethod
    def _walk(res):
        yield res
        for child in res.children:
            yield from OmniScanGUI._walk(child)

    @staticmethod
    def _fmt_qty(qty) -> str:
        if qty is None:
            return ""
        return str(int(qty)) if float(qty).is_integer() else str(qty)

    def _fmt_out_line(self, li: "omniscan.LineItem") -> str:
        mark = "T" if li.taught else ("*" if li.corrected else " ")
        desc = li.description or li.raw
        if len(desc) > 43:
            desc = desc[:42] + "…"
        return (f"{mark} {self._fmt_qty(li.qty) or '-':>7}  {li.uom or '-':<4} "
                f"{(li.part or '-'):<18} {desc:<43} "
                f"{(li.gauge or '-'):<9} {li.confidence:4.2f}")

    def _clear_review(self):
        self.pairs = []
        self.cur = None
        self.src_res = None
        self.pv_res = None
        self.picked_line = None
        self._set_editor(None)
        self.pv_canvas.delete("all")
        self.pv_images = []
        self.pv_pages = []
        for widget in (self.src, self.out, self.warn, self.pv_fallback):
            widget.configure(state="normal")
            widget.delete("1.0", "end")
            widget.configure(state="disabled")

    # -------------------------------------------------- document preview
    def _load_preview(self, owner: "omniscan.ScanResult"):
        if self.pv_res is owner:
            return
        self.pv_res = owner
        self.pv_canvas.delete("all")
        self.pv_images = []
        self.pv_pages = []
        path = owner.meta.get("path")
        rendered = False
        if _PIL_OK and path and os.path.exists(path):
            try:
                degs = owner.meta.get("deskew_deg") or []

                def straighten(pil_page, page_idx):
                    # mirror the engine's pre-OCR deskew so overlay boxes align
                    deg = degs[page_idx] if page_idx < len(degs) else 0.0
                    if deg:
                        pil_page = pil_page.rotate(deg, expand=True,
                                                   fillcolor=(255, 255, 255)
                                                   if pil_page.mode != "L" else 255)
                    return pil_page

                if owner.fmt == "pdf" and omniscan.DEPS.pypdfium2 is not None:
                    pdf = omniscan.DEPS.pypdfium2.PdfDocument(path)
                    y = 0
                    for i in range(len(pdf)):
                        w_pt, _h = pdf[i].get_size()
                        scale = _PREVIEW_W / max(w_pt, 1)
                        pil = straighten(pdf[i].render(scale=scale).to_pil(), i)
                        self._pv_add_page(pil, y)
                        y += pil.height + 8
                    pdf.close()
                    rendered = bool(self.pv_pages)
                elif owner.fmt == "image":
                    pil = straighten(PILImage.open(path).convert("RGB"), 0)
                    ratio = _PREVIEW_W / max(pil.width, 1)
                    pil = pil.resize((_PREVIEW_W, max(1, int(pil.height * ratio))),
                                     PILImage.LANCZOS)
                    self._pv_add_page(pil, 0)
                    rendered = True
            except Exception:
                rendered = False
        if rendered:
            total_h = self.pv_pages[-1]["y"] + self.pv_pages[-1]["h"]
            self.pv_canvas.configure(scrollregion=(0, 0, _PREVIEW_W, total_h))
            self.pv_fallback_frame.pack_forget()
            self.pv_frame.pack(fill="both", expand=True)
        else:
            # attachments / non-visual formats: fall back to extracted text
            note = ("(no visual preview for this item — attachment or "
                    "non-renderable format; showing extracted text)\n\n")
            self.pv_fallback.configure(state="normal")
            self.pv_fallback.delete("1.0", "end")
            self.pv_fallback.insert("1.0", note + (owner.text or ""))
            self.pv_fallback.configure(state="disabled")
            self.pv_frame.pack_forget()
            self.pv_fallback_frame.pack(fill="both", expand=True)

    def _pv_add_page(self, pil, y: int):
        photo = ImageTk.PhotoImage(pil)
        self.pv_images.append(photo)
        self.pv_canvas.create_image(0, y, anchor="nw", image=photo)
        self.pv_pages.append({"y": y, "w": pil.width, "h": pil.height})

    def _preview_box(self, owner: "omniscan.ScanResult",
                     li: Optional["omniscan.LineItem"]):
        """Draw the colored rectangle around the item's row ON the document
        image, using the engine's OCR line boxes."""
        self.pv_canvas.delete("ibox")
        if (li is None or li.line_no is None or owner is not self.pv_res
                or not self.pv_pages):
            return
        info = (owner.meta.get("ocr_lines") or {}).get(str(li.line_no))
        pages_meta = owner.meta.get("ocr_pages") or []
        if not info or info.get("page", 0) >= min(len(self.pv_pages), len(pages_meta)):
            return
        p = info["page"]
        page = self.pv_pages[p]
        src_w, src_h = pages_meta[p]
        rx = page["w"] / max(src_w, 1)
        ry = page["h"] / max(src_h, 1)
        x0, y0, x1, y1 = info["box"]
        cx0 = x0 * rx - 4
        cy0 = y0 * ry - 4 + page["y"]
        cx1 = x1 * rx + 4
        cy1 = y1 * ry + 4 + page["y"]
        self.pv_canvas.create_rectangle(cx0, cy0, cx1, cy1,
                                        outline=_COL_BOX, width=3, tags="ibox")
        total_h = self.pv_pages[-1]["y"] + self.pv_pages[-1]["h"]
        self.pv_canvas.yview_moveto(max(0.0, (cy0 - 60) / max(total_h, 1)))

    def _rebuild_review(self, keep_index: bool = False):
        want = self.cur if keep_index else 0
        res = self._selected_result()
        self.pairs = []
        self.cur = None
        self.src_res = None
        self.picked_line = None
        if res is None:
            self._clear_review()
            return
        if self.var_aggregate.get():
            self.pairs = [(o, li) for o in self._walk(res) for li in o.line_items]
        else:
            self.pairs = [(res, li) for li in res.line_items]
        # output pane
        self.out.configure(state="normal")
        self.out.delete("1.0", "end")
        self.out.insert("end", _OUT_HEADER + "\n", "hdr")
        for owner, li in self.pairs:
            tags = ("taught",) if (li.taught or li.corrected) else ()
            self.out.insert("end", self._fmt_out_line(li) + "\n", tags)
        self.out.configure(state="disabled")
        # warnings tab
        self.warn.configure(state="normal")
        self.warn.delete("1.0", "end")
        blocks = []
        if res.warnings:
            blocks.append("WARNINGS:\n" + "\n".join(f"  ! {w}" for w in res.warnings))
        blocks.append("META:\n" + json.dumps(res.meta, indent=2))
        self.warn.insert("1.0", "\n\n".join(blocks))
        self.warn.configure(state="disabled")
        # source pane + selection
        if self.pairs:
            idx = want if (keep_index and want is not None
                           and want < len(self.pairs)) else 0
            self._select_pair(idx)
        else:
            self._load_src(res)

    def _load_src(self, res: "omniscan.ScanResult"):
        self._load_preview(res)
        if self.src_res is res:
            return
        self.src_res = res
        self.picked_line = None
        self.src.configure(state="normal")
        self.src.delete("1.0", "end")
        self.src.insert("1.0", res.text or "(no text extracted)")
        self.src.configure(state="disabled")

    def _select_pair(self, idx: int):
        """Draw the synced colored rectangles: source line (left) and
        extracted output line (right) for pair idx."""
        if not self.pairs:
            return
        idx = max(0, min(idx, len(self.pairs) - 1))
        self.cur = idx
        owner, li = self.pairs[idx]
        self._load_src(owner)
        self._set_editor(li, owner)

        self.src.configure(state="normal")
        self.src.tag_remove("box", "1.0", "end")
        line_index = self._find_line(owner, li)
        if line_index is not None:
            tk_line = line_index + 1
            self.src.tag_add("box", f"{tk_line}.0", f"{tk_line}.end")
            self.src.see(f"{tk_line}.0")
        self.src.configure(state="disabled")

        self.out.configure(state="normal")
        self.out.tag_remove("box", "1.0", "end")
        out_line = idx + 2                      # +1 header, +1 1-based
        self.out.tag_add("box", f"{out_line}.0", f"{out_line}.end")
        self.out.see(f"{out_line}.0")
        self.out.configure(state="disabled")

        self._preview_box(owner, li)

        where = "" if owner is self._selected_result() else f"  [in: {owner.name}]"
        self.status_var.set(f"Item {idx + 1}/{len(self.pairs)}"
                            f"{where} — source line "
                            f"{'?' if line_index is None else line_index + 1}.")

    @staticmethod
    def _find_line(owner, li) -> Optional[int]:
        lines = (owner.text or "").splitlines()
        if li.line_no is not None and li.line_no < len(lines) \
                and lines[li.line_no].strip() == li.raw:
            return li.line_no
        for idx, line in enumerate(lines):
            if line.strip() == li.raw:
                return idx
        return None

    def _step(self, delta: int):
        if self.pairs:
            self._select_pair((self.cur or 0) + delta)

    def _on_out_click(self, event):
        index = self.out.index(f"@{event.x},{event.y}")
        line = int(index.split(".")[0])
        idx = line - 2                          # header offset
        if 0 <= idx < len(self.pairs):
            self._select_pair(idx)
        return "break"

    def _on_src_click(self, event):
        if self.src_res is None:
            return "break"
        index = self.src.index(f"@{event.x},{event.y}")
        line0 = int(index.split(".")[0]) - 1
        # If this source line belongs to an item of the displayed doc, jump to it
        for idx, (owner, li) in enumerate(self.pairs):
            if owner is self.src_res and self._find_line(owner, li) == line0:
                self._select_pair(idx)
                return "break"
        # Otherwise mark it as an add candidate
        self.picked_line = line0
        self.src.configure(state="normal")
        self.src.tag_remove("pick", "1.0", "end")
        self.src.tag_add("pick", f"{line0 + 1}.0", f"{line0 + 1}.end")
        self.src.configure(state="disabled")
        self.status_var.set(f"Source line {line0 + 1} selected — "
                            f"'Add From Selected Line' to add it as an item.")
        return "break"

    def _set_editor(self, li, owner=None):
        self._editing = (owner, li)
        self.var_e_qty.set("" if li is None else self._fmt_qty(li.qty))
        self.var_e_uom.set("" if li is None or not li.uom else li.uom)
        self.var_e_part.set("" if li is None or not li.part else li.part)
        self.var_e_gauge.set("" if li is None or not li.gauge else li.gauge)
        self.var_e_desc.set("" if li is None else (li.description or li.raw))

    # ------------------------------------------------------ correct/teach
    def _teach_fields(self, li) -> dict:
        return {"qty": li.qty, "uom": li.uom, "part": li.part,
                "gauge": li.gauge, "description": li.description}

    def _on_apply(self):
        owner, li = getattr(self, "_editing", (None, None))
        if li is None:
            messagebox.showinfo(APP_TITLE, "Select a line item first.")
            return
        qty_raw = self.var_e_qty.get().strip()
        if qty_raw:
            qty = omniscan.LineItemParser._to_qty(qty_raw)
            if qty is None:
                messagebox.showerror(APP_TITLE, f"Qty is not a number: {qty_raw!r}")
                return
        else:
            qty = None
        li.qty = qty
        li.uom = self.var_e_uom.get().strip().upper() or None
        new_part = self.var_e_part.get().strip().upper() or None
        part_changed = new_part != li.part
        li.part = new_part
        li.gauge = self.var_e_gauge.get().strip().upper() or None
        li.description = self.var_e_desc.get().strip() or li.raw
        li.confidence = 1.0
        li.corrected = True
        if part_changed:
            li.catalog_match = (self.catalog.lookup(li.part)
                                if (self.catalog and li.part) else None)
        taught_msg = ""
        if self.var_teach_on.get() and self.teach:
            self.teach.set_rule(li.raw, self._teach_fields(li))
            if self._teach_write_ok():
                taught_msg = f" Taught ({len(self.teach.rules)} rules)."
        self._rebuild_review(keep_index=True)
        self._update_tree_counts(owner)
        self.status_var.set(f"Correction applied: {li.part or li.description}.{taught_msg}")

    def _on_delete_item(self):
        owner, li = getattr(self, "_editing", (None, None))
        if li is None:
            messagebox.showinfo(APP_TITLE, "Select a line item first.")
            return
        try:
            owner.line_items.remove(li)
        except ValueError:
            pass
        taught_msg = ""
        if self.var_teach_on.get() and self.teach:
            self.teach.suppress(li.raw)
            if self._teach_write_ok():
                taught_msg = f" Taught suppression ({len(self.teach.rules)} rules)."
        cur = self.cur
        self._rebuild_review()
        if self.pairs and cur:
            self._select_pair(min(cur, len(self.pairs) - 1))
        self._update_tree_counts(owner)
        self.status_var.set(f"Line item deleted.{taught_msg}")

    def _on_add_from_line(self):
        if self.src_res is None or self.picked_line is None:
            messagebox.showinfo(APP_TITLE, "Click an unmatched line in the "
                                           "Extracted / Merged Text tab first.")
            return
        owner = self.src_res
        lines = (owner.text or "").splitlines()
        if self.picked_line >= len(lines) or not lines[self.picked_line].strip():
            messagebox.showinfo(APP_TITLE, "That line is empty.")
            return
        raw = lines[self.picked_line].strip()
        parser = omniscan.LineItemParser(self._current_cfg(), teach=self.teach)
        li = parser.parse_one(raw) or omniscan.LineItem(raw=raw, description=raw)
        li.line_no = self.picked_line
        li.confidence = 1.0
        li.corrected = True
        if self.catalog and li.part:
            li.catalog_match = self.catalog.lookup(li.part)
        owner.line_items.append(li)
        taught_msg = ""
        if self.var_teach_on.get() and self.teach:
            self.teach.set_rule(raw, self._teach_fields(li))
            if self._teach_write_ok():
                taught_msg = f" Taught ({len(self.teach.rules)} rules)."
        self._rebuild_review()
        for idx, (o, item) in enumerate(self.pairs):
            if item is li:
                self._select_pair(idx)
                break
        self._update_tree_counts(owner)
        self.status_var.set(f"Added item from source line {li.line_no + 1}.{taught_msg}")

    def _update_tree_counts(self, owner):
        if owner is None:
            return
        iid = self.res_to_iid.get(id(owner))
        if iid:
            self.tree.item(iid, values=(owner.fmt, "OK" if owner.ok else "ERR",
                                        len(owner.line_items), len(owner.warnings)))


# ------------------------------------------------------------------ smoke
def run_smoke() -> int:
    """Headless end-to-end test: renders a REAL line-numbered PO table image
    (the exact failure pattern from Steve's Tri-State test PO), OCRs it
    through the GUI, and asserts correct quantities, the on-image overlay
    rectangle, the new tab structure, column order, and teach persistence."""
    tmpdir = tempfile.mkdtemp(prefix="omniscan_gui_smoke_")
    teach_path = os.path.join(tmpdir, "teach.json")
    app = OmniScanGUI()
    app.var_teach.set(teach_path)
    app._load_teach_store()

    ocr_ready = False
    try:
        omniscan.DEPS.pytesseract.get_tesseract_version()
        ocr_ready = _PIL_OK
    except Exception:
        pass

    files = []
    if ocr_ready:
        from PIL import Image, ImageDraw, ImageFont
        rows = ["1  300 FT THREADED ROD ATR-056",
                "2  500 EA HEX NUTS HN-050",
                "3  250 EA FLAT WASHER FW-050"]
        img = Image.new("L", (1500, 90 + 90 * len(rows)), 255)
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.load_default(size=44)
        except TypeError:
            font = None
        for i, row in enumerate(rows):
            draw.text((30, 40 + 90 * i), row, font=font, fill=0)
        if font is None:
            img = img.resize((3000, img.height * 2))
        p_img = os.path.join(tmpdir, "test_po.png")
        img.save(p_img)
        files.append(p_img)
    p_docx = os.path.join(tmpdir, "rfq.docx")
    with open(p_docx, "wb") as fh:
        fh.write(omniscan._build_docx(
            ["RFQ 4512", "6 EA TEST-PART-1 WIDGET",
             "wire assembly per spec sheet A"]))
    files.append(p_docx)
    app._add_paths(files)

    outcome = {"code": 1}
    phase = {"n": 0, "ok": True}
    deadline = [150]

    def ok_and(cond, label):
        if not cond:
            phase["ok"] = False
            print(f"  smoke-fail: {label}")

    def watch():
        deadline[0] -= 1
        if deadline[0] <= 0:
            print("SMOKE FAIL: timeout")
            app.destroy()
            return
        if app.scanning or not app.results:
            app.after(200, watch)
            return
        if phase["n"] == 0:
            phase["n"] = 1
            tabs = [app.notebook.tab(t, "text") for t in app.notebook.tabs()]
            ok_and(tabs == ["Preview", "Extracted / Merged Text", "Warnings & Meta"],
                   f"tabs {tabs}")
            ok_and(_OUT_HEADER.index("PART") < _OUT_HEADER.index("DESCRIPTION")
                   < _OUT_HEADER.index("GAUGE"), "column order")
            if ocr_ready:
                # select the OCR image result
                img_iid = next(iid for iid, r in app.node_map.items()
                               if r.fmt == "image")
                app.tree.selection_set(img_iid)
                app._rebuild_review()
                qtys = {li.qty for li in app.node_map[img_iid].line_items}
                ok_and(qtys == {300.0, 500.0, 250.0},
                       f"OCR table qtys (line numbers stripped): {qtys}")
                ok_and(len(app.pv_pages) == 1 and len(app.pv_images) == 1,
                       "document preview rendered")
                app._select_pair(0)
                ok_and(bool(app.pv_canvas.find_withtag("ibox")),
                       "on-image overlay rectangle drawn")
                ok_and(bool(app.out.tag_ranges("box")), "output box drawn")
            else:
                print("  SKIP  OCR-image phase (tesseract not installed)")
            # teach round-trip on the docx
            docx_iid = next(iid for iid, r in app.node_map.items()
                            if r.fmt == "docx")
            app.tree.selection_set(docx_iid)
            app._rebuild_review()
            app._select_pair(0)
            app.var_e_qty.set("99")
            app.var_e_part.set("FIXED-1")
            app._on_apply()
            app.picked_line = 2
            app._on_add_from_line()
            rules = json.load(open(teach_path))
            ok_and(rules.get("6 EA TEST-PART-1 WIDGET", {}).get("qty") == 99
                   and "WIRE ASSEMBLY PER SPEC SHEET A" in rules,
                   f"teach rules written: {list(rules)}")
            app._start_scan()
            app.after(200, watch)
            return
        docx_res = next(r for r in app.results if r.fmt == "docx")
        taught_fix = any(li.taught and li.qty == 99 and li.part == "FIXED-1"
                         for li in docx_res.line_items)
        taught_add = any(li.taught and "WIRE ASSEMBLY" in (li.raw or "").upper()
                         for li in docx_res.line_items)
        ok_and(taught_fix, "taught fix applied on rescan")
        ok_and(taught_add, "taught add applied on rescan")
        print(f"SMOKE {'PASS' if phase['ok'] else 'FAIL'}")
        outcome["code"] = 0 if phase["ok"] else 1
        app.destroy()

    app._start_scan()
    app.after(200, watch)
    app.mainloop()
    return outcome["code"]


def main() -> int:
    if "--smoke" in sys.argv:
        return run_smoke()
    app = OmniScanGUI()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
