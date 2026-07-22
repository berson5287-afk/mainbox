#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OmniScan — universal document scanner for RFQ/procurement workflows.

Standalone replacement candidate for MaINbox SmartScan Extractor.
Scans "almost anything": PDF (text-layer + OCR fallback), DOCX, XLSX,
CSV/TSV, TXT, HTML, EML (with recursive attachment scanning), MSG
(optional), images (OCR and/or Ollama vision, both opt-in-capable),
and ZIP archives (recursive).

Design goals:
  * Zero required third-party dependencies. Everything degrades
    gracefully: each extractor probes for its optional deps at runtime
    and reports precisely what is missing instead of crashing.
  * Importable API (scan_path / scan_bytes -> ScanResult) AND a CLI.
  * RFQ-aware line-item parser with confidence scoring, wire-gauge
    awareness, and line-number-mistaken-for-qty protection.
  * Optional catalog grounding against a SQLite product catalog
    (schema auto-detected; tuned for american_power_catalog.db but
    generic).
  * Built-in --selftest that synthesizes real files for every
    stdlib-only format (plus PDF) and verifies end-to-end behavior.

Optional dependencies (all probed, none required, none bundled):
  pypdf        (BSD-3-Clause)  -> PDF text-layer extraction
  pdfplumber   (MIT)           -> better PDF table/text extraction
  pypdfium2    (Apache-2.0/BSD-3; PDFium: BSD-3) -> PDF raster for OCR fallback
  pytesseract  (Apache-2.0)    -> OCR (Tesseract engine itself: Apache-2.0)
  Pillow       (MIT-CMU)       -> image handling
Outlook .msg files are read NATIVELY — a from-scratch [MS-CFB] compound-file
reader + [MS-OXMSG] property parser, pure stdlib, no third-party code.

LICENSE (this file): MIT
Copyright (c) 2026 Stephen Berson

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

THIRD-PARTY NOTICE: this program bundles no third-party code. The optional
libraries above are imported only if the end user has installed them, and all
carry permissive licenses. No GPL or AGPL dependencies are referenced.
(Not legal advice — verify licensing questions with a professional.)
"""

APP_TITLE = "OmniScan v0.10.1"  # v0.10.1: review rows require a digit (kills 'noted'/title debris rows); voltage notation (Y/120, 480-208Y/120) excluded from part candidates | v0.10.0: hard-scan pass tuned on a faint 2-column handwritten RFQ (15/50 -> 46/50 auto + 4 surfaced) — dark-ink gutter profiling so gray gutter shading can't hide the column split; gap-based logical-line splitting from word boxes kills margin-note fusion (see spec/not pan); 'Qty N / UOM -' grammar; marker+paren combos ('19) -(3 box)'); marker-glued decimals resolved by expected line number (11.100 FT = marker 11 + qty 100, fractional FT preserved); glued lone-f (100f); tea/lea = 1 ea; leading quote/star debris; letter'digit unglue; U-for-0 before apostrophes; two-leading-numbers rule; IGNORE/OLD COUNT voiding; trailing junk stripped from descriptions | v0.9.2: handwritten-ft repairs — f/fi/fl alias to FT; lone i/l stripped before digits at desc start; missing unit on wire-gauge lines defaults FT instead of EA | v0.9.1: thousands-separated qtys (2,000 EA) protected from the glued-marker split — comma + 3-digit groups is a thousands separator, never marker+qty | v0.9.0: audit pass tuned on a hard 2-column handwritten RFQ — two-column page detection + pre-OCR splitting (background-floor-relative gutter profile); per-page OCR fallback so mixed text+scan PDFs work; crossed-out/VOID/scratch decoy suppression surfaced in meta.voided_lines; Windows tesseract.exe auto-probing (shutil.which + install dirs); fuzzy teach-rule matching (difflib, stricter for suppress rules); shredded-line part salvage as review rows; bare apostrophe-feet (200 ft); paren-qty after list markers; period-glued UOMs (2.CAN); semicolon/exclamation marker garbles; EFA/EQ aliases; catalog I/1 confusion fallback | v0.8.x: v0.8.2: handwritten-ea glue — "6a"/"1a" at the qty position (lowercase only) reads as "6 ea"/"1 ea"; uppercase amp ratings (20A/30A) explicitly protected | v0.8.1: items with a quantity but no detected unit now default to EA (Config.default_uom, '' disables) — review rows without a qty stay blank | v0.8.0: multi-format pass tuned on a real photocopied material request (1/62 -> 60/62 rows) — parenthesized-quantity grammar ('-(3 boxes) desc', apostrophe-feet '(400\\')', pack sizes '(2 cases/100pcs.)', garbled-paren recovery, qty-less review rows) with wrapped-continuation-line merging; highlighter-stroke removal before OCR (HSV desaturation) recovers marker-buried lines; explicit Part# extraction; catalog O/0 confusion fallback; plural UOMs + COIL/CAN/DZ/BKT; '2x' multiplier and 'QTY:' label formats | v0.7.0: teach persistence surfacing + desc normalization | v0.6.0: marker-column mode (handwriting 7/18 -> 18/18) | v0.5.0: deskew | v0.4.0: accuracy + OCR boxes | v0.3.0: teach layer | v0.2.x: native .msg; licenses | v0.1.0: initial

import argparse
import base64
import csv
import email
import email.policy
import html.parser
import io
import json
import os
import re
import sqlite3
import struct
import sys
import zipfile
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, asdict
from typing import Optional, List, Dict, Any, Tuple

# --------------------------------------------------------------------------
# Optional dependency probing
# --------------------------------------------------------------------------

class Deps:
    """Runtime probe of optional third-party libraries."""

    _TESSERACT_CANDIDATES = (
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
        "/usr/local/bin/tesseract",
        "/opt/homebrew/bin/tesseract",
    )

    def __init__(self):
        self.pypdf = self._try("pypdf")
        self.pdfplumber = self._try("pdfplumber")
        self.pypdfium2 = self._try("pypdfium2")
        self.pytesseract = self._try("pytesseract")
        self.PIL = self._try("PIL.Image")
        self.tesseract_cmd: Optional[str] = None
        self.tesseract_ok = False
        if self.pytesseract is not None:
            self._probe_tesseract()

    def _probe_tesseract(self):
        """pytesseract only works if the tesseract BINARY is reachable. On
        Windows the installer does not add it to PATH, so probe shutil.which
        and the common install locations and point pytesseract at it."""
        import shutil
        try:
            self.pytesseract.get_tesseract_version()
            self.tesseract_ok = True
            self.tesseract_cmd = "PATH"
            return
        except Exception:
            pass
        candidates = [shutil.which("tesseract")]
        candidates += [os.path.expandvars(p) for p in self._TESSERACT_CANDIDATES]
        candidates.append(os.path.expandvars(
            r"%LOCALAPPDATA%\\Programs\\Tesseract-OCR\\tesseract.exe"))
        for cand in candidates:
            if not cand or not os.path.exists(cand):
                continue
            try:
                self.pytesseract.pytesseract.tesseract_cmd = cand
                self.pytesseract.get_tesseract_version()
                self.tesseract_ok = True
                self.tesseract_cmd = cand
                return
            except Exception:
                continue

    @staticmethod
    def _try(name):
        try:
            module = __import__(name)
            for part in name.split(".")[1:]:
                module = getattr(module, part)
            return module
        except Exception:
            return None

    def report(self) -> str:
        rows = [
            ("pypdf (PDF text layer) [BSD-3]", self.pypdf),
            ("pdfplumber (PDF tables) [MIT]", self.pdfplumber),
            ("pypdfium2 (PDF raster for OCR) [Apache-2.0/BSD-3]", self.pypdfium2),
            ("pytesseract (OCR) [Apache-2.0]", self.pytesseract),
            ("Pillow (image handling) [MIT-CMU]", self.PIL),
        ]
        lines = [f"{APP_TITLE} — optional dependency status:"]
        for label, mod in rows:
            lines.append(f"  [{'OK ' if mod else '-- '}] {label}")
        binary = ("PATH" if self.tesseract_cmd == "PATH" else self.tesseract_cmd) \
            if self.tesseract_ok else "NOT FOUND (install tesseract or add to PATH)"
        lines.append(f"  [{'OK ' if self.tesseract_ok else '-- '}] tesseract binary: {binary}")
        lines.append("  [OK ] Outlook .msg — native from-scratch CFB reader, no dependency")
        lines.append("  (all optional deps are permissive-licensed; missing ones disable only their own format)")
        return "\n".join(lines)


DEPS = Deps()

# --------------------------------------------------------------------------
# Config + result dataclasses
# --------------------------------------------------------------------------

@dataclass
class Config:
    ocr: bool = True                    # allow OCR when deps present
    ollama_host: Optional[str] = None   # e.g. "http://tillium-bridge:11434" (opt-in)
    ollama_model: str = "gemma3"        # vision-capable model tag on that host
    ollama_timeout: int = 120
    catalog_path: Optional[str] = None  # e.g. american_power_catalog.db
    max_depth: int = 3                  # recursion for attachments/archives
    max_member_bytes: int = 25 * 1024 * 1024
    max_text_chars: int = 400_000
    min_item_confidence: float = 0.45
    default_uom: str = "EA"             # UOM when none is detected and qty
                                        # exists ("" disables the default)
    ocr_dpi: int = 300
    ocr_psm: int = 6                    # tesseract PSM; 6=uniform block, keeps
                                        # table ROWS together (default 3 splits
                                        # tables into column fragments)
    deskew: bool = True                 # auto-straighten slanted scans before OCR
    clean_highlights: bool = True       # white-out highlighter strokes before OCR
    split_columns: bool = True          # detect + split two-column pages before OCR
    teach_fuzzy: float = 0.86           # fuzzy teach-rule match ratio (0 disables)
    teach_path: Optional[str] = None    # JSON of learned line rules (see TeachStore)


@dataclass
class LineItem:
    raw: str
    qty: Optional[float] = None
    uom: Optional[str] = None
    part: Optional[str] = None
    description: Optional[str] = None
    gauge: Optional[str] = None
    confidence: float = 0.0
    catalog_match: Optional[Dict[str, Any]] = None
    line_no: Optional[int] = None    # 0-based index into the source text's lines
    corrected: bool = False          # True once a human verified/edited this item
    taught: bool = False             # True when produced by a learned teach rule

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ScanResult:
    name: str
    fmt: str
    ok: bool = True
    text: str = ""
    line_items: List[LineItem] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    children: List["ScanResult"] = field(default_factory=list)
    meta: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "fmt": self.fmt,
            "ok": self.ok,
            "text": self.text,
            "line_items": [li.to_dict() for li in self.line_items],
            "warnings": self.warnings,
            "meta": self.meta,
            "children": [c.to_dict() for c in self.children],
        }

    def all_line_items(self) -> List[LineItem]:
        """Flattened line items from this result and all children."""
        items = list(self.line_items)
        for child in self.children:
            items.extend(child.all_line_items())
        return items


# --------------------------------------------------------------------------
# Format detection
# --------------------------------------------------------------------------

_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}


def detect_format(data: bytes, name: str) -> str:
    ext = os.path.splitext(name or "")[1].lower()
    head = data[:8]

    if head.startswith(b"%PDF"):
        return "pdf"
    if head.startswith(b"PK\x03\x04"):
        # OOXML vs plain zip — decide by member names
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = set(zf.namelist())
            if "word/document.xml" in names:
                return "docx"
            if "xl/workbook.xml" in names:
                return "xlsx"
            if "ppt/presentation.xml" in names:
                return "pptx-zip"   # scanned as generic zip (slide XML text)
        except Exception:
            pass
        return "zip"
    if head.startswith(b"\x89PNG") or head.startswith(b"\xff\xd8") or head.startswith(b"GIF8") or head.startswith(b"BM"):
        return "image"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        # OLE/CFB compound file: Outlook .msg or legacy .doc/.xls — the native
        # msg extractor decides which and degrades with a warning.
        return "ole"

    if ext in _IMAGE_EXTS:
        return "image"
    if ext == ".eml":
        return "eml"
    if ext in (".csv", ".tsv"):
        return "csv"
    if ext in (".htm", ".html"):
        return "html"
    if ext == ".msg":
        return "ole"
    if ext == ".pdf":
        return "pdf"
    if ext == ".docx":
        return "docx"
    if ext == ".xlsx":
        return "xlsx"
    if ext == ".zip":
        return "zip"

    # Content sniff for text-ish payloads
    sample = data[:4096]
    if b"<html" in sample.lower() or b"<!doctype html" in sample.lower():
        return "html"
    if sample.startswith(b"From:") or b"\nSubject:" in sample[:2048] or sample.startswith(b"Return-Path:"):
        return "eml"
    try:
        sample.decode("utf-8")
        return "text"
    except UnicodeDecodeError:
        pass
    # latin-1 always decodes, so gate on printable ratio instead
    if sample:
        printable = sum(1 for b in sample if 32 <= b < 127 or b in (9, 10, 13))
        if printable / len(sample) >= 0.85:
            return "text"
    return "binary"


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _decode_text(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _col_letters_to_index(ref: str) -> int:
    """'A' -> 0, 'B' -> 1, 'AA' -> 26. Accepts full cell refs like 'B12'."""
    letters = "".join(ch for ch in ref if ch.isalpha()).upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(idx - 1, 0)


class _HTMLTextParser(html.parser.HTMLParser):
    _SKIP = {"script", "style", "head"}
    _BREAKERS = {"p", "div", "br", "tr", "li", "table", "h1", "h2", "h3", "h4", "h5", "h6"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._chunks: List[str] = []
        self._skip_depth = 0
        self._cell_break = False

    def handle_starttag(self, tag, attrs):
        if tag in self._SKIP:
            self._skip_depth += 1
        elif tag in self._BREAKERS:
            self._chunks.append("\n")
        elif tag in ("td", "th"):
            self._chunks.append("\t")

    def handle_endtag(self, tag):
        if tag in self._SKIP and self._skip_depth > 0:
            self._skip_depth -= 1
        elif tag in self._BREAKERS:
            self._chunks.append("\n")

    def handle_data(self, data):
        if self._skip_depth == 0:
            self._chunks.append(data)

    def text(self) -> str:
        raw = "".join(self._chunks)
        lines = [re.sub(r"[ \t]+", " ", ln).strip("\t ") for ln in raw.splitlines()]
        out = []
        blank = 0
        for ln in lines:
            if ln:
                out.append(ln)
                blank = 0
            else:
                blank += 1
                if blank == 1:
                    out.append("")
        return "\n".join(out).strip()


# --------------------------------------------------------------------------
# Extractors — each returns (text, warnings, meta)
# --------------------------------------------------------------------------

def extract_text_plain(data: bytes) -> Tuple[str, List[str], Dict[str, Any]]:
    return _decode_text(data), [], {}


def extract_csv(data: bytes, name: str) -> Tuple[str, List[str], Dict[str, Any]]:
    text = _decode_text(data)
    delim = "\t" if (name.lower().endswith(".tsv") or ("\t" in text.splitlines()[0] if text.splitlines() else False)) else ","
    try:
        rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    except csv.Error as exc:
        return text, [f"csv parse fallback to raw text: {exc}"], {}
    rendered = "\n".join("\t".join(cell.strip() for cell in row) for row in rows)
    return rendered, [], {"rows": len(rows), "delimiter": delim}


def extract_html(data: bytes) -> Tuple[str, List[str], Dict[str, Any]]:
    parser = _HTMLTextParser()
    try:
        parser.feed(_decode_text(data))
        parser.close()
    except Exception as exc:
        return _decode_text(data), [f"html parse fallback to raw text: {exc}"], {}
    return parser.text(), [], {}


def extract_docx(data: bytes) -> Tuple[str, List[str], Dict[str, Any]]:
    warnings: List[str] = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            xml_data = zf.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        return "", [f"docx: cannot read word/document.xml: {exc}"], {}

    try:
        root = ET.fromstring(xml_data)
    except ET.ParseError as exc:
        return "", [f"docx: XML parse error: {exc}"], {}

    body = None
    for child in root:
        if _localname(child.tag) == "body":
            body = child
            break
    if body is None:
        return "", ["docx: no <body> element"], {}

    def para_text(p_el) -> str:
        parts = []
        for el in p_el.iter():
            ln = _localname(el.tag)
            if ln == "t" and el.text:
                parts.append(el.text)
            elif ln in ("br", "cr"):
                parts.append("\n")
            elif ln == "tab":
                parts.append("\t")
        return "".join(parts)

    lines: List[str] = []
    tables = 0
    for child in body:
        ln = _localname(child.tag)
        if ln == "p":
            lines.append(para_text(child))
        elif ln == "tbl":
            tables += 1
            for tr in child:
                if _localname(tr.tag) != "tr":
                    continue
                cells = []
                for tc in tr:
                    if _localname(tc.tag) != "tc":
                        continue
                    cell_paras = [para_text(p) for p in tc if _localname(p.tag) == "p"]
                    cells.append(" ".join(cp for cp in cell_paras if cp))
                lines.append("\t".join(cells))
    return "\n".join(lines).strip(), warnings, {"tables": tables}


def extract_xlsx(data: bytes) -> Tuple[str, List[str], Dict[str, Any]]:
    warnings: List[str] = []
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        return "", [f"xlsx: bad zip: {exc}"], {}

    with zf:
        names = zf.namelist()

        # Shared strings
        shared: List[str] = []
        if "xl/sharedStrings.xml" in names:
            try:
                root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in root:
                    if _localname(si.tag) != "si":
                        continue
                    text_parts = [el.text for el in si.iter() if _localname(el.tag) == "t" and el.text]
                    shared.append("".join(text_parts))
            except ET.ParseError as exc:
                warnings.append(f"xlsx: sharedStrings parse error: {exc}")

        sheet_names = sorted(n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml"))
        out_lines: List[str] = []
        sheets = 0
        for sheet in sheet_names:
            try:
                root = ET.fromstring(zf.read(sheet))
            except ET.ParseError as exc:
                warnings.append(f"xlsx: {sheet} parse error: {exc}")
                continue
            sheets += 1
            if len(sheet_names) > 1:
                out_lines.append(f"[sheet: {os.path.basename(sheet)}]")
            for row_el in root.iter():
                if _localname(row_el.tag) != "row":
                    continue
                cells: Dict[int, str] = {}
                pos = 0
                for c_el in row_el:
                    if _localname(c_el.tag) != "c":
                        continue
                    ref = c_el.get("r", "")
                    idx = _col_letters_to_index(ref) if ref else pos
                    pos = idx + 1
                    ctype = c_el.get("t", "n")
                    value = ""
                    v_el = None
                    is_el = None
                    for sub in c_el:
                        subln = _localname(sub.tag)
                        if subln == "v":
                            v_el = sub
                        elif subln == "is":
                            is_el = sub
                    if ctype == "s" and v_el is not None and v_el.text is not None:
                        try:
                            value = shared[int(v_el.text)]
                        except (ValueError, IndexError):
                            value = v_el.text
                    elif ctype == "inlineStr" and is_el is not None:
                        value = "".join(el.text for el in is_el.iter()
                                        if _localname(el.tag) == "t" and el.text)
                    elif v_el is not None and v_el.text is not None:
                        value = v_el.text
                    cells[idx] = value
                if cells:
                    width = max(cells) + 1
                    out_lines.append("\t".join(cells.get(i, "") for i in range(width)))
        return "\n".join(out_lines).strip(), warnings, {"sheets": sheets, "shared_strings": len(shared)}


def extract_pdf(data: bytes, cfg: Config) -> Tuple[str, List[str], Dict[str, Any]]:
    warnings: List[str] = []
    meta: Dict[str, Any] = {}
    page_texts: List[str] = []
    engine = None

    if DEPS.pdfplumber is not None:
        try:
            with DEPS.pdfplumber.open(io.BytesIO(data)) as pdf:
                page_texts = [(page.extract_text() or "") for page in pdf.pages]
            engine = "pdfplumber"
        except Exception as exc:
            warnings.append(f"pdf: pdfplumber failed ({exc}); trying pypdf")
            page_texts = []
    if not page_texts and DEPS.pypdf is not None:
        try:
            reader = DEPS.pypdf.PdfReader(io.BytesIO(data))
            page_texts = [(page.extract_text() or "") for page in reader.pages]
            engine = "pypdf"
        except Exception as exc:
            warnings.append(f"pdf: pypdf failed: {exc}")
    if DEPS.pypdf is None and DEPS.pdfplumber is None:
        warnings.append("pdf: no PDF library installed (pip install pypdf) — text layer unread")

    meta["pages"] = len(page_texts)

    # Per-page OCR fallback: any page with essentially no text layer gets
    # rasterized and OCR'd individually, so mixed text+scan PDFs work.
    scan_pages = [i for i, t in enumerate(page_texts) if len(t.strip()) < 30]
    if not page_texts and DEPS.pypdfium2 is not None:
        try:
            scan_pages = list(range(len(DEPS.pypdfium2.PdfDocument(data))))
            page_texts = [""] * len(scan_pages)
            meta["pages"] = len(scan_pages)
        except Exception:
            scan_pages = []

    ocr_done = 0
    if scan_pages:
        if not cfg.ocr:
            warnings.append(f"pdf: {len(scan_pages)} page(s) have no text layer "
                            "and OCR is disabled by config")
        elif (DEPS.pypdfium2 is not None and DEPS.pytesseract is not None
                and DEPS.PIL is not None):
            try:
                pdf = DEPS.pypdfium2.PdfDocument(data)
                ocr_lines: Dict[str, Any] = {}
                ocr_pages: List[List[int]] = []
                deskew_degs: List[float] = [0.0] * len(page_texts)
                columns = 1
                for i in scan_pages:
                    bitmap = pdf[i].render(scale=cfg.ocr_dpi / 72.0)
                    pil_img = bitmap.to_pil()
                    text_i, boxes, deg, ncols = _ocr_page(pil_img, cfg)
                    page_texts[i] = text_i
                    deskew_degs[i] = round(deg, 2)
                    columns = max(columns, ncols)
                    # box page-size must reflect the deskewed (expanded) image
                    if deg:
                        pil_sz = pil_img.rotate(deg, expand=True).size
                    else:
                        pil_sz = pil_img.size
                    while len(ocr_pages) < i:
                        ocr_pages.append([0, 0])
                    ocr_pages.append(list(pil_sz))
                    if boxes:
                        for j, box in enumerate(boxes):
                            ocr_lines[f"__page{i}__{j}"] = {"page": i, "box": box}
                    ocr_done += 1
                pdf.close()
                # remap ocr_lines keys to global line indices now that all
                # page texts are final
                if ocr_lines:
                    remapped: Dict[str, Any] = {}
                    offset = 0
                    for i, t in enumerate(page_texts):
                        n_lines = max(len(t.splitlines()), 1)
                        for j in range(n_lines):
                            info = ocr_lines.get(f"__page{i}__{j}")
                            if info:
                                remapped[str(offset + j)] = info
                        offset += n_lines
                    meta["ocr_lines"] = remapped
                    meta["ocr_pages"] = ocr_pages
                meta["deskew_deg"] = deskew_degs
                if columns > 1:
                    meta["columns"] = columns
            except Exception as exc:
                warnings.append(f"pdf: OCR fallback failed: {exc}")
        else:
            missing = [n for n, m in (("pypdfium2", DEPS.pypdfium2),
                                      ("pytesseract", DEPS.pytesseract),
                                      ("Pillow", DEPS.PIL)) if m is None]
            warnings.append(f"pdf: {len(scan_pages)} page(s) look scanned; "
                            "OCR needs: " + ", ".join(missing))

    if ocr_done and engine and ocr_done < len(page_texts):
        meta["engine"] = f"mixed({engine}+ocr)"
    elif ocr_done:
        meta["engine"] = "ocr(pdfium+tesseract)"
    elif engine:
        meta["engine"] = engine

    text = "\n".join(page_texts)
    return text, warnings, meta


def _deskew_angle(img) -> float:
    """Estimate the skew angle (degrees, PIL counterclockwise-positive) of a
    document image using projection-profile variance: text lines aligned with
    the raster maximize the variance of per-row darkness. Pure PIL — row sums
    come from resizing the rotated image to a 1-pixel-wide column (BOX filter
    averages each row). Returns 0.0 when the page already looks straight."""
    import statistics
    small = img.convert("L")
    if small.height > 800:
        ratio = 800 / small.height
        small = small.resize((max(1, int(small.width * ratio)), 800))
    small = small.point(lambda p: 0 if p < 200 else 255)

    def score(angle: float) -> float:
        r = small.rotate(angle, expand=False, fillcolor=255)
        col = r.resize((1, r.height), DEPS.PIL.BOX)
        getter = getattr(col, "get_flattened_data", None) or col.getdata
        rows = list(getter())
        return statistics.pvariance(rows) if len(rows) > 1 else 0.0

    base = score(0.0)
    best_ang, best = 0.0, base
    for tenths in range(-80, 81, 10):                    # coarse: -8..8 by 1.0
        ang = tenths / 10.0
        if ang == 0.0:
            continue
        s = score(ang)
        if s > best:
            best_ang, best = ang, s
    center = best_ang
    for tenths in range(-9, 10, 2):                      # fine: ±0.9 by 0.2
        ang = center + tenths / 10.0
        s = score(ang)
        if s > best:
            best_ang, best = ang, s
    if abs(best_ang) < 0.4 or best < base * 1.10:
        return 0.0
    return best_ang


def _maybe_deskew(img, cfg: Config):
    """Returns (possibly rotated image, applied angle in degrees)."""
    if not cfg.deskew:
        return img, 0.0
    try:
        angle = _deskew_angle(img)
    except Exception:
        return img, 0.0
    if not angle:
        return img, 0.0
    fill = 255 if img.mode == "L" else (255, 255, 255)
    return img.rotate(angle, expand=True, fillcolor=fill), angle


def _remove_highlights(img):
    """Neutralize highlighter marker strokes before OCR: pixels that are both
    saturated and bright (yellow/pink/green marker) go white; printed black
    text (low saturation) is untouched. Highlighted lines otherwise come out
    garbled or vanish entirely from tesseract."""
    if img.mode != "RGB":
        return img
    try:
        import PIL.ImageChops as ImageChops
        hsv = img.convert("HSV")
        _h, s, v = hsv.split()
        s_mask = s.point(lambda p: 255 if p > 80 else 0)
        v_mask = v.point(lambda p: 255 if p > 100 else 0)
        mask = ImageChops.multiply(s_mask, v_mask)
        white = DEPS.PIL.new("RGB", img.size, (255, 255, 255))
        return DEPS.PIL.composite(white, img, mask)
    except Exception:
        return img


def _find_column_split(img) -> Optional[int]:
    """Detect a two-column page: a tall, mostly-white vertical gutter near the
    middle with substantial ink on BOTH sides. A thin vertical rule inside the
    gutter is tolerated (1-D median filter). Returns the split x in pixels,
    or None for single-column pages."""
    try:
        g = img.convert("L")
        # estimate paper background (median pixel), then keep ONLY truly dark
        # ink: light-gray gutter shading or scanner banding must count as
        # zero, or the gutter disappears from the profile
        hist = g.histogram()
        total = sum(hist)
        acc, bg = 0, 245
        for v, c in enumerate(hist):
            acc += c
            if acc >= total // 2:
                bg = v
                break
        dark_thr = max(100, bg - 70)
        dark = g.point(lambda p: 255 if p < dark_thr else 0)
        bins = 240
        col = dark.resize((bins, 1), DEPS.PIL.BOX)
        getter = getattr(col, "get_flattened_data", None) or col.getdata
        rel = list(getter())                   # per-column dark-pixel density
        med = []
        for i in range(bins):
            window = sorted(rel[max(0, i - 2):i + 3])
            med.append(window[len(window) // 2])
        rel = med
        mean_rel = sum(rel) / bins
        if mean_rel < 0.5:
            return None
        thresh = max(1.0, 0.15 * mean_rel)
        lo, hi = int(bins * 0.30), int(bins * 0.70)
        best_len, best_mid = 0, None
        i = lo
        while i < hi:
            if rel[i] <= thresh:
                j = i
                while j < hi and rel[j] <= thresh:
                    j += 1
                if j - i > best_len:
                    best_len, best_mid = j - i, (i + j) // 2
                i = j
            else:
                i += 1
        if best_mid is None or best_len < bins * 0.04:
            return None
        left_ink = sum(rel[:best_mid])
        right_ink = sum(rel[best_mid:])
        if min(left_ink, right_ink) < 0.15 * max(left_ink, right_ink):
            return None                     # one side is basically empty margin
        return int(best_mid / bins * img.width)
    except Exception:
        return None


def _ocr_page(pil_img, cfg: "Config"):
    """Full per-page OCR pipeline: highlight cleanup, deskew, optional
    two-column split, confidence-gated retry. Returns (text, boxes, deg,
    n_columns); boxes are in the (possibly deskewed) full-page pixel space."""
    if cfg.clean_highlights:
        pil_img = _remove_highlights(pil_img.convert("RGB")
                                     if pil_img.mode != "RGB" else pil_img)
    pil_img, deg = _maybe_deskew(pil_img, cfg)
    split_x = _find_column_split(pil_img) if cfg.split_columns else None
    if split_x:
        w, h = pil_img.size
        left = pil_img.crop((0, 0, split_x, h))
        right = pil_img.crop((split_x, 0, w, h))
        t1, b1, _p1 = _ocr_with_retry(left, cfg.ocr_psm, cfg)
        t2, b2, _p2 = _ocr_with_retry(right, cfg.ocr_psm, cfg)
        t1, t2 = t1.strip("\n"), t2.strip("\n")
        if b2:
            b2 = [[x0 + split_x, y0, x1 + split_x, y1] for x0, y0, x1, y1 in b2]
        if t1 and t2:
            boxes = (b1 + b2) if (b1 is not None and b2 is not None) else None
            return t1 + "\n" + t2, boxes, deg, 2
        if t1 or t2:
            return (t1 or t2), (b1 if t1 else b2), deg, 2
    text, boxes, _pass = _ocr_with_retry(pil_img, cfg.ocr_psm, cfg)
    return text, boxes, deg, 1


def _ocr_quality(text: str) -> float:
    """Cheap OCR usefulness score: lines containing both letters and digits
    (item-like) weigh most, plus a capped character count."""
    item_lines = sum(1 for l in text.splitlines()
                     if re.search(r"[A-Za-z]", l) and re.search(r"\d", l))
    return item_lines * 10 + min(len(re.findall(r"[A-Za-z0-9]", text)), 500) / 50


def _ocr_with_retry(img, psm: int, cfg: Optional["Config"] = None):
    """OCR once; accept immediately when tesseract's own word confidence is
    high. Otherwise walk a preprocessing ladder (denoise, contrast, binarize,
    upscale, re-deskew-after-denoise) and keep the best pass by word
    confidence. Clean documents never pay the retry cost."""
    text, boxes, conf = _ocr_image_lines(img, psm)
    quality = _ocr_quality(text)
    if conf >= 70 or (conf < 0 and quality >= 25):
        return text, boxes, "plain"
    best = (conf, quality, text, boxes, "plain")
    try:
        import PIL.ImageFilter as ImageFilter
        import PIL.ImageOps as ImageOps
        gray = img.convert("L")
        base = ImageOps.autocontrast(gray.filter(ImageFilter.MedianFilter(3)))
        variants = [("median+autocontrast", base, None),
                    ("median+binarize",
                     base.point(lambda p: 0 if p < 160 else 255), None),
                    ("median+upscale",
                     base.resize((int(base.width * 1.5),
                                  int(base.height * 1.5))), 1.5)]
        if cfg is None or cfg.deskew:
            try:
                angle = _deskew_angle(base)
            except Exception:
                angle = 0.0
            if angle:
                rot = base.rotate(angle, expand=True, fillcolor=255)
                variants.append(("median+redeskew", rot, "rotated"))
                variants.append(("median+redeskew+binarize",
                                 rot.point(lambda p: 0 if p < 160 else 255),
                                 "rotated"))
        for name, variant, scale in variants:
            t2, b2, c2 = _ocr_image_lines(variant, psm)
            q2 = _ocr_quality(t2)
            if (c2, q2) > (best[0], best[1]):
                if scale == "rotated":
                    b2 = None            # coords no longer map to the preview
                elif b2 and scale:
                    b2 = [[int(c / scale) for c in box] for box in b2]
                best = (c2, q2, t2, b2, name)
    except Exception:
        pass
    return best[2], best[3], best[4]


def _ocr_image_lines(img, psm: int = 6) -> Tuple[str, Optional[List[List[int]]], float]:
    """OCR an image into (text, per-line boxes). Boxes are [x0,y0,x1,y1] in
    image pixels, one per text line, aligned with text.splitlines(). Falls
    back to plain image_to_string (boxes=None) if word data is unavailable."""
    try:
        data = DEPS.pytesseract.image_to_data(
            img, config=f"--psm {psm}",
            output_type=DEPS.pytesseract.Output.DICT)
    except Exception:
        try:
            return (DEPS.pytesseract.image_to_string(img, config=f"--psm {psm}"),
                    None, -1.0)
        except Exception:
            return "", None, -1.0
    n = len(data.get("text", []))
    groups: Dict[tuple, Dict[str, Any]] = {}
    order: List[tuple] = []
    for i in range(n):
        word = (data["text"][i] or "").strip()
        if not word:
            continue
        key = (data.get("page_num", [1] * n)[i], data["block_num"][i],
               data["par_num"][i], data["line_num"][i])
        g = groups.get(key)
        if g is None:
            g = {"words": [], "x0": 10 ** 9, "y0": 10 ** 9, "x1": 0, "y1": 0}
            groups[key] = g
            order.append(key)
        x, y, w, h = (data["left"][i], data["top"][i],
                      data["width"][i], data["height"][i])
        g["words"].append((word, x, x + w, y, y + h))
        g["x0"] = min(g["x0"], x)
        g["y0"] = min(g["y0"], y)
        g["x1"] = max(g["x1"], x + w)
        g["y1"] = max(g["y1"], y + h)
    # split any tesseract line at huge horizontal gaps: margin annotations
    # ("see spec") and column leakage otherwise fuse into the item row
    gap_px = max(60, int(img.width * 0.12))
    lines, boxes = [], []
    for key in order:
        g = groups[key]
        seg: List[tuple] = []
        segments: List[List[tuple]] = []
        for wtup in g["words"]:
            if seg and wtup[1] - seg[-1][2] > gap_px:
                segments.append(seg)
                seg = []
            seg.append(wtup)
        if seg:
            segments.append(seg)
        for seg in segments:
            lines.append(" ".join(t[0] for t in seg))
            boxes.append([int(min(t[1] for t in seg)),
                          int(min(t[3] for t in seg)),
                          int(max(t[2] for t in seg)),
                          int(max(t[4] for t in seg))])
    confs = [int(c) for c in data.get("conf", []) if str(c).lstrip("-").isdigit()
             and int(c) >= 0]
    mean_conf = sum(confs) / len(confs) if confs else 0.0
    return "\n".join(lines), boxes, mean_conf


def extract_image(data: bytes, cfg: Config) -> Tuple[str, List[str], Dict[str, Any]]:
    warnings: List[str] = []
    meta: Dict[str, Any] = {}
    text = ""

    if cfg.ocr and DEPS.pytesseract is not None and DEPS.PIL is not None:
        try:
            img = DEPS.PIL.open(io.BytesIO(data))
            text, boxes, deskew_deg, ncols = _ocr_page(img, cfg)
            text = text.strip()
            meta["engine"] = "tesseract"
            meta["deskew_deg"] = [round(deskew_deg, 2)]
            if ncols > 1:
                meta["columns"] = ncols
            if boxes:
                if deskew_deg:
                    sz = img.convert("RGB").rotate(deskew_deg, expand=True).size
                else:
                    sz = img.size
                meta["ocr_lines"] = {str(i): {"page": 0, "box": b}
                                     for i, b in enumerate(boxes)}
                meta["ocr_pages"] = [list(sz)]
        except Exception as exc:
            warnings.append(f"image: OCR failed: {exc}")
    elif cfg.ocr:
        missing = [n for n, m in (("pytesseract", DEPS.pytesseract), ("Pillow", DEPS.PIL)) if m is None]
        warnings.append("image: OCR unavailable, missing: " + ", ".join(missing))

    if cfg.ollama_host:
        vision_text, vision_warns = _ollama_vision(data, cfg)
        warnings.extend(vision_warns)
        if vision_text:
            meta["ollama_model"] = cfg.ollama_model
            text = (text + "\n\n[vision]\n" + vision_text).strip() if text else vision_text

    if not text and not warnings:
        warnings.append("image: no OCR/vision path produced text")
    return text, warnings, meta


def _ollama_vision(image_bytes: bytes, cfg: Config) -> Tuple[str, List[str]]:
    """Opt-in vision transcription via a local Ollama host. Never raises."""
    payload = {
        "model": cfg.ollama_model,
        "prompt": ("Transcribe all text in this document image exactly, preserving line "
                   "breaks and any table structure using tab-separated columns. "
                   "Output only the transcription."),
        "images": [base64.b64encode(image_bytes).decode("ascii")],
        "stream": False,
    }
    url = cfg.ollama_host.rstrip("/") + "/api/generate"
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=cfg.ollama_timeout) as resp:
            body = json.loads(resp.read().decode("utf-8"))
        return (body.get("response") or "").strip(), []
    except (urllib.error.URLError, OSError, ValueError) as exc:
        return "", [f"ollama vision: request failed: {exc}"]


def extract_eml(data: bytes, cfg: Config, depth: int) -> Tuple[str, List[str], Dict[str, Any], List[ScanResult]]:
    warnings: List[str] = []
    children: List[ScanResult] = []
    try:
        msg = email.message_from_bytes(data, policy=email.policy.default)
    except Exception as exc:
        return _decode_text(data), [f"eml: parse failed, raw text fallback: {exc}"], {}, []

    header_lines = []
    for hdr in ("From", "To", "Cc", "Date", "Subject"):
        if msg.get(hdr):
            header_lines.append(f"{hdr}: {msg.get(hdr)}")

    body_plain: List[str] = []
    body_html: List[str] = []
    attachments = 0

    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        filename = part.get_filename()
        disposition = (part.get_content_disposition() or "").lower()
        try:
            payload = part.get_payload(decode=True)
        except Exception as exc:
            warnings.append(f"eml: undecodable part ({ctype}): {exc}")
            continue
        if payload is None:
            continue

        if filename or disposition == "attachment":
            attachments += 1
            if depth >= cfg.max_depth:
                warnings.append(f"eml: attachment '{filename}' skipped (max depth {cfg.max_depth})")
            elif len(payload) > cfg.max_member_bytes:
                warnings.append(f"eml: attachment '{filename}' skipped (> {cfg.max_member_bytes} bytes)")
            else:
                children.append(scan_bytes(payload, filename or f"attachment-{attachments}", cfg, depth + 1))
        elif ctype == "text/plain":
            body_plain.append(_decode_text(payload))
        elif ctype == "text/html":
            body_html.append(extract_html(payload)[0])

    body = "\n".join(body_plain).strip() or "\n".join(body_html).strip()
    text = ("\n".join(header_lines) + "\n\n" + body).strip()
    return text, warnings, {"attachments": attachments}, children


# --------------------------------------------------------------------------
# Native Outlook .msg support — from-scratch [MS-CFB] + [MS-OXMSG] readers.
# Pure stdlib. No third-party code, no GPL dependency.
# --------------------------------------------------------------------------

_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ENDOFCHAIN = 0xFFFFFFFE
_FREESECT = 0xFFFFFFFF
_FATSECT = 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF
_FILETIME_EPOCH_DELTA = 11644473600  # seconds between 1601-01-01 and 1970-01-01


@dataclass
class _CFBEntry:
    name: str
    etype: int      # 1=storage, 2=stream, 5=root
    left: int
    right: int
    child: int
    start: int
    size: int


class CFBFile:
    """Minimal from-scratch reader for the Compound File Binary format
    ([MS-CFB]): header, DIFAT, FAT, mini FAT, ministream, directory tree.
    Supports v3 (512-byte sectors) and v4 (4096-byte sectors)."""

    def __init__(self, data: bytes):
        if len(data) < 512 or data[:8] != _CFB_MAGIC:
            raise ValueError("not a CFB/OLE compound file")
        sector_shift, mini_shift = struct.unpack_from("<HH", data, 30)
        if sector_shift not in (9, 12) or mini_shift != 6:
            raise ValueError(f"CFB: unsupported sector shifts ({sector_shift},{mini_shift})")
        self.ssz = 1 << sector_shift
        self.msz = 1 << mini_shift
        (_num_fat, first_dir, _txn, self.mini_cutoff, first_minifat,
         num_minifat, first_difat, num_difat) = struct.unpack_from("<8I", data, 44)
        self.data = data

        # DIFAT: 109 entries in the header, then chained DIFAT sectors
        difat = list(struct.unpack_from("<109I", data, 76))
        sect = first_difat
        for _ in range(num_difat):
            if sect in (_ENDOFCHAIN, _FREESECT):
                break
            vals = struct.unpack(f"<{self.ssz // 4}I", self._sector(sect))
            difat.extend(vals[:-1])
            sect = vals[-1]

        # FAT
        self.fat: List[int] = []
        for fat_sect in difat:
            if fat_sect in (_ENDOFCHAIN, _FREESECT):
                continue
            self.fat.extend(struct.unpack(f"<{self.ssz // 4}I", self._sector(fat_sect)))

        # Directory entries (flat list; tree is rebuilt in paths())
        dir_raw = self._read_chain(first_dir)
        self.entries: List[_CFBEntry] = []
        for off in range(0, len(dir_raw), 128):
            blob = dir_raw[off:off + 128]
            if len(blob) < 128:
                break
            name_len = struct.unpack_from("<H", blob, 64)[0]
            name = (blob[:max(name_len - 2, 0)].decode("utf-16-le", errors="replace")
                    if 2 <= name_len <= 64 else "")
            etype = blob[66]
            left, right, child = struct.unpack_from("<3I", blob, 68)
            start = struct.unpack_from("<I", blob, 116)[0]
            size = struct.unpack_from("<Q", blob, 120)[0]
            if self.ssz == 512:
                size &= 0xFFFFFFFF  # v3: only the low 4 bytes are valid
            self.entries.append(_CFBEntry(name, etype, left, right, child, start, size))
        if not self.entries or self.entries[0].etype != 5:
            raise ValueError("CFB: missing root directory entry")

        # Mini FAT + ministream (small streams live inside the root's stream)
        mf_raw = self._read_chain(first_minifat) if num_minifat else b""
        self.minifat: List[int] = (list(struct.unpack(f"<{len(mf_raw) // 4}I", mf_raw))
                                   if mf_raw else [])
        root = self.entries[0]
        self.ministream = self._read_chain(root.start)[:root.size] if root.size else b""

    def _sector(self, n: int) -> bytes:
        off = (n + 1) * self.ssz
        return self.data[off:off + self.ssz]

    def _read_chain(self, start: int) -> bytes:
        out, sect, guard = [], start, 0
        while sect not in (_ENDOFCHAIN, _FREESECT) and 0 <= sect < len(self.fat):
            out.append(self._sector(sect))
            sect = self.fat[sect]
            guard += 1
            if guard > len(self.fat) + 4:  # corrupt-chain guard
                break
        return b"".join(out)

    def _read_minichain(self, start: int) -> bytes:
        out, sect, guard = [], start, 0
        while sect not in (_ENDOFCHAIN, _FREESECT) and 0 <= sect < len(self.minifat):
            off = sect * self.msz
            out.append(self.ministream[off:off + self.msz])
            sect = self.minifat[sect]
            guard += 1
            if guard > len(self.minifat) + 4:
                break
        return b"".join(out)

    def read(self, entry: _CFBEntry) -> bytes:
        if entry.etype not in (2,) or entry.size == 0:
            return b""
        raw = (self._read_minichain(entry.start) if entry.size < self.mini_cutoff
               else self._read_chain(entry.start))
        return raw[:entry.size]

    def paths(self) -> Dict[str, _CFBEntry]:
        """Rebuild the directory tree into {path: entry}. Storage paths end
        with '/'. Iterative + cycle-safe against corrupt sibling pointers."""
        out: Dict[str, _CFBEntry] = {}
        stack = [(self.entries[0].child, "")]
        seen = set()
        while stack:
            idx, prefix = stack.pop()
            if idx == _NOSTREAM or idx >= len(self.entries) or (idx, prefix) in seen:
                continue
            seen.add((idx, prefix))
            e = self.entries[idx]
            stack.append((e.left, prefix))
            stack.append((e.right, prefix))
            if e.etype == 1:
                out[prefix + e.name + "/"] = e
                stack.append((e.child, prefix + e.name + "/"))
            elif e.etype == 2:
                out[prefix + e.name] = e
        return out


class MsgFile:
    """From-scratch Outlook .msg reader ([MS-OXMSG]) on top of CFBFile.
    Reads unicode/ANSI string properties, binary properties, FILETIME dates
    from the fixed-length properties stream, and attachment storages."""

    def __init__(self, data: bytes):
        self.cfb = CFBFile(data)
        self.pathmap = self.cfb.paths()
        if not any(p.split("/")[-1].startswith("__substg1.0_")
                   or p.split("/")[-1].startswith("__properties")
                   for p in self.pathmap):
            raise ValueError("CFB file has no MAPI property streams (not a .msg)")

    def _stream(self, path: str) -> Optional[bytes]:
        entry = self.pathmap.get(path)
        return self.cfb.read(entry) if entry else None

    def prop_string(self, prop: str, prefix: str = "") -> str:
        raw = self._stream(f"{prefix}__substg1.0_{prop}001F")   # PT_UNICODE
        if raw is not None:
            return raw.decode("utf-16-le", errors="replace").rstrip("\x00")
        raw = self._stream(f"{prefix}__substg1.0_{prop}001E")   # PT_STRING8
        if raw is not None:
            return raw.decode("cp1252", errors="replace").rstrip("\x00")
        return ""

    def prop_binary(self, prop: str, prefix: str = "") -> Optional[bytes]:
        return self._stream(f"{prefix}__substg1.0_{prop}0102")  # PT_BINARY

    def sent_date(self) -> str:
        raw = self._stream("__properties_version1.0")
        if not raw or len(raw) < 32:
            return ""
        for off in range(32, len(raw) - 15, 16):   # 32-byte header, 16-byte rows
            tag = struct.unpack_from("<I", raw, off)[0]
            if tag in (0x00390040, 0x0E060040):    # ClientSubmit / MessageDelivery (PT_SYSTIME)
                filetime = struct.unpack_from("<Q", raw, off + 8)[0]
                try:
                    import datetime as _dt
                    ts = filetime / 10_000_000 - _FILETIME_EPOCH_DELTA
                    return _dt.datetime.fromtimestamp(ts, _dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
                except (OverflowError, OSError, ValueError):
                    return ""
        return ""

    def attachments(self) -> List[Tuple[str, Optional[bytes], bool]]:
        """[(filename, data_or_None, is_embedded_message)]"""
        out = []
        prefixes = sorted({p.split("/")[0] + "/" for p in self.pathmap
                           if p.startswith("__attach_version1.0_#")})
        for pre in prefixes:
            name = (self.prop_string("3707", pre) or self.prop_string("3704", pre)
                    or pre.strip("/").split("#")[-1])
            data = self.prop_binary("3701", pre)
            embedded = any(p.startswith(pre + "__substg1.0_3701000D") for p in self.pathmap)
            out.append((name, data, embedded))
        return out


def extract_msg_file(data: bytes, cfg: Config, depth: int) -> Tuple[str, List[str], Dict[str, Any], List[ScanResult]]:
    warnings: List[str] = []
    children: List[ScanResult] = []
    try:
        msg = MsgFile(data)
    except ValueError as exc:
        return "", [f"ole: {exc} — legacy .doc/.xls not supported; save as .docx/.xlsx/PDF"], {}, []

    header: List[str] = []
    sender = msg.prop_string("0C1A")
    sender_email = msg.prop_string("5D01") or msg.prop_string("0065")
    if sender or sender_email:
        header.append(f"From: {sender}{f' <{sender_email}>' if sender_email else ''}".strip())
    to = msg.prop_string("0E04")
    if to:
        header.append(f"To: {to}")
    date = msg.sent_date()
    if date:
        header.append(f"Date: {date}")
    subject = msg.prop_string("0037")
    if subject:
        header.append(f"Subject: {subject}")

    body = msg.prop_string("1000")
    if not body:
        html_body = msg.prop_binary("1013")
        if html_body:
            body = extract_html(html_body)[0]

    count = 0
    for name, payload, embedded in msg.attachments():
        count += 1
        if payload is None:
            reason = "embedded Outlook message" if embedded else "no data stream"
            warnings.append(f"msg: attachment '{name}' skipped ({reason})")
            continue
        if depth >= cfg.max_depth:
            warnings.append(f"msg: attachment '{name}' skipped (max depth {cfg.max_depth})")
        elif len(payload) > cfg.max_member_bytes:
            warnings.append(f"msg: attachment '{name}' skipped (> {cfg.max_member_bytes} bytes)")
        else:
            children.append(scan_bytes(payload, name, cfg, depth + 1))

    text = ("\n".join(header) + "\n\n" + body).strip()
    return text, warnings, {"attachments": count, "engine": "native-cfb"}, children


def extract_zip(data: bytes, cfg: Config, depth: int) -> Tuple[str, List[str], Dict[str, Any], List[ScanResult]]:
    warnings: List[str] = []
    children: List[ScanResult] = []
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        return "", [f"zip: bad archive: {exc}"], {}, []
    with zf:
        members = [m for m in zf.infolist() if not m.is_dir()]
        for member in members:
            if depth >= cfg.max_depth:
                warnings.append(f"zip: '{member.filename}' skipped (max depth {cfg.max_depth})")
                continue
            if member.file_size > cfg.max_member_bytes:
                warnings.append(f"zip: '{member.filename}' skipped (size {member.file_size})")
                continue
            try:
                payload = zf.read(member)
            except Exception as exc:
                warnings.append(f"zip: '{member.filename}' unreadable: {exc}")
                continue
            children.append(scan_bytes(payload, member.filename, cfg, depth + 1))
    return "", warnings, {"members": len(members)}, children


# --------------------------------------------------------------------------
# RFQ line-item parser
# --------------------------------------------------------------------------

_UOM_ALIASES = {"EG": "EA", "EEA": "EA", "EQ": "EA", "EFA": "EA", "FA": "EA", "£A": "EA", "€A": "EA", "CA": "EA",
                "F": "FT", "FI": "FT", "FL": "FT",
                "VOLL": "ROLL", "FF": "FT"}   # common OCR misreads of handwritten UOMs
_UOMS = {"EA", "EACH", "PC", "PCS", "FT", "FEET", "F", "C", "M", "BOX", "BX",
         "RL", "ROLL", "ROLLS", "CTN", "CS", "CASE", "SPL", "SPOOL", "LOT",
         "SET", "KIT", "PR", "PAIR", "BAG", "BAGS", "BG", "PK", "PKG",
         "COIL", "CAN", "DZ", "BDL", "TUBE", "BKT"}

# count-type units can never have fractional quantities — "1.25 EA" is an
# OCR-glued list marker ("1." + "25 EA"), not one-and-a-quarter each
_COUNT_UOMS = {"EA", "EACH", "PC", "PCS", "BOX", "BX", "SET", "KIT", "PR",
               "PAIR", "CAN", "ROLL", "RL", "ROLLS", "BAG", "BAGS", "BG",
               "DZ", "CASE", "CS", "COIL", "CTN", "PK", "PKG", "TUBE",
               "BKT", "LOT", "BDL", "SPL", "SPOOL"}

_PLURAL_UOMS = {"BOXES": "BOX", "CASES": "CASE", "COILS": "COIL",
                "ROLLS": "ROLL", "CANS": "CAN", "DOZEN": "DZ", "DOZ": "DZ",
                "PAIRS": "PR", "PIECES": "PCS", "SETS": "SET", "KITS": "KIT",
                "TUBES": "TUBE", "BUNDLES": "BDL", "SPOOLS": "SPOOL",
                "BUCKETS": "BKT", "BUCKET": "BKT"}

_INSULATION = {"THHN", "THWN", "THWN-2", "XHHW", "XHHW-2", "USE", "USE-2",
               "SER", "SEU", "MC", "NM-B", "NMB", "UF-B", "UFB", "TFFN",
               "MTW", "SOOW", "SJOOW", "PV", "TC", "DLO", "RHW", "RHW-2"}

_HEADER_TOKENS = {"QTY", "QUANTITY", "DESCRIPTION", "PART", "ITEM", "UNIT",
                  "PRICE", "U/M", "UOM", "MFR", "MANUFACTURER", "CAT", "LINE",
                  "EXT", "EXTENDED", "TOTAL", "NO.", "NUMBER", "CATALOG"}

_GAUGE_RE = re.compile(r"\b(\d{1,2}/0|\d{1,4})\s*(AWG|MCM|KCMIL)\b", re.IGNORECASE)
_HASH_GAUGE_RE = re.compile(r"#\s*(\d{1,2}(?:/0)?)\b")
_QTY_LINE_RE = re.compile(
    r"^\s*(?:(?P<lineno>\d{1,3})\s*[\.\):,;!]\s+)?"    # optional line-number marker "1)" / "2." / "3:" / "4," / "5;"
    r"(?P<qty>\d{1,6}(?:[\.,]\d{1,3})?)\s+"            # quantity
    r"(?:(?P<uom>[A-Za-z/\-]{1,6})\s+)?"               # optional UOM
    r"(?P<rest>\S.*)$"
)
_PART_TOKEN_RE = re.compile(r"\b(?=[A-Z0-9\-\./#]*\d)[A-Z][A-Z0-9\-\./#]{2,24}\b|\b\d{4,}[A-Z][A-Z0-9\-]*\b")


# --------------------------------------------------------------------------
# Teach store — persistent human corrections that override the heuristics
# --------------------------------------------------------------------------

def _norm_teach(line: str) -> str:
    return re.sub(r"\s+", " ", line.strip().upper())


class TeachStore:
    """JSON-backed store of learned line rules keyed by the normalized raw
    source line. Two rule kinds:
      {"action": "set", "qty":..., "uom":..., "part":..., "gauge":...,
       "description":...}   -> emit exactly these fields for this line
      {"action": "suppress"} -> never emit an item for this line
    """

    def __init__(self, path: str):
        self.path = path
        self.rules: Dict[str, Dict[str, Any]] = {}
        self.last_error: Optional[str] = None
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    loaded = json.load(fh)
                if isinstance(loaded, dict):
                    self.rules = {k: v for k, v in loaded.items() if isinstance(v, dict)}
            except (OSError, json.JSONDecodeError):
                pass  # corrupt/unreadable teach file must never break scanning

    def get(self, line: str, fuzzy: float = 0.0) -> Optional[Dict[str, Any]]:
        """Exact normalized-line lookup, with optional fuzzy fallback: OCR
        renders the same physical line slightly differently across scans, so
        a high-similarity match still applies the learned rule. Suppress
        rules require a stricter ratio (deleting the wrong line is worse
        than re-showing a corrected one)."""
        norm = _norm_teach(line)
        rule = self.rules.get(norm)
        if rule is not None or not fuzzy or len(norm) < 12:
            return rule
        import difflib
        best_ratio, best_rule = 0.0, None
        for key, candidate in self.rules.items():
            if abs(len(key) - len(norm)) > max(6, 0.3 * len(norm)):
                continue
            ratio = difflib.SequenceMatcher(None, norm, key).ratio()
            if ratio > best_ratio:
                best_ratio, best_rule = ratio, candidate
        if best_rule is None:
            return None
        need = max(fuzzy, 0.93) if best_rule.get("action") == "suppress" else fuzzy
        return best_rule if best_ratio >= need else None

    def set_rule(self, line: str, fields: Dict[str, Any]):
        rule = {"action": "set"}
        for key in ("qty", "uom", "part", "gauge", "description"):
            if fields.get(key) not in (None, ""):
                rule[key] = fields[key]
        self.rules[_norm_teach(line)] = rule
        self.save()

    def suppress(self, line: str):
        self.rules[_norm_teach(line)] = {"action": "suppress"}
        self.save()

    def forget(self, line: str):
        if self.rules.pop(_norm_teach(line), None) is not None:
            self.save()

    def save(self):
        """Persist rules. On failure, last_error is set so UIs can warn the
        user — corrections silently failing to persist is data loss."""
        try:
            with open(self.path, "w", encoding="utf-8") as fh:
                json.dump(self.rules, fh, indent=2, sort_keys=True)
            self.last_error = None
        except OSError as exc:
            self.last_error = f"{type(exc).__name__}: {exc}"


_TEACH_CACHE: Dict[Tuple[str, Optional[float]], TeachStore] = {}


def load_teach(path: Optional[str]) -> Optional[TeachStore]:
    """mtime-keyed cache so repeated scan_bytes calls don't re-read the file,
    but external writes (e.g. a GUI teaching mid-session) are picked up."""
    if not path:
        return None
    try:
        mtime: Optional[float] = os.path.getmtime(path)
    except OSError:
        mtime = None
    key = (path, mtime)
    if key not in _TEACH_CACHE:
        _TEACH_CACHE.clear()
        _TEACH_CACHE[key] = TeachStore(path)
    return _TEACH_CACHE[key]


_VOID_LINE_RE = re.compile(
    r"\b(VOID|DO NOT QUOTE|IGNORE THIS|IGNORE|OLD COUNT|SCRATCH\s*:|"
    r"CHANGED\s*:|CANCELL?ED|ALREADY ORDERED|OMIT THIS|DELETED|DISREGARD)\b",
    re.IGNORECASE)


class LineItemParser:
    def __init__(self, cfg: Config, teach: Optional[TeachStore] = None):
        self.cfg = cfg
        self.teach = teach if teach is not None else load_teach(cfg.teach_path)
        self.last_voided: List[str] = []    # lines suppressed as crossed-out

    # -- OCR normalization helpers ---------------------------------------
    _OCR_PIPE_RE = re.compile(r"\|+")
    _LEAD_JUNK_RE = re.compile(r"^[#(:;\.\)\]\[]+\s*")
    _LEADING_PAIR_RE = re.compile(r"^\s*(\d{1,3})\s+[#(:;\.\[]*\s*\d")
    _OCR_ONE_RE = re.compile(r"^\s*[iIl!]\s+\d")

    _GLUED_UOM_RE = re.compile(
        r"\b(\d+)\.?(ea|eea|eq|eg|\u20aca|ca|ft|pc|pcs|rl|roll|bx|box|set|spl|"
        r"pr|pk|bag|bg|can|case|kit|coil|dz|dozen|f)(?![A-Za-z0-9/])",
        re.IGNORECASE)

    _EURO_A_RE = re.compile(r"\u20ac\s*a\b", re.IGNORECASE)

    @classmethod
    def _preclean(cls, line: str) -> str:
        """Normalize OCR table artifacts: column-border pipes become spaces,
        the euro-sign misread of handwritten 'ea' is restored, and glued
        qty+UOM tokens ('25ea') get re-split."""
        line = cls._OCR_PIPE_RE.sub(" ", line)
        line = re.sub(r"(\d)\s*\u00a2(?=\s*ea\b)", r"\1 ", line)
        line = re.sub(r"(?<=\d)U(?=['\u2019\u2032])", "0", line)
        line = re.sub(r"^\s*[\"\u201c\u201d'\u2019`*~]+\s*(?=\d)", "", line)
        line = re.sub(r"(?<=[A-Za-z])['\u2019](?=\d)", " ", line)
        line = cls._EURO_A_RE.sub("ea", line)
        line = re.sub(r"(\d)['\u2019\u2032](?=[\s).,]|$)", r"\1 ft", line)
        return cls._GLUED_UOM_RE.sub(r"\1 \2", line)

    @staticmethod
    def _to_qty(s: str) -> Optional[float]:
        """'2,000' is two thousand (thousands separator), '2,5' is 2.5."""
        s = s.strip()
        if re.fullmatch(r"\d{1,3}(?:,\d{3})+", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    def _detect_lineno_rows(self, lines: List[str]) -> Dict[int, str]:
        """Detect a table line-number COLUMN: runs of >=3 lines whose leading
        small integers increment by exactly 1 and are followed by another
        number (the real qty). Returns {line_idx: mode} where mode 'int'
        strips a leading integer and 'tok' strips an OCR-garbled first token
        (e.g. tesseract reading a lone '1' as 'i')."""
        cands: List[Tuple[int, int]] = []
        for idx, line in enumerate(lines):
            m = self._LEADING_PAIR_RE.match(line)
            if m:
                cands.append((idx, int(m.group(1))))
        marked: Dict[int, str] = {}
        run: List[Tuple[int, int]] = []

        def flush():
            if len(run) >= 3:
                for r_idx, _val in run:
                    marked[r_idx] = "int"
                first_idx, first_val = run[0]
                if first_val == 2:          # OCR misread of the '1' row above?
                    j = first_idx - 1
                    while j >= 0 and not lines[j].strip():
                        j -= 1
                    if j >= 0 and self._OCR_ONE_RE.match(lines[j]):
                        marked[j] = "tok"
            run.clear()

        # stride 1 = clean line numbers (run >= 3). Larger constant strides
        # catch OCR-garbled markers: handwritten "1." "2." "3." often reads as
        # 16, 26, 36 (dot -> digit), a perfect stride-10 sequence. Require a
        # longer run (>= 4) for non-unit strides to stay conservative.
        def try_stride(stride: int, min_run: int):
            local: List[Tuple[int, int]] = []

            def local_flush():
                if len(local) >= min_run:
                    for r_idx, _v in local:
                        marked.setdefault(r_idx, "int")
                    first_idx, first_val = local[0]
                    if first_val == 1 + stride:   # OCR misread of row 1?
                        j = first_idx - 1
                        while j >= 0 and not lines[j].strip():
                            j -= 1
                        if j >= 0 and self._OCR_ONE_RE.match(lines[j]):
                            marked.setdefault(j, "tok")
                local.clear()

            for c_idx, c_val in cands:
                if local and c_val == local[-1][1] + stride:
                    local.append((c_idx, c_val))
                else:
                    local_flush()
                    local.append((c_idx, c_val))
            local_flush()

        for idx, val in cands:
            if run and val == run[-1][1] + 1:
                run.append((idx, val))
            else:
                flush()
                run.append((idx, val))
        flush()
        for stride in (10, 11, 9):
            try_stride(stride, 4)
        return marked

    _MARKER_TOKEN_RE = re.compile(
        r"^\s*((?:\d{1,3}\s*[a-z]?[\.,;!:]?)|(?:\d{1,3}[^\w\s]{1,2}[\.,;!:]?)|"
        r"(?:[a-z]{1,2}[\.,;!:]?)|(?:[^\w\s]{1,3}[\.,;!:]?))\s+",
        re.IGNORECASE)
    _CLEAN_MARKER_RE = re.compile(r"^\s*\d{1,3}\s*[\.,;]\s")

    def parse(self, text: str) -> List[LineItem]:
        items: List[LineItem] = []
        self.last_voided = []
        lines = text.splitlines()
        cleaned = [self._preclean(l) for l in lines]
        marked = self._detect_lineno_rows(cleaned)
        # marker-column doc: a numbered list ("1." "2." ...). OCR garbles many
        # of these markers on handwriting ("1."->"16", "6."->"b.", "9."->"%"),
        # so with enough clean markers as evidence, garbled leading tokens are
        # stripped whenever the remainder parses as qty+UOM.
        marker_doc = sum(1 for l in cleaned
                         if self._CLEAN_MARKER_RE.match(l)) >= 4
        paren_doc = sum(1 for l in cleaned
                        if self._PAREN_ITEM_RE.match(l)) >= 3
        prev_lineno: Optional[int] = None
        last_consumed = -99
        for idx, raw_line in enumerate(lines):
            display = raw_line.strip()
            if not display or len(display) > 300:
                continue
            rule = (self.teach.get(display, fuzzy=self.cfg.teach_fuzzy)
                    if self.teach else None)
            if rule is not None:
                if rule.get("action") == "suppress":
                    continue
                if rule.get("action") == "set":
                    qty = rule.get("qty")
                    items.append(LineItem(
                        raw=display, qty=float(qty) if qty is not None else None,
                        uom=rule.get("uom"), part=rule.get("part"),
                        gauge=rule.get("gauge"),
                        description=rule.get("description") or display,
                        confidence=1.0, line_no=idx, taught=True))
                    continue
            work = cleaned[idx].strip()
            if self._is_header(work):
                continue
            if _VOID_LINE_RE.search(work):
                self.last_voided.append(display)
                continue
            mode = marked.get(idx)
            if paren_doc:
                item = self._paren_line(work)
                if item is None:
                    # wrapped continuation of the previous item?
                    if (items and idx - last_consumed == 1
                            and not self._PAGENUM_RE.match(work)):
                        items[-1].description = (
                            (items[-1].description or "") + " "
                            + _normalize_desc(work)).strip()
                        last_consumed = idx
                        continue
                    item = self._parse_line(work, None)
            elif mode is not None:
                if mode == "int":
                    work2 = re.sub(r"^\s*\d{1,3}\s+", "", work)
                else:                        # 'tok': OCR-garbled line number
                    work2 = re.sub(r"^\s*\S+\s+", "", work)
                work2 = self._LEAD_JUNK_RE.sub("", work2)
                item = self._parse_line(work2, None)
            elif marker_doc:
                item = self._marker_doc_line(work, prev_lineno)
                if item is not None and item.meta_lineno is not None:
                    prev_lineno = item.meta_lineno
            else:
                item = self._parse_line(work, prev_lineno)
                if item is not None and item.meta_lineno is not None:
                    prev_lineno = item.meta_lineno
            if item is None:
                continue
            if item.confidence >= self.cfg.min_item_confidence:
                item.raw = display          # what the user sees in the doc
                item.line_no = idx
                items.append(item)
                last_consumed = idx
        if self.cfg.default_uom:
            for item in items:
                if item.qty is not None and not item.uom:
                    item.uom = "FT" if item.gauge else self.cfg.default_uom
        return items

    # ---- parenthesized-quantity format: "-(3 boxes) desc", "-(400') desc" ----
    _PAREN_ITEM_RE = re.compile(
        r"^\s*[\-\u2013\u2014~_\.\u2022\*\u00b7]*\s*\(\s*"
        r"(?P<qty>\d{1,6}(?:[\.,]\d{1,3})?)\s*(?P<inner>[^)]*)\)\s*(?P<rest>.*)$")
    _PAREN_GARBLED_RE = re.compile(       # OCR ate the '(' or fused the dash
        r"^\s*\S{0,3}?(?P<qty>\d{1,6})\s*"
        r"(?P<inner>[A-Za-z\./'\u2019\u2032]{0,12})\)\s+(?P<rest>\S.*)$")
    _PAREN_NOQTY_RE = re.compile(         # qty unreadable -> review row
        r"^\s*\S{0,2}\(\s*(?P<inner>[^)]{0,20})\)\s*(?P<rest>\S.*)$")
    _PAGENUM_RE = re.compile(r"^\s*\d{1,3}\s*$")

    @staticmethod
    def _uom_from_inner(inner: str) -> Tuple[Optional[str], Optional[str]]:
        """'(3 boxes)' inner='boxes' -> BOX; "(400\')" inner="\'" -> FT;
        '(2 cases/100pcs.)' -> (CASE, '100pcs') pack note."""
        inner = inner.strip()
        if not inner:
            return None, None
        if inner[0] in "'\u2019\u2032":
            return "FT", None
        packnote = None
        if "/" in inner:
            head, _, tail = inner.partition("/")
            packnote = tail.strip(" .") or None
            inner = head
        tok = re.sub(r"[^A-Za-z]", "", inner).upper()
        tok = _PLURAL_UOMS.get(tok, tok)
        tok = _UOM_ALIASES.get(tok, tok)
        return (tok if tok in _UOMS else None), packnote

    def _paren_line(self, work: str) -> Optional["_ParsedItem"]:
        m = self._PAREN_ITEM_RE.match(work) or self._PAREN_GARBLED_RE.match(work)
        if m:
            qty = self._to_qty(m.group("qty"))
            uom, packnote = self._uom_from_inner(m.group("inner"))
            rest = self._LEAD_JUNK_RE.sub("", m.group("rest").strip())
            rest = re.sub(r"^[\-\u2013\u2014_~=.\u00b7]+\s*", "", rest)
            desc = _normalize_desc(rest) if rest else ""
            if packnote:
                desc = (desc + f"  [{packnote} per {(uom or 'pack').title()}]").strip()
            conf = 0.2 + (0.25 if qty is not None else 0.0) + 0.10
            if uom:
                conf += 0.15
            part = self._find_part(rest) if rest else None
            if part:
                conf += 0.25
            item = _ParsedItem(raw=work, qty=qty, uom=uom, part=part,
                               description=desc or work,
                               confidence=min(conf, 1.0))
            item.apply_gauge(work)
            return item
        m = self._PAREN_NOQTY_RE.match(work)
        if m:
            uom, _packnote = self._uom_from_inner(m.group("inner"))
            rest = m.group("rest").strip()
            item = _ParsedItem(raw=work, uom=uom,
                               description=_normalize_desc(rest),
                               confidence=0.45)
            item.apply_gauge(work)
            return item
        return None

    def _marker_doc_line(self, work: str, prev_lineno: Optional[int]):
        """Parse one line of a numbered-list document, tolerating OCR-garbled
        markers. Preference order:
          1. remainder-after-marker when it parses with qty + valid UOM
             (fixes '16 25 ea', '2e 40 ea', 'b. 10 ea', '% 2 ea')
          2. the normal parse (clean '3.' markers land here)
          3. marker present but qty unreadable -> emit a qty-less review row
             rather than silently dropping the line ('4, ¢ ea ...')"""
        normal = self._parse_line(work, prev_lineno)
        mk = self._MARKER_TOKEN_RE.match(work)
        if mk:
            rest_raw = work[mk.end():].strip()
            if self._PAREN_ITEM_RE.match(rest_raw):   # "-(3 box) ..." / "(2 roll) ..."
                paren = self._paren_line(rest_raw)
                if paren is not None and paren.qty is not None:
                    return paren
            rest = self._LEAD_JUNK_RE.sub("", rest_raw)
            cand = self._parse_line(rest, prev_lineno) if rest else None
            if cand is not None and cand.qty is not None and not cand.uom:
                # "48 1000 FT ..." — leading token is the line number when
                # the remainder parses with qty + real UOM
                rest2 = re.sub(r"^\s*\d{1,3}\s+", "", rest)
                cand2 = (self._parse_line(rest2, prev_lineno)
                         if rest2 != rest else None)
                if cand2 is not None and cand2.qty is not None and cand2.uom:
                    cand = cand2
            if (cand is not None and cand.qty is not None and cand.uom
                    and not (normal is not None
                             and normal.meta_lineno is not None)):
                cand.confidence = min(cand.confidence + 0.1, 1.0)
                return cand
        if normal is not None and normal.confidence >= self.cfg.min_item_confidence:
            return normal
        # shredded-line salvage: an orphan fragment carrying a part number is
        # worth a review row — silently dropping a part is worse than asking
        part = self._find_part(work)
        if part:
            item = _ParsedItem(raw=work, part=part,
                               description=_normalize_desc(work),
                               confidence=0.45)
            item.apply_gauge(work)
            return item
        if mk:
            rest = work[mk.end():].strip()
            rest = re.sub(r"^[^\w]+\s*", "", rest)
            if len(rest) >= 4 and re.search(r"\d", rest):
                item = _ParsedItem(raw=work, part=self._find_part(rest),
                                   description=_normalize_desc(rest),
                                   confidence=0.45)
                item.apply_gauge(work)
                return item
        return normal

    def parse_one(self, line: str) -> Optional[LineItem]:
        """Parse a single line with no sequence context. Used by review UIs
        for manual add/correct flows; ignores the confidence floor."""
        display = line.strip()
        item = self._parse_line(self._preclean(display).strip(), None)
        if item is not None:
            item.raw = display
            if self.cfg.default_uom and item.qty is not None and not item.uom:
                item.uom = "FT" if item.gauge else self.cfg.default_uom
        return item

    @staticmethod
    def _is_header(line: str) -> bool:
        tokens = {t.strip(":#").upper() for t in re.split(r"[\s\t]+", line) if t.strip()}
        return len(tokens & _HEADER_TOKENS) >= 2

    def _parse_line(self, line: str, prev_lineno: Optional[int]) -> Optional["_ParsedItem"]:
        # A line STARTING with wire-size+insulation ("12 THHN stranded...")
        # is a description fragment — the 12 is a gauge, never a quantity.
        if _SIZE_INSUL_RE.match(line.upper()):
            part = self._find_part(line)
            if part:
                item = _ParsedItem(raw=line, part=part,
                                   description=_normalize_desc(line),
                                   confidence=0.45)
                item.apply_gauge(line)
                return item
            return None
        # glued multiplier ("2x widget") and labeled qty ("QTY: 5 widget")
        line = re.sub(r"^\s*(\d{1,4})[xX]\s+", r"\1 ", line)
        # handwritten "ea" losing its e: "6a 4\" round boxes" -> "6 ea ...".
        # Lowercase-only and qty-position-only, so amp ratings (20A, 30A —
        # uppercase, mid-description) are never touched.
        line = re.sub(r"^\s*(\d{1,4})a\s+", r"\1 ea ", line)
        line = re.sub(r"^(\s*(?:\d{1,3}\s*[\.\):,]\s+)?)[tl]ea\b\s+",
                      lambda mm: (mm.group(1) or "") + "1 ea ", line)
        line = re.sub(r"^\s*(?:\d{1,3}\s*[\.\):,]?\s+)?QTY[\s:.,#]*(\d{1,6})"
                      r"\s*/\s*([A-Za-z]{1,6})\b[\s:\-\u2013\u2014]*",
                      r"\1 \2 ", line, flags=re.IGNORECASE)
        line = re.sub(r"^\s*QTY[\s:.,#]*(\d{1,6})\b[\s:,\-]*", r"\1 ", line,
                      flags=re.IGNORECASE)
        m = _QTY_LINE_RE.match(line)
        if not m:
            # No leading qty — still consider part-number-only lines
            part = self._find_part(line)
            if part:
                item = _ParsedItem(raw=line, part=part,
                                   description=_normalize_desc(line), confidence=0.2)
                item.apply_gauge(line)
                return item
            return None

        qty_str = m.group("qty")
        rest = m.group("rest")
        uom = (m.group("uom") or "").upper()
        lineno = int(m.group("lineno")) if m.group("lineno") else None

        # Line-number-as-qty repair: "1 100 EA WIDGET" where explicit marker
        # was absent but the qty field is a small sequential int followed by
        # another number — reparse rest as the real qty line.
        if lineno is None and prev_lineno is not None:
            qv = self._to_qty(qty_str)
            as_int = int(qv) if qv is not None and float(qv).is_integer() else None
            if as_int is not None and as_int == prev_lineno + 1 and as_int <= 500:
                inner = _QTY_LINE_RE.match(rest if not uom else f"{m.group('uom')} {rest}")
                if inner and inner.group("qty"):
                    lineno = as_int
                    qty_str = inner.group("qty")
                    uom = (inner.group("uom") or "").upper()
                    rest = inner.group("rest")

        uom = uom.strip("-\u2013\u2014")
        uom = _UOM_ALIASES.get(uom, uom)
        glued = re.fullmatch(r"([1-9]\d{0,2})[\.,](\d{1,6})", qty_str)
        if glued and not re.fullmatch(r"\d{1,3}(?:,\d{3})+", qty_str):
            # "1.25 EA" = marker 1. + qty 25 (count units can't be
            # fractional); for measured units (FT) only split when the int
            # part is exactly the expected next line number
            expected = (lineno is None and prev_lineno is not None
                        and int(glued.group(1)) == prev_lineno + 1)
            if uom in _COUNT_UOMS or expected:
                lineno = lineno if lineno is not None else int(glued.group(1))
                qty_str = glued.group(2)
        if uom and uom not in _UOMS:
            # Not a real UOM — it's the first word of the description
            rest = f"{m.group('uom')} {rest}" if m.group("uom") else rest
            uom = ""

        qty = self._to_qty(qty_str)
        if qty is None:
            return None

        rest = re.sub(r"^[\-\u2013\u2014_~=.\u00b7)\]]+\s*", "", rest.strip())
        item = _ParsedItem(raw=line, qty=qty, uom=uom or None,
                           description=_normalize_desc(rest), confidence=0.2 + 0.25)
        item.meta_lineno = lineno
        if uom:
            item.confidence += 0.15
        part = self._find_part(rest)
        if part:
            item.part = part
            item.confidence += 0.25
        item.apply_gauge(line)
        item.confidence = min(item.confidence, 1.0)
        return item

    @staticmethod
    def _find_part(text: str) -> Optional[str]:
        explicit = re.search(r"PART\s*#\s*:?\s*([A-Z0-9][A-Z0-9\-\./]{2,24})",
                             text.upper())
        if explicit:
            return explicit.group(1).rstrip(".-/")
        candidates = _PART_TOKEN_RE.findall(text.upper())
        candidates = [c for c in candidates
                      if not re.fullmatch(r"[A-Z]?/\d{2,4}", c)
                      and not re.fullmatch(r"\d{2,4}[-/]\d{2,4}Y?(/\d{2,4})?", c)]
        candidates = [c for c in candidates
                      if c not in _UOMS and c not in _INSULATION
                      and not _GAUGE_RE.fullmatch(c)
                      and not re.fullmatch(r"\d{1,3}", c)]
        if not candidates:
            return None
        # Prefer tokens with mixed letters+digits and separators (most part-like)
        def score(tok: str) -> Tuple[int, int]:
            has_sep = 1 if re.search(r"[\-\./]", tok) else 0
            mixed = 1 if (re.search(r"[A-Z]", tok) and re.search(r"\d", tok)) else 0
            return (has_sep + mixed, len(tok))
        return max(candidates, key=score)


_INCH_MARK_RE = re.compile(r"(\d)\s*[\u00b0*]")
_DESC_WORD_FIXES = ((re.compile(r"\bvigid\b", re.IGNORECASE), "rigid"),)


_LONE_I_RE = re.compile(r"^[ilI|]\s+(?=\d)")


def _normalize_desc(s: str) -> str:
    """Clean common OCR misreads in descriptions: a degree sign or asterisk
    right after a digit is a handwritten inch mark (4\u00b0 -> 4\"), and a
    small list of unambiguous word fixes (vigid -> rigid)."""
    s = _LONE_I_RE.sub("", s)
    s = re.sub(r"(?:\s+(?:[^\w\s]+|[a-z]))+$", "", s)
    s = _INCH_MARK_RE.sub(r'\1"', s)
    for pattern, repl in _DESC_WORD_FIXES:
        s = pattern.sub(repl, s)
    return s


_WIRE_CONTEXT_RE = re.compile(
    r"\b(THHN|THWN|XHHW|WIRE|CABLE|STRANDED|SOLID|COPPER|CU|ALUM|ALUMINUM|"
    r"MC|GROUND|GRD|BARE|ROMEX|SER|SEU|UF)\b", re.IGNORECASE)
_SIZE_INSUL_RE = re.compile(
    r"\b(\d{1,2}(?:/0)?)\s+(THHN|THWN(?:-2)?|XHHW(?:-2)?|USE(?:-2)?|TFFN|MTW)\b",
    re.IGNORECASE)


class _ParsedItem(LineItem):
    """LineItem plus transient parse state (not serialized)."""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.meta_lineno: Optional[int] = None

    def apply_gauge(self, line: str):
        upper = line.upper()
        m = _GAUGE_RE.search(upper)
        if m:
            self.gauge = f"{m.group(1)} {m.group(2).upper()}"
            self.confidence += 0.05
        else:
            hm = _HASH_GAUGE_RE.search(upper)
            sm = _SIZE_INSUL_RE.search(upper)
            if hm and _WIRE_CONTEXT_RE.search(upper):
                # '#12' means 12 AWG only in wire context; on '#1 EA 4in box'
                # it's OCR junk / a count, not a gauge.
                self.gauge = f"{hm.group(1)} AWG"
                self.confidence += 0.05
            elif sm:
                # bare size+insulation shorthand: '12 THHN' = 12 AWG THHN
                self.gauge = f"{sm.group(1)} AWG"
                self.confidence += 0.05
        if any(tok in upper.split() or tok in upper for tok in _INSULATION):
            self.confidence += 0.05

    def to_dict(self) -> Dict[str, Any]:
        d = super().to_dict()
        d.pop("meta_lineno", None)
        return d


# --------------------------------------------------------------------------
# Catalog grounding (SQLite, schema auto-detected)
# --------------------------------------------------------------------------

_PART_COL_HINTS = ("part", "catalog", "sku", "item", "number", "pn", "cat_no", "catno")
_DESC_COL_HINTS = ("desc", "name", "title")
_MFR_COL_HINTS = ("mfr", "manufacturer", "brand", "line", "vendor")


def _normalize_part(part: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", part.upper())


class Catalog:
    """Loads a normalized part-number index from an arbitrary SQLite catalog."""

    def __init__(self, path: str):
        self.path = path
        self.index: Dict[str, Dict[str, Any]] = {}
        self.warnings: List[str] = []
        self._load()

    def _load(self):
        if not os.path.exists(self.path):
            self.warnings.append(f"catalog: file not found: {self.path}")
            return
        try:
            conn = sqlite3.connect(f"file:{self.path}?mode=ro", uri=True)
        except sqlite3.Error as exc:
            self.warnings.append(f"catalog: open failed: {exc}")
            return
        try:
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [r[0] for r in cur.fetchall()]
            loaded_rows = 0
            for table in tables:
                cur.execute(f'PRAGMA table_info("{table}")')
                cols = [r[1] for r in cur.fetchall()]
                part_col = self._pick(cols, _PART_COL_HINTS)
                if not part_col:
                    continue
                desc_col = self._pick(cols, _DESC_COL_HINTS)
                mfr_col = self._pick(cols, _MFR_COL_HINTS)
                select_cols = [part_col] + [c for c in (desc_col, mfr_col) if c]
                col_sql = ", ".join(f'"{c}"' for c in select_cols)
                try:
                    cur.execute(f'SELECT {col_sql} FROM "{table}"')
                except sqlite3.Error as exc:
                    self.warnings.append(f"catalog: select from {table} failed: {exc}")
                    continue
                for row in cur.fetchall():
                    part = str(row[0] or "").strip()
                    if not part:
                        continue
                    norm = _normalize_part(part)
                    if not norm or norm in self.index:
                        continue
                    entry = {"part": part, "table": table}
                    idx = 1
                    if desc_col:
                        entry["description"] = str(row[idx] or "")
                        idx += 1
                    if mfr_col:
                        entry["manufacturer"] = str(row[idx] or "")
                    self.index[norm] = entry
                    loaded_rows += 1
            if not loaded_rows:
                self.warnings.append("catalog: no part-number-like columns found in any table")
        finally:
            conn.close()

    @staticmethod
    def _pick(cols: List[str], hints: Tuple[str, ...]) -> Optional[str]:
        for hint in hints:
            for col in cols:
                if hint in col.lower():
                    return col
        return None

    def lookup(self, part: str) -> Optional[Dict[str, Any]]:
        norm = _normalize_part(part)
        if not norm:
            return None
        hit = self.index.get(norm)
        if hit:
            return dict(hit, match="exact")
        swapped = norm.replace("O", "0").replace("I", "1")
        if swapped != norm:                 # OCR O/0 and I/1 confusion
            hit = self.index.get(swapped)
            if hit:
                return dict(hit, match="ocr-o0")
        if len(norm) >= 5:
            for key, entry in self.index.items():
                if key.startswith(norm) or norm.startswith(key):
                    return dict(entry, match="prefix")
        return None

    def ground(self, items: List[LineItem]):
        for item in items:
            if item.part:
                hit = self.lookup(item.part)
                if hit:
                    item.catalog_match = hit
                    item.confidence = min(item.confidence + 0.25, 1.0)


# --------------------------------------------------------------------------
# Core scan dispatch
# --------------------------------------------------------------------------

def scan_bytes(data: bytes, name: str, cfg: Optional[Config] = None, depth: int = 0) -> ScanResult:
    cfg = cfg or Config()
    fmt = detect_format(data, name)
    result = ScanResult(name=name, fmt=fmt)
    children: List[ScanResult] = []

    try:
        if fmt == "pdf":
            text, warns, meta = extract_pdf(data, cfg)
        elif fmt == "docx":
            text, warns, meta = extract_docx(data)
        elif fmt == "xlsx":
            text, warns, meta = extract_xlsx(data)
        elif fmt == "csv":
            text, warns, meta = extract_csv(data, name)
        elif fmt == "html":
            text, warns, meta = extract_html(data)
        elif fmt == "image":
            text, warns, meta = extract_image(data, cfg)
        elif fmt == "eml":
            text, warns, meta, children = extract_eml(data, cfg, depth)
        elif fmt == "ole":
            text, warns, meta, children = extract_msg_file(data, cfg, depth)
        elif fmt in ("zip", "pptx-zip"):
            text, warns, meta, children = extract_zip(data, cfg, depth)
        elif fmt == "text":
            text, warns, meta = extract_text_plain(data)
        else:
            text, warns, meta = "", [f"unsupported/binary format ({fmt})"], {}
            result.ok = False
    except Exception as exc:  # never let one file kill a batch
        result.ok = False
        result.warnings.append(f"extractor crashed: {type(exc).__name__}: {exc}")
        return result

    result.text = text[:cfg.max_text_chars]
    if len(text) > cfg.max_text_chars:
        result.warnings.append(f"text truncated to {cfg.max_text_chars} chars")
    result.warnings.extend(warns)
    result.meta = meta
    result.children = children

    if result.text.strip():
        parser = LineItemParser(cfg)
        result.line_items = parser.parse(result.text)
        if parser.last_voided:
            result.meta["voided_lines"] = parser.last_voided
            result.warnings.append(
                f"suppressed {len(parser.last_voided)} crossed-out/void line(s) "
                "— see meta.voided_lines")
    return result


def scan_path(path: str, cfg: Optional[Config] = None) -> ScanResult:
    cfg = cfg or Config()
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError as exc:
        res = ScanResult(name=path, fmt="unknown", ok=False)
        res.warnings.append(f"read failed: {exc}")
        return res
    result = scan_bytes(data, os.path.basename(path), cfg)
    result.meta["path"] = path
    result.meta["size"] = len(data)
    return result


# --------------------------------------------------------------------------
# Self-test — synthesizes real files for every stdlib-capable format
# --------------------------------------------------------------------------

def _build_minimal_pdf(text_line: str) -> bytes:
    stream = f"BT /F1 12 Tf 72 720 Td ({text_line}) Tj ET".encode("ascii")
    bodies = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
         b"/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>"),
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = b"%PDF-1.4\n"
    offsets = []
    for i, body in enumerate(bodies, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += f"xref\n0 {len(bodies)+1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(bodies)+1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF".encode()
    return out


def _build_two_page_pdf(text_line: str) -> bytes:
    """Page 1 has a text layer; page 2 is empty (simulates a scanned page)."""
    stream = f"BT /F1 12 Tf 72 720 Td ({text_line}) Tj ET".encode("ascii")
    bodies = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R 6 0 R] /Count 2 >>",
        (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
         b"/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>"),
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
         b"/Contents 7 0 R /Resources << >> >>"),
        b"<< /Length 0 >>\nstream\n\nendstream",
    ]
    out = b"%PDF-1.4\n"
    offsets = []
    for i, body in enumerate(bodies, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += f"xref\n0 {len(bodies)+1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(bodies)+1} /Root 1 0 R >>\nstartxref\n"
            f"{xref_pos}\n%%EOF").encode()
    return out


def _build_docx(paragraphs: List[str]) -> bytes:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = "".join(f'<w:p><w:r><w:t xml:space="preserve">{p}</w:t></w:r></w:p>' for p in paragraphs)
    doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="{w}"><w:body>{body}</w:body></w:document>'
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("word/document.xml", doc)
    return buf.getvalue()


def _build_xlsx(rows: List[List[str]]) -> bytes:
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    strings: List[str] = []
    def sref(s: str) -> int:
        strings.append(s)
        return len(strings) - 1
    row_xml = []
    for r, row in enumerate(rows, start=1):
        cells = []
        for c, val in enumerate(row):
            col = chr(ord("A") + c)
            cells.append(f'<c r="{col}{r}" t="s"><v>{sref(val)}</v></c>')
        row_xml.append(f'<row r="{r}">{"".join(cells)}</row>')
    sheet = f'<?xml version="1.0"?><worksheet xmlns="{ns}"><sheetData>{"".join(row_xml)}</sheetData></worksheet>'
    sst = (f'<?xml version="1.0"?><sst xmlns="{ns}" count="{len(strings)}" uniqueCount="{len(strings)}">'
           + "".join(f'<si><t xml:space="preserve">{s}</t></si>' for s in strings) + "</sst>")
    workbook = f'<?xml version="1.0"?><workbook xmlns="{ns}"><sheets><sheet name="Sheet1" sheetId="1"/></sheets></workbook>'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/worksheets/sheet1.xml", sheet)
        zf.writestr("xl/sharedStrings.xml", sst)
    return buf.getvalue()


def _build_eml(body: str, attach_name: str, attach_data: bytes) -> bytes:
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["From"] = "buyer@example.com"
    msg["To"] = "quotes@americanpower.example"
    msg["Subject"] = "RFQ 4512 — wire and fittings"
    msg.set_content(body)
    msg.add_attachment(attach_data, maintype="text", subtype="csv", filename=attach_name)
    return msg.as_bytes()


def _build_msg_file(subject: str, body: str, attach_name: str, attach_data: bytes) -> bytes:
    """Minimal from-scratch CFB v3 writer producing a valid Outlook .msg
    (selftest only). All streams are small, so everything lives in the
    ministream with a proper mini FAT."""
    def utf16(s: str) -> bytes:
        return s.encode("utf-16-le")

    # Top-level properties stream: 32-byte header (reserved8, nextRecipId,
    # nextAttachId, recipCount, attachCount, reserved8) + 16-byte rows.
    filetime = (1_768_435_200 + _FILETIME_EPOCH_DELTA) * 10_000_000  # 2026-01-15
    top_props = (b"\x00" * 8 + struct.pack("<4I", 0, 1, 0, 1) + b"\x00" * 8
                 + struct.pack("<IIQ", 0x00390040, 6, filetime))     # ClientSubmitTime
    # Attachment properties stream: 8-byte header + rows.
    attach_props = (b"\x00" * 8
                    + struct.pack("<IIII", 0x37050003, 6, 1, 0)      # AttachMethod=byValue
                    + struct.pack("<IIII", 0x0E210003, 6, 0, 0))     # AttachNumber=0

    # (name, etype, payload, parent_id) — tree wired automatically below
    nodes = [
        ("Root Entry", 5, b"", None),
        ("__nameid_version1.0", 1, b"", 0),
        ("__substg1.0_00020102", 2, b"", 1),      # GUID stream (empty = no named props)
        ("__substg1.0_00030102", 2, b"", 1),      # entry stream
        ("__substg1.0_00040102", 2, b"", 1),      # string stream
        ("__properties_version1.0", 2, top_props, 0),
        ("__substg1.0_0037001F", 2, utf16(subject), 0),
        ("__substg1.0_1000001F", 2, utf16(body), 0),
        ("__attach_version1.0_#00000000", 1, b"", 0),
        ("__properties_version1.0", 2, attach_props, 8),
        ("__substg1.0_3707001F", 2, utf16(attach_name), 8),
        ("__substg1.0_37010102", 2, attach_data, 8),
    ]
    ids = range(len(nodes))
    left = {i: _NOSTREAM for i in ids}
    right = {i: _NOSTREAM for i in ids}
    child = {i: _NOSTREAM for i in ids}
    for parent in ids:
        kids = [i for i in ids if nodes[i][3] == parent]
        if kids:
            child[parent] = kids[0]
            for a, b in zip(kids, kids[1:]):
                right[a] = b

    # Ministream layout: pack each stream into 64-byte mini sectors
    ministream = b""
    minifat: List[int] = []
    starts: Dict[int, int] = {}
    for i in ids:
        _name, etype, payload, _parent = nodes[i]
        if etype != 2 or not payload:
            continue
        n_mini = (len(payload) + 63) // 64
        starts[i] = len(minifat)
        for k in range(n_mini):
            minifat.append(len(minifat) + 1 if k < n_mini - 1 else _ENDOFCHAIN)
        ministream += payload + b"\x00" * (n_mini * 64 - len(payload))

    # Regular sector plan: 0=FAT, 1..d=directory, d+1=miniFAT, d+2..=ministream
    dir_sectors = (len(nodes) * 128 + 511) // 512
    mini_sectors = (len(ministream) + 511) // 512
    ministream_padded = ministream + b"\x00" * (mini_sectors * 512 - len(ministream))
    first_dir_sect, minifat_sect = 1, 1 + dir_sectors
    first_mini_sect = minifat_sect + 1

    fat = [_FREESECT] * 128
    fat[0] = _FATSECT
    for k in range(dir_sectors):               # directory chain
        fat[first_dir_sect + k] = (first_dir_sect + k + 1
                                   if k < dir_sectors - 1 else _ENDOFCHAIN)
    fat[minifat_sect] = _ENDOFCHAIN            # mini FAT chain
    for k in range(mini_sectors):              # ministream chain
        fat[first_mini_sect + k] = (first_mini_sect + k + 1
                                    if k < mini_sectors - 1 else _ENDOFCHAIN)

    def dirent(name: str, etype: int, l: int, r: int, c: int, start: int, size: int) -> bytes:
        nm = name.encode("utf-16-le")
        blob = (nm + b"\x00\x00").ljust(64, b"\x00")
        blob += struct.pack("<HBB3I", len(nm) + 2, etype, 1, l, r, c)
        blob += b"\x00" * 16                  # CLSID
        blob += b"\x00" * 4                   # state bits
        blob += b"\x00" * 16                  # ctime + mtime
        blob += struct.pack("<IQ", start, size)
        return blob

    directory = b""
    for i in ids:
        name, etype, payload, _parent = nodes[i]
        if etype == 5:
            directory += dirent(name, 5, _NOSTREAM, _NOSTREAM, child[i],
                                first_mini_sect, len(ministream_padded))  # root: ministream
        elif etype == 1:
            directory += dirent(name, 1, left[i], right[i], child[i], 0, 0)
        else:
            directory += dirent(name, 2, left[i], right[i], _NOSTREAM,
                                starts.get(i, _ENDOFCHAIN), len(payload))
    directory = directory.ljust(dir_sectors * 512, b"\x00")       # pad w/ free entries

    minifat_sector = b"".join(struct.pack("<I", v) for v in minifat)
    minifat_sector = minifat_sector.ljust(512, b"\xff")           # FREESECT fill

    header = _CFB_MAGIC + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6)        # ver, order, shifts
    header += b"\x00" * 6
    header += struct.pack("<8I", 0, 1, first_dir_sect, 0, 4096,
                          minifat_sect, 1, _ENDOFCHAIN)
    # ^ num_dir(v4 only), num_fat, first_dir, txn, mini_cutoff,
    #   first_minifat, num_minifat, first_difat
    header += struct.pack("<I", 0)                                # num_difat
    header += struct.pack("<I", 0)                                # DIFAT[0] -> FAT sector 0
    header += b"\xff" * 4 * 108                                   # remaining DIFAT free
    assert len(header) == 512

    fat_sector = b"".join(struct.pack("<I", v) for v in fat)
    return header + fat_sector + directory + minifat_sector + ministream_padded


def _build_catalog_db(path: str):
    conn = sqlite3.connect(path)
    conn.execute("CREATE TABLE products (part_number TEXT, description TEXT, manufacturer TEXT)")
    conn.executemany("INSERT INTO products VALUES (?,?,?)", [
        ("TEST-PART-1", "Test widget, 1in", "Acme"),
        ("SIM-500-BK", "500 MCM THHN black copper", "Southwire"),
        ("VS450", "4in square box, deep", "Raco"),
    ])
    conn.commit()
    conn.close()


def selftest(cfg: Config) -> int:
    import tempfile
    passed, failed = 0, 0
    tmpdir = tempfile.mkdtemp(prefix="omniscan_selftest_")

    def check(label: str, cond: bool, detail: str = ""):
        nonlocal passed, failed
        if cond:
            passed += 1
            print(f"  PASS  {label}")
        else:
            failed += 1
            print(f"  FAIL  {label}  {detail}")

    print(f"{APP_TITLE} selftest ({tmpdir})")

    # txt
    res = scan_bytes(b"25 EA VS450 4SQ BOX DEEP\n2) 500 FT SIM-500-BK 500 MCM THHN BLACK\n", "test.txt", cfg)
    check("txt: format", res.fmt == "text", res.fmt)
    check("txt: line items >= 2", len(res.line_items) >= 2, str(len(res.line_items)))
    qty_ok = any(li.qty == 25 for li in res.line_items) and any(li.qty == 500 for li in res.line_items)
    check("txt: qty parsed (25, 500)", qty_ok, json.dumps([li.to_dict() for li in res.line_items]))
    check("txt: gauge detected (500 MCM)", any(li.gauge == "500 MCM" for li in res.line_items))

    # line-number repair
    res = scan_bytes(b"1) 10 EA TEST-PART-1 WIDGET\n2 100 FT SIM-500-BK THHN\n", "rfq.txt", cfg)
    repaired = [li for li in res.line_items if li.qty == 100]
    check("parser: line-number-as-qty repair", bool(repaired),
          json.dumps([li.to_dict() for li in res.line_items]))

    # source-line linkage + single-line parse (review-loop support)
    res = scan_bytes(b"header junk\n25 EA VS450 4SQ BOX\n", "ln.txt", cfg)
    check("parser: line_no recorded",
          bool(res.line_items) and res.line_items[0].line_no == 1,
          json.dumps([li.to_dict() for li in res.line_items]))
    one = LineItemParser(cfg).parse_one("7 EA TEST-PART-1 WIDGET")
    check("parser: parse_one", one is not None and one.qty == 7,
          "None" if one is None else json.dumps(one.to_dict()))

    # OCR table repro (Steve's Tri-State test PO failure): line-number column,
    # 'i' misread for '1', pipe/hash column-border junk, thousands separator
    ocr_doc = (b"LineQty Unit | Description Part # Notes\n"
               b"i 300 FT 1/2 threaded rod galvanized ATR-056\n"
               b"2 2,000 EA Deep kindorf strap 1-5/8 inch KOS-158\n"
               b"3 500 EA 1/2 hex nuts zinc plated HN-050\n"
               b"4 |500 EA 1/2 flat washer FW-050\n"
               b"5 #1 EA 4 inch square box 2-1/8 deep BOX-4218\n"
               b"6 250 FT 12 THHN stranded green copper THHN-12-GRN\n")
    res = scan_bytes(ocr_doc, "ocr.txt", cfg)
    got = {li.qty for li in res.line_items}
    check("ocr-repair: line-number column stripped (qtys 300/2000/500/500/1/250)",
          got == {300.0, 2000.0, 500.0, 1.0, 250.0},
          json.dumps([li.to_dict() for li in res.line_items]))
    box_item = next((li for li in res.line_items if li.part == "BOX-4218"), None)
    check("ocr-repair: '#1 EA box' has qty 1 and NO false gauge",
          box_item is not None and box_item.qty == 1 and box_item.gauge is None,
          "missing" if box_item is None else json.dumps(box_item.to_dict()))
    wire_item = next((li for li in res.line_items if li.part == "THHN-12-GRN"), None)
    check("ocr-repair: wire context keeps 12 AWG gauge",
          wire_item is not None and wire_item.gauge == "12 AWG",
          "missing" if wire_item is None else json.dumps(wire_item.to_dict()))
    check("ocr-repair: raw preserves original line (pipes intact) + line_no",
          any(li.raw.startswith("4 |500") and li.line_no == 4 for li in res.line_items),
          json.dumps([li.to_dict() for li in res.line_items]))
    frag = LineItemParser(cfg).parse_one("12 THHN stranded green copper THHN-12-GRN")
    check("parser: wire-size fragment never becomes qty",
          frag is not None and frag.qty is None and frag.gauge == "12 AWG"
          and frag.part == "THHN-12-GRN",
          "None" if frag is None else json.dumps(frag.to_dict()))

    # garbled handwritten markers: '1.' OCR'd as 16, '2.' as 26 ... (stride 10)
    garbled = (b"16 25 EA EMT CONNECTORS CON-34\n"
               b"26 40 EA EMT COUPLINGS CPL-34\n"
               b"36 100 FT 12 THHN BLACK WIRE THHN-12\n"
               b"46 10 EA SQUARE BOXES SQ-4\n")
    res = scan_bytes(garbled, "hw.txt", cfg)
    check("stride-run: garbled dot markers stripped (qtys 25/40/100/10)",
          {li.qty for li in res.line_items} == {25.0, 40.0, 100.0, 10.0},
          json.dumps([li.to_dict() for li in res.line_items]))
    one = LineItemParser(cfg).parse_one("10, 500 FT 1/2 FLEX CONDUIT FC-050")
    check("parser: comma line-number marker",
          one is not None and one.qty == 500,
          "None" if one is None else json.dumps(one.to_dict()))
    one = LineItemParser(cfg).parse_one("25 EG 3/4 EMT CONNECTORS CC-075")
    check("parser: handwriting UOM alias eg->EA",
          one is not None and one.qty == 25 and one.uom == "EA",
          "None" if one is None else json.dumps(one.to_dict()))
    one = LineItemParser(cfg).parse_one("5 EA - 1/2 FLEX CONDUIT FC-012")
    check("parser: leading dash stripped from description",
          one is not None and one.description.startswith("1/2"),
          "None" if one is None else json.dumps(one.to_dict()))
    one = LineItemParser(cfg).parse_one("1. 25ea - 3/4 EMT connectors")
    check("parser: glued qty+UOM split (25ea)",
          one is not None and one.qty == 25 and one.uom == "EA",
          "None" if one is None else json.dumps(one.to_dict()))
    one = LineItemParser(cfg).parse_one("3. 100 ft- 12 THHN black wire")
    check("parser: dash-glued UOM (ft-)",
          one is not None and one.qty == 100 and one.uom == "FT"
          and one.gauge == "12 AWG",
          "None" if one is None else json.dumps(one.to_dict()))

    # REAL-WORLD BENCHMARK: verbatim tesseract output from a slanted
    # handwritten job material list (garbled list markers, euro-a, etc.)
    hw = ("Tob material _tict\n"
          "16 25 ea \u2014 3/4\" EMT connectors\n"
          "2e 40 ea \u2014 3/4\" EMT couplings\n"
          "3. 100 ft \u2014 #12 THHN black\n"
          "4, 100 ft \u2014 #12 THHN white\n"
          "5. 100 ft \u2014 #12 THHN q@reen\n"
          "b. 10 \u20aca \u2014 4\" square boxes\n"
          "1. 10 \u20aca \u2014 4* square blank covers\n"
          "g, 6 ea \u2014 single gang mud rings\n"
          "% 2 \u20aca \u2014 100A 3P breakers\n"
          "40, 500 ft \u2014 1/2\" flex conduit\n"
          "\u201c. 1 voll \u2014 red electrical tape\n"
          "12. 12 ea \u2014 1\u00b0 PVC straps\n"
          "13, 20 ff \u2014 2\u00b0 vigid conduit\n"
          "4, \u00a2 ea \u2014 2\u00b0 rigid couplings\n"
          "15. 4 ea \u2014 2\" vigid locknuts\n"
          "be 6 Ca \u2014 3/4* LB bodies\n"
          "11, 3 bags \u2014 blue wire nuts\n"
          "%, 1\u20aca \u2014 24 x 24 x 6 Pull box\n").encode("utf-8")
    res = scan_bytes(hw, "handwriting.txt", cfg)
    hw_qtys = [li.qty for li in res.line_items]
    expected = [25.0, 40.0, 100.0, 100.0, 100.0, 10.0, 10.0, 6.0, 2.0,
                500.0, 1.0, 12.0, 20.0, 4.0, 6.0, 3.0, 1.0]
    hit = sum(1 for q in expected if q in hw_qtys and hw_qtys.remove(q) is None)
    check("handwriting benchmark: >=16 of 17 readable qtys correct",
          hit >= 16, f"hit={hit} items="
          + json.dumps([li.to_dict() for li in res.line_items]))
    check("handwriting benchmark: inch marks normalized in descriptions",
          any('1" PVC straps' in (li.description or "") for li in res.line_items)
          and not any("\u00b0" in (li.description or "") for li in res.line_items),
          json.dumps([li.description for li in res.line_items]))
    check("handwriting benchmark: vigid->rigid word fix",
          not any("vigid" in (li.description or "") for li in res.line_items),
          json.dumps([li.description for li in res.line_items]))
    pull = next((li for li in res.line_items
                 if "pull box" in (li.description or "").lower()), None)
    check("handwriting benchmark: pull-box row parses (1e\u20aca glue)",
          pull is not None and pull.qty == 1 and pull.uom == "EA",
          "missing" if pull is None else json.dumps(pull.to_dict()))

    # teach persistence: save failures must surface, not vanish
    bad = TeachStore("/proc/omniscan_cannot_write_here.json")
    bad.set_rule("X 1 EA THING", {"qty": 1})
    check("teach: unwritable path surfaces last_error",
          bad.last_error is not None, str(bad.last_error))
    good = TeachStore(os.path.join(tmpdir, "ok_teach.json"))
    good.set_rule("X 1 EA THING", {"qty": 1})
    check("teach: writable path clears last_error", good.last_error is None,
          str(good.last_error))

    # PAREN-FORMAT BENCHMARK: verbatim OCR lines from a real photocopied
    # material request (parenthesized quantities, highlighter damage, wraps)
    paren = ("Danny\n"
             "Subject: FW: 1595-Material Request #1.\n"
             "-(3) Cases of Water\n"
             "-(3 boxes) \"TOPAZ\" 170's\n"
             "-(1 coil) 3/4\" UL Rated GF\n"
             "G400') 3/4\" EMT\n"
             "-(1box) 3/4\" EMT one-hole straps\n"
             "~(3 boxes) DEWALT-(PFM2211200)\n"
             "L( box) DEWALT-(1/4\" x 1-3/8\" long Acorn Anchors\n"
             "-(2 cases/100pcs.) 4-11/16\" x 5/8\" Deep Raised SIngle-Gang\n"
             "cover plates\n"
             "-(500pcs.) 3M Red/Yellow wire nuts\n"
             "-(2 dozen) Work Gloves\n"
             "-(1 can) Cherry Red Spray Paint\n"
             "-(1) \"ITM\"-(Part#HF0016) 7/8\" Tungsten Carbide tipped hole\n"
             "cutter saws, 3/16\" deep\n"
             "2\n").encode("utf-8")
    pres = scan_bytes(paren, "paren.txt", cfg)
    def find_kw(kw):
        return next((li for li in pres.line_items
                     if kw in (li.description or "").lower()), None)
    li = find_kw("cases of water")
    check("paren: bare qty defaults to EA",
          li is not None and li.qty == 3 and li.uom == "EA",
          str(li and li.to_dict()))
    li = find_kw("topaz")
    check("paren: qty+boxes uom", li is not None and li.qty == 3 and li.uom == "BOX",
          str(li and li.to_dict()))
    li = find_kw("ul rated")
    check("paren: coil uom", li is not None and li.qty == 1 and li.uom == "COIL",
          str(li and li.to_dict()))
    li = find_kw("emt") if find_kw("emt") else None
    li = next((x for x in pres.line_items if x.qty == 400), None)
    check("paren: garbled-paren apostrophe-feet (G400')",
          li is not None and li.uom == "FT", str(li and li.to_dict()))
    li = find_kw("one-hole")
    check("paren: glued (1box)", li is not None and li.qty == 1 and li.uom == "BOX",
          str(li and li.to_dict()))
    li = find_kw("dewalt-(pfm")
    check("paren: tilde marker + part", li is not None and li.part == "PFM2211200",
          str(li and li.to_dict()))
    li = find_kw("acorn")
    check("paren: qty-less review row (L( box))",
          li is not None and li.qty is None and li.uom == "BOX",
          str(li and li.to_dict()))
    res_blank = scan_bytes(b"L( box) mystery widget line here\n"
                           b"-(3) first\n-(4) second\n-(5) third\n", "b.txt", cfg)
    blanks = [li for li in res_blank.line_items if li.qty is None]
    check("default-uom: qty-less rows keep blank UOM unless parens said otherwise",
          all(li.uom in (None, "BOX") for li in blanks),
          json.dumps([li.to_dict() for li in res_blank.line_items]))
    li = find_kw("single-gang")
    check("paren: pack note + continuation merge",
          li is not None and li.qty == 2 and li.uom == "CASE"
          and "cover plates" in (li.description or "")
          and "100 pcs" in (li.description or "").replace("100pcs", "100 pcs"),
          str(li and li.to_dict()))
    li = find_kw("wire nuts")
    check("paren: (500pcs.)", li is not None and li.qty == 500 and li.uom == "PCS",
          str(li and li.to_dict()))
    li = find_kw("work gloves")
    check("paren: dozen -> DZ", li is not None and li.qty == 2 and li.uom == "DZ",
          str(li and li.to_dict()))
    li = find_kw("tungsten")
    check("paren: explicit Part# + continuation",
          li is not None and li.part == "HF0016"
          and "cutter saws" in (li.description or ""),
          str(li and li.to_dict()))
    check("paren: page number not swallowed",
          not any((li.raw or "").strip() == "2" for li in pres.line_items),
          json.dumps([li.raw for li in pres.line_items]))

    one = LineItemParser(cfg).parse_one("2x 4\" square box VS450")
    check("parser: Nx multiplier", one is not None and one.qty == 2
          and one.part == "VS450", str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("QTY: 12 EA VS450 4SQ BOX")
    check("parser: QTY label", one is not None and one.qty == 12 and one.uom == "EA",
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("6a 4\" round boxes")
    check("parser: handwritten-ea glue (6a)",
          one is not None and one.qty == 6 and one.uom == "EA"
          and "round boxes" in one.description,
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("1a panel schedule sleeve")
    check("parser: handwritten-ea glue (1a)",
          one is not None and one.qty == 1 and one.uom == "EA",
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("2 ea 100A 3P breakers")
    check("parser: mid-line amp rating untouched",
          one is not None and one.qty == 2 and "100A" in one.description,
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("1.25 EA WIDGET AAA-1")
    check("parser: marker glued to qty (1.25 EA = 1. + 25 EA)",
          one is not None and one.qty == 25 and one.uom == "EA",
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("4.7 EA THING DDD-4")
    ft1 = LineItemParser(cfg).parse_one("150 fi 10 THHN green")
    check("parser: ft misread alias (fi)",
          ft1 is not None and ft1.qty == 150 and ft1.uom == "FT"
          and ft1.gauge == "10 AWG",
          "None" if ft1 is None else str(ft1.to_dict()))
    ft2 = LineItemParser(cfg).parse_one("150 f 8 THHN red")
    check("parser: ft misread alias (f)",
          ft2 is not None and ft2.uom == "FT" and ft2.gauge == "8 AWG",
          "None" if ft2 is None else str(ft2.to_dict()))
    ft3 = LineItemParser(cfg).parse_one("150 ft i 10 THHN green")
    check("parser: lone-i OCR noise stripped from desc",
          ft3 is not None and ft3.description.startswith("10 THHN"),
          "None" if ft3 is None else str(ft3.to_dict()))
    ft4 = LineItemParser(cfg).parse_one("150 10 THHN green")
    check("parser: missing unit on wire line defaults FT not EA",
          ft4 is not None and ft4.uom == "FT" and ft4.gauge == "10 AWG",
          "None" if ft4 is None else str(ft4.to_dict()))
    ft5 = LineItemParser(cfg).parse_one("36 toggle switch plates")
    check("parser: missing unit on non-wire line still defaults EA",
          ft5 is not None and ft5.uom == "EA",
          "None" if ft5 is None else str(ft5.to_dict()))
    # HARD-SCAN BENCHMARK: verbatim OCR lines from a faint 2-column RFQ
    hp = LineItemParser(cfg)
    hl = hp.parse_one("7 Qty 8 / EA - 1900 BOX W/ 1/2 KO")
    check("hard-scan: Qty N / UOM grammar",
          hl is not None and hl.qty == 8 and hl.uom == "EA",
          str(hl and hl.to_dict()))
    hl = hp.parse_one("4.12 EA 1/2 EMT CPLG COMP")
    check("hard-scan: marker-glued count qty (4.12 EA)",
          hl is not None and hl.qty == 12 and hl.uom == "EA",
          str(hl and hl.to_dict()))
    hres = scan_bytes(b"9. 100 FT #12 THHN BLACK\n"
                      b"10, 100 FT #12 THHN WHITE\n"
                      b"11.100 FT #12 THHN GREEN\n"
                      b"12. 250 FT #10 THHN RED\n"
                      b"13. 75 FT 12/2 MC CABLE ALUM\n"
                      b": 14.60 FT 3/4 LIQUIDTIGHT FLEX\n"
                      b"15. 10 EA 3/4 LT CONN STRAIGHT\n", "h.txt", cfg)
    hq = [li.qty for li in hres.line_items]
    check("hard-scan: marker-glued measured qty via expected lineno",
          hq.count(100.0) == 3 and 60.0 in hq and 75.0 in hq,
          json.dumps([li.to_dict() for li in hres.line_items]))
    hl = hp.parse_one("10, 100f #12 THHN WHITE")
    check("hard-scan: glued lone-f (100f) -> 100 FT",
          hl is not None and hl.qty == 100 and hl.uom == "FT",
          str(hl and hl.to_dict()))
    hl = hp.parse_one("25. tea 45KVA XFMR 480-208Y/120")
    check("hard-scan: tea -> 1 ea (t/1 confusion)",
          hl is not None and hl.qty == 1 and hl.uom == "EA",
          str(hl and hl.to_dict()))
    hres = scan_bytes(b"3. 12 EA 1/2 EMT CONN COMP\n"
                      b"9. 100 FT #12 THHN BLACK\n"
                      b"19) -(3 BOX) 1/4-20 SPRING NUTS\n"
                      b"22. 5 EA 13/16 STRUT 90 BRACKET\n"
                      b"'43. 7 EA'2 IN EMT CPLG COMP\n"
                      b"- 48 1000 FT #14 THHN GRAY\n", "h2.txt", cfg)
    hitems = {li.qty: li for li in hres.line_items}
    check("hard-scan: marker+paren combo (19) -(3 BOX))",
          3.0 in hitems and hitems[3.0].uom == "BOX",
          json.dumps([li.to_dict() for li in hres.line_items]))
    check("hard-scan: quote-prefixed marker + letter'digit unglue",
          7.0 in hitems and "2 IN EMT" in (hitems[7.0].description or ""),
          json.dumps([li.to_dict() for li in hres.line_items]))
    check("hard-scan: two-leading-numbers (- 48 1000 FT)",
          1000.0 in hitems and hitems[1000.0].uom == "FT",
          json.dumps([li.to_dict() for li in hres.line_items]))
    hl = hp.parse_one("25. 1ea 45KVA XFMR 480-208Y/120")
    check("hard-scan: voltage notation never becomes a part",
          hl is not None and hl.qty == 1 and hl.part is None,
          str(hl and hl.to_dict()))
    nres = scan_bytes(b"1. 14 EA 3/4 EMT CONN SS\n2. 25 EA 3/4 EMT CPLG SS\n"
                      b"3. 12 EA 1/2 EMT CONN COMP\n4. 12 EA 1/2 EMT CPLG COMP\n"
                      b"as noted\n", "n.txt", cfg)
    check("hard-scan: digit-less margin debris never becomes a review row",
          not any("noted" in (li.description or "") for li in nres.line_items),
          json.dumps([li.to_dict() for li in nres.line_items]))
    hl = hp.parse_one("41, 150' 3/8 THREADED ROD")
    check("hard-scan: bare apostrophe-feet after comma marker",
          hl is not None and hl.qty == 150 and hl.uom == "FT",
          str(hl and hl.to_dict()))
    thou = LineItemParser(cfg).parse_one("2,000 EA Deep kindorf strap KOS-158")
    check("parser: thousands qty survives glued-marker logic (2,000 EA)",
          thou is not None and thou.qty == 2000,
          "None" if thou is None else str(thou.to_dict()))
    check("parser: marker glued to qty single-digit (4.7 EA)",
          one is not None and one.qty == 7,
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("2.5 FT CABLE CCC-3")
    check("parser: fractional FT stays fractional",
          one is not None and one.qty == 2.5 and one.uom == "FT",
          str(one and one.to_dict()))
    one = LineItemParser(cfg).parse_one("20A duplex receptacles")
    check("parser: uppercase amp rating never becomes a qty",
          one is None or one.qty != 20,
          "None" if one is None else str(one.to_dict()))

    # void/crossed-out decoy suppression
    vres = scan_bytes(b"1. 25 EA VS450 4SQ BOX\n"
                      b"VOID - 99 EA old-style boxes - do not quote\n"
                      b"changed: 18 EA 1/2 locknuts - ignore this note\n"
                      b"scratch: 200 ft #12 red was already ordered\n"
                      b"2. 10 EA TEST-PART-1 WIDGET\n"
                      b"3. 4 EA TP248 THING\n", "void.txt", cfg)
    check("void: decoys suppressed, real items kept",
          {li.qty for li in vres.line_items} == {25.0, 10.0, 4.0}
          and len(vres.meta.get("voided_lines", [])) == 3
          and any("suppressed 3" in w for w in vres.warnings),
          json.dumps(vres.to_dict()))

    # fuzzy teach: OCR jitter still matches; suppress rules stay strict
    fstore = TeachStore(os.path.join(tmpdir, "fuzzy.json"))
    fstore.set_rule("16. 1 CASE case of clear safety glasses SG-CL",
                    {"qty": 1, "uom": "CASE", "part": "SG-CL"})
    hit = fstore.get("16, 1 CASE case of c1ear safety glasses SG-CL", fuzzy=0.86)
    check("teach: fuzzy match tolerates OCR jitter",
          hit is not None and hit.get("part") == "SG-CL", str(hit))
    fstore.suppress("99 EA definitely bad line of stuff")
    near = fstore.get("99 EA definitely bad line of Stuff xx", fuzzy=0.86)
    check("teach: fuzzy suppress requires stricter ratio",
          near is None or near.get("action") != "suppress"
          or fstore.get("99 EA definitely bad line of stuff", fuzzy=0.86) is not None,
          str(near))

    # marker-doc grammar additions from the audit document
    mres = scan_bytes(b"1. 300 FT 3/4 EMT conduit EMT-075\n"
                      b"2. 25 EA 4SQ deep box VS450\n"
                      b"6. (2 roll) caution tape red electrical - TAPE-RED\n"
                      b"10; 200' 2 inch PVC schedule 40 PVC40-200\n"
                      b"24. 2.CAN cold galvanizing spray CGS-16\n"
                      b"ff! 3 BAG red/yellow wire nuts RY-WN\n"
                      b"ok 30. 30ea 30A 2 pole breaker BR230BO\n"
                      b"ground bar kit 12 terminal GBK12\n", "hard.txt", cfg)
    def mfind(kw):
        return next((li for li in mres.line_items
                     if kw in ((li.part or "") + " " + (li.description or ""))), None)
    li = mfind("caution tape")
    check("marker-doc: paren qty after marker (2 roll)",
          li is not None and li.qty == 2 and li.uom == "ROLL", str(li and li.to_dict()))
    li = mfind("PVC40-200")
    check("marker-doc: semicolon marker + bare apostrophe-feet",
          li is not None and li.qty == 200 and li.uom == "FT", str(li and li.to_dict()))
    li = mfind("CGS-16")
    check("marker-doc: period-glued UOM (2.CAN)",
          li is not None and li.qty == 2 and li.uom == "CAN", str(li and li.to_dict()))
    li = mfind("wire nuts")
    check("marker-doc: garbled letter marker (ff!)",
          li is not None and li.qty == 3 and li.uom == "BAG", str(li and li.to_dict()))
    li = mfind("BR230BO")
    check("marker-doc: annotation prefix (ok) + glued 30ea",
          li is not None and li.qty == 30 and li.uom == "EA", str(li and li.to_dict()))
    li = mfind("GBK12")
    check("marker-doc: shredded-line part salvage as review row",
          li is not None and li.qty is None and li.confidence == 0.45,
          str(li and li.to_dict()))

    # per-page OCR fallback: mixed text+scan PDF
    if DEPS.pypdf or DEPS.pdfplumber:
        mixed = scan_bytes(_build_two_page_pdf(
            "12 EA TEST-PART-1 WIDGET for the mixed text-plus-scan page test"),
            "mixed.pdf", cfg)
        check("pdf: mixed text+scan engine label",
              str(mixed.meta.get("engine", "")).startswith("mixed"),
              str(mixed.meta))
        check("pdf: mixed text page still parses",
              any(li.qty == 12 for li in mixed.line_items),
              json.dumps(mixed.to_dict()))

    # column split: synthetic two-column page (needs OCR)
    col_ocr_ready = False
    try:
        DEPS.pytesseract.get_tesseract_version()
        col_ocr_ready = DEPS.PIL is not None
    except Exception:
        pass
    if col_ocr_ready:
        from PIL import Image as _Img, ImageDraw as _Draw, ImageFont as _Font
        two = _Img.new("L", (2000, 400), 255)
        _d = _Draw.Draw(two)
        try:
            _f = _Font.load_default(size=40)
        except TypeError:
            _f = None
        _d.text((60, 60), "1. 25 EA WIDGET AAA-1", font=_f, fill=0)
        _d.text((60, 160), "2. 40 EA GADGET BBB-2", font=_f, fill=0)
        _d.text((1150, 60), "3. 100 FT CABLE CCC-3", font=_f, fill=0)
        _d.text((1150, 160), "4. 7 EA THING DDD-4", font=_f, fill=0)
        split = _find_column_split(two)
        check("columns: gutter detected on two-column page",
              split is not None and 800 < split < 1200, str(split))
        buf2 = io.BytesIO()
        two.save(buf2, format="PNG")
        cres = scan_bytes(buf2.getvalue(), "twocol.png", cfg)
        check("columns: both columns extracted",
              {li.qty for li in cres.line_items} >= {25.0, 40.0, 100.0, 7.0}
              and cres.meta.get("columns") == 2,
              json.dumps(cres.to_dict()))
        single = _Img.new("L", (2000, 400), 255)
        _d = _Draw.Draw(single)
        _d.text((60, 60), "1. 25 EA WIDGET AAA-1 long single column line", font=_f, fill=0)
        _d.text((60, 160), "2. 40 EA GADGET BBB-2 more single column text", font=_f, fill=0)
        check("columns: single-column page not split",
              _find_column_split(single) is None, str(_find_column_split(single)))

    check("deps: tesseract binary probe reports status",
          "tesseract binary" in DEPS.report(), DEPS.report())

    # catalog O/0 confusion
    db_o0 = os.path.join(tmpdir, "cat_o0.db")
    _build_catalog_db(db_o0)
    hit = Catalog(db_o0).lookup("VS45O")   # letter O for zero
    check("catalog: O/0 confusion fallback",
          hit is not None and hit.get("match") == "ocr-o0", str(hit))
    check("handwriting benchmark: all 18 rows surfaced (incl. review rows)",
          len(res.line_items) == 18, str(len(res.line_items)))
    check("handwriting benchmark: unreadable-qty row surfaced for review",
          any(li.qty is None and "rigid couplings" in (li.description or "")
              for li in res.line_items),
          json.dumps([li.to_dict() for li in res.line_items]))

    # teach layer: set-rule, suppress-rule, forget
    teach_path = os.path.join(tmpdir, "teach.json")
    store = TeachStore(teach_path)
    store.set_rule("wire assembly per spec sheet A", {"qty": 40, "uom": "EA",
                                                      "part": "ASSY-A", "description": "Wire assembly A"})
    store.suppress("25 EA VS450 4SQ BOX")
    tcfg = Config(teach_path=teach_path, ocr=cfg.ocr)
    res = scan_bytes(b"wire assembly per spec sheet A\n25 EA VS450 4SQ BOX\n", "t.txt", tcfg)
    taught = [li for li in res.line_items if li.taught]
    check("teach: set-rule emits taught item",
          len(taught) == 1 and taught[0].qty == 40 and taught[0].part == "ASSY-A",
          json.dumps([li.to_dict() for li in res.line_items]))
    check("teach: suppress-rule kills heuristic item",
          not any(li.qty == 25 for li in res.line_items),
          json.dumps([li.to_dict() for li in res.line_items]))
    store.forget("25 EA VS450 4SQ BOX")
    res = scan_bytes(b"25 EA VS450 4SQ BOX\n", "t2.txt", tcfg)
    check("teach: forget restores heuristics",
          any(li.qty == 25 for li in res.line_items),
          json.dumps([li.to_dict() for li in res.line_items]))

    # csv
    res = scan_bytes(b"QTY,UOM,PART,DESCRIPTION\n12,EA,VS450,4SQ BOX\n", "items.csv", cfg)
    check("csv: format + item", res.fmt == "csv" and any(li.qty == 12 for li in res.line_items),
          json.dumps(res.to_dict()))

    # docx
    res = scan_bytes(_build_docx(["RFQ 4512", "6 EA TEST-PART-1 WIDGET"]), "rfq.docx", cfg)
    check("docx: format", res.fmt == "docx", res.fmt)
    check("docx: item qty 6", any(li.qty == 6 for li in res.line_items), res.text)

    # xlsx
    res = scan_bytes(_build_xlsx([["QTY", "PART", "DESCRIPTION"],
                                  ["40", "VS450", "4SQ BOX DEEP"]]), "rfq.xlsx", cfg)
    check("xlsx: format", res.fmt == "xlsx", res.fmt)
    check("xlsx: item qty 40", any(li.qty == 40 for li in res.line_items), res.text)

    # html
    res = scan_bytes(b"<html><body><p>15 EA TEST-PART-1 WIDGET</p></body></html>", "q.html", cfg)
    check("html: item qty 15", any(li.qty == 15 for li in res.line_items), res.text)

    # eml with csv attachment (recursion)
    eml = _build_eml("Please quote:\n8 EA VS450 4SQ BOX\n",
                     "items.csv", b"QTY,PART\n30,TEST-PART-1\n")
    res = scan_bytes(eml, "rfq.eml", cfg)
    check("eml: format", res.fmt == "eml", res.fmt)
    check("eml: body item qty 8", any(li.qty == 8 for li in res.line_items), res.text)
    check("eml: attachment scanned", len(res.children) == 1, str(len(res.children)))
    check("eml: attachment item qty 30",
          any(li.qty == 30 for li in res.all_line_items()), json.dumps(res.to_dict()))

    # zip recursion
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("inner/rfq.txt", "9 EA VS450 4SQ BOX\n")
    res = scan_bytes(buf.getvalue(), "bundle.zip", cfg)
    check("zip: child item qty 9", any(li.qty == 9 for li in res.all_line_items()),
          json.dumps(res.to_dict()))

    # native .msg (from-scratch CFB writer -> from-scratch CFB/MSG reader)
    msg_bytes = _build_msg_file("RFQ 991", "Please quote:\n14 EA VS450 4SQ BOX\n",
                                "items.csv", b"QTY,PART\n21,TEST-PART-1\n")
    res = scan_bytes(msg_bytes, "rfq.msg", cfg)
    check("msg: native engine", res.fmt == "ole" and res.meta.get("engine") == "native-cfb",
          f"{res.fmt} {res.meta} {res.warnings}")
    check("msg: subject read", "RFQ 991" in res.text, res.text)
    check("msg: body item qty 14", any(li.qty == 14 for li in res.line_items), res.text)
    check("msg: attachment scanned", len(res.children) == 1, str(res.warnings))
    check("msg: attachment item qty 21",
          any(li.qty == 21 for li in res.all_line_items()), json.dumps(res.to_dict()))
    res = scan_bytes(_CFB_MAGIC + b"\x00" * 1024, "old.doc", cfg)
    check("msg: legacy OLE degrades to warning", bool(res.warnings) and not res.line_items,
          str(res.warnings))

    # image OCR (only if the tesseract binary is actually installed)
    ocr_ready = False
    if cfg.ocr and DEPS.pytesseract and DEPS.PIL:
        try:
            DEPS.pytesseract.get_tesseract_version()
            ocr_ready = True
        except Exception:
            pass
    if ocr_ready:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new("L", (700, 120), 255)
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.load_default(size=48)
        except TypeError:
            font = None
        draw.text((20, 30), "25 EA VS450 4SQ BOX", font=font, fill=0)
        if font is None:
            img = img.resize((2100, 360))
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        res = scan_bytes(buf.getvalue(), "scan.png", cfg)
        check("image OCR: qty 25", any(li.qty == 25 for li in res.line_items),
              res.text + " | " + "; ".join(res.warnings))
        li0 = res.line_items[0] if res.line_items else None
        boxed = (li0 is not None and res.meta.get("ocr_lines")
                 and str(li0.line_no) in res.meta["ocr_lines"]
                 and len(res.meta["ocr_lines"][str(li0.line_no)]["box"]) == 4)
        check("image OCR: line bounding box emitted and aligned", bool(boxed),
              json.dumps(res.meta.get("ocr_lines", {})))

        # deskew: rotate a 3-row table by 6 degrees, must still extract
        rows = ["1  300 FT THREADED ROD ATR-056",
                "2  500 EA HEX NUTS HN-050",
                "3  250 EA FLAT WASHER FW-050"]
        img = Image.new("L", (1500, 90 + 90 * len(rows)), 255)
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.load_default(size=42)
        except TypeError:
            font = None
        for i, row in enumerate(rows):
            draw.text((60, 40 + 90 * i), row, font=font, fill=0)
        img = img.rotate(-6, expand=True, fillcolor=255)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        res = scan_bytes(buf.getvalue(), "slanted.png", cfg)
        deg = (res.meta.get("deskew_deg") or [0.0])[0]
        check("deskew: ~6 degree slant detected", 4.0 <= abs(deg) <= 8.0,
              f"deskew_deg={deg} warnings={res.warnings}")
        check("deskew: slanted table extracts correct qtys",
              {li.qty for li in res.line_items} >= {300.0, 500.0, 250.0},
              json.dumps([li.to_dict() for li in res.line_items]) + f" text={res.text!r}")
    else:
        print("  SKIP  image OCR (tesseract engine not installed)")

    # pdfium render machinery (used by scanned-PDF OCR fallback)
    if DEPS.pypdfium2:
        try:
            pdf = DEPS.pypdfium2.PdfDocument(_build_minimal_pdf("X"))
            pil = pdf[0].render(scale=2).to_pil()
            pdf.close()
            check("pypdfium2: page render", pil.size[0] > 0 and pil.size[1] > 0, str(pil.size))
        except Exception as exc:
            check("pypdfium2: page render", False, str(exc))
    else:
        print("  SKIP  pypdfium2 render (not installed)")

    # pdf (only if a pdf lib is present)
    if DEPS.pypdf or DEPS.pdfplumber:
        res = scan_bytes(_build_minimal_pdf("12 EA TEST-PART-1 WIDGET"), "rfq.pdf", cfg)
        check("pdf: format", res.fmt == "pdf", res.fmt)
        check("pdf: item qty 12", any(li.qty == 12 for li in res.line_items),
              res.text + " | " + "; ".join(res.warnings))
    else:
        print("  SKIP  pdf (no pypdf/pdfplumber installed)")

    # catalog grounding
    db = os.path.join(tmpdir, "cat.db")
    _build_catalog_db(db)
    catalog = Catalog(db)
    check("catalog: loaded 3 parts", len(catalog.index) == 3,
          f"{len(catalog.index)} | {catalog.warnings}")
    res = scan_bytes(b"5 EA TEST-PART-1 WIDGET\n", "g.txt", cfg)
    catalog.ground(res.line_items)
    check("catalog: exact match grounded",
          bool(res.line_items and res.line_items[0].catalog_match
               and res.line_items[0].catalog_match.get("match") == "exact"),
          json.dumps([li.to_dict() for li in res.line_items]))

    print(f"selftest: {passed} passed, {failed} failed")
    return 0 if failed == 0 else 1


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _print_result(res: ScanResult, indent: int = 0):
    pad = "  " * indent
    print(f"{pad}{res.name}  [{res.fmt}]  ok={res.ok}  items={len(res.line_items)}")
    for warning in res.warnings:
        print(f"{pad}  ! {warning}")
    for li in res.line_items:
        gm = f" gauge={li.gauge}" if li.gauge else ""
        cm = f" catalog={li.catalog_match['part']}({li.catalog_match['match']})" if li.catalog_match else ""
        print(f"{pad}  - qty={li.qty} uom={li.uom or '-'} part={li.part or '-'}"
              f" conf={li.confidence:.2f}{gm}{cm} :: {li.description or li.raw}")
    for child in res.children:
        _print_result(child, indent + 1)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(prog="omniscan", description=APP_TITLE)
    parser.add_argument("paths", nargs="*", help="files to scan")
    parser.add_argument("--json", metavar="OUT", help="write full results JSON to file ('-' for stdout)")
    parser.add_argument("--catalog", metavar="DB", help="SQLite product catalog for grounding")
    parser.add_argument("--teach", metavar="JSON", help="teach-rules JSON file (learned corrections)")
    parser.add_argument("--ollama-host", metavar="URL", help="opt-in Ollama vision host, e.g. http://tillium-bridge:11434")
    parser.add_argument("--ollama-model", default="gemma3", help="vision model tag (default: gemma3)")
    parser.add_argument("--no-ocr", action="store_true", help="disable OCR fallback")
    parser.add_argument("--min-confidence", type=float, default=0.45, help="line-item confidence floor")
    parser.add_argument("--deps", action="store_true", help="show optional dependency status")
    parser.add_argument("--selftest", action="store_true", help="run built-in end-to-end selftest")
    args = parser.parse_args(argv)

    cfg = Config(ocr=not args.no_ocr, ollama_host=args.ollama_host,
                 ollama_model=args.ollama_model, catalog_path=args.catalog,
                 min_item_confidence=args.min_confidence, teach_path=args.teach)

    if args.deps:
        print(DEPS.report())
        return 0
    if args.selftest:
        return selftest(cfg)
    if not args.paths:
        parser.print_help()
        return 2

    catalog = Catalog(args.catalog) if args.catalog else None
    if catalog:
        for warning in catalog.warnings:
            print(f"! {warning}", file=sys.stderr)

    results = []
    for path in args.paths:
        res = scan_path(path, cfg)
        if catalog:
            def ground_tree(r: ScanResult):
                catalog.ground(r.line_items)
                for ch in r.children:
                    ground_tree(ch)
            ground_tree(res)
        results.append(res)
        _print_result(res)

    if args.json:
        payload = json.dumps([r.to_dict() for r in results], indent=2)
        if args.json == "-":
            print(payload)
        else:
            with open(args.json, "w", encoding="utf-8") as fh:
                fh.write(payload)
            print(f"JSON written: {args.json}")
    return 0


# ==========================================================================
# MaINbox SmartScan bridge — OmniScan engine behind the SmartScan surface
# ==========================================================================

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

GUI_APP_TITLE = "OmniScan GUI v0.5.1"  # v0.5.1: manual Qty corrections use the engine thousands-aware parser (typing 1,000 means one thousand, not 1.0) | v0.5.0: teach persistence fix — default teach file now lives next to the script (writability-probed, home-dir fallback) instead of the process CWD, which on double-clicked Windows apps is system32 where writes silently failed; every teach write is verified and a failure raises a loud error dialog instead of losing the correction | v0.4.1: preview pages rotate by the engine's per-page deskew angle so on-image overlay boxes stay aligned on slanted scans | v0.4.0: SmartScan-style document Preview tab — renders the actual PDF pages (pypdfium2) or image and draws the colored rectangle on the document itself via engine OCR line boxes, synced with the extracted-output box; tabs restructured to Preview / Extracted-Merged Text / Warnings & Meta; Description column moved right of Part; smoke test scans a rendered OCR table image end-to-end and asserts correct qtys + on-image overlay | v0.3.0: side-by-side review w/ synced rectangles + teach-on-correct persistence | v0.2.0: review loop (highlight/correct/delete/add) | v0.1.0: initial release

import json
import os
import queue
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict, List, Optional, Tuple

omniscan = sys.modules[__name__]  # single-file bridge: engine is this module

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
        self.title(f"{GUI_APP_TITLE.split(' # ')[0]}  |  engine: "
                   f"{omniscan.GUI_APP_TITLE.split(' # ')[0]}")
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
                GUI_APP_TITLE.split(" # ")[0],
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
                GUI_APP_TITLE.split(" # ")[0],
                f"TEACH SAVE FAILED — this correction was NOT persisted:\n"
                f"{self.teach.path}\n\n{self.teach.last_error}\n\n"
                "Pick a writable teach file location with the … button, "
                "then Apply again.")
            return False
        return True

    def _on_clear(self):
        if self.scanning:
            messagebox.showinfo(GUI_APP_TITLE, "Wait for the current scan to finish.")
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
            messagebox.showinfo(GUI_APP_TITLE, "Nothing to export yet — run a scan first.")
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
            messagebox.showerror(GUI_APP_TITLE, f"Export failed: {exc}")

    def _on_copy_items(self):
        if not self.pairs:
            messagebox.showinfo(GUI_APP_TITLE, "No line items in the current view.")
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
            messagebox.showinfo(GUI_APP_TITLE, "Add files first.")
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
                    messagebox.showwarning(GUI_APP_TITLE, payload)
                elif kind == "done":
                    self._finish_scan(f"Done — {payload} file(s), "
                                      f"{sum(len(r.all_line_items()) for r in self.results)} "
                                      f"item(s). Review left-vs-right; corrections teach the parser.")
                elif kind == "error":
                    self._finish_scan(f"Scan error: {payload}")
                    messagebox.showerror(GUI_APP_TITLE, payload)
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
            messagebox.showinfo(GUI_APP_TITLE, "Select a line item first.")
            return
        qty_raw = self.var_e_qty.get().strip()
        if qty_raw:
            qty = omniscan.LineItemParser._to_qty(qty_raw)
            if qty is None:
                messagebox.showerror(GUI_APP_TITLE, f"Qty is not a number: {qty_raw!r}")
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
            messagebox.showinfo(GUI_APP_TITLE, "Select a line item first.")
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
            messagebox.showinfo(GUI_APP_TITLE, "Click an unmatched line in the "
                                           "Extracted / Merged Text tab first.")
            return
        owner = self.src_res
        lines = (owner.text or "").splitlines()
        if self.picked_line >= len(lines) or not lines[self.picked_line].strip():
            messagebox.showinfo(GUI_APP_TITLE, "That line is empty.")
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


# ---- SmartScan-compatible surface -----------------------------------------
BRIDGE_TITLE = "SmartScan (OmniScan engine v0.10.1)"


def _bridge_data_dir() -> str:
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


def _bridge_find_catalog() -> str:
    """Best-effort discovery of the MaINbox parts catalog for grounding."""
    names = ("american_power_catalog.db",)
    roots = [_bridge_data_dir(), os.path.dirname(_bridge_data_dir()), os.getcwd()]
    env = os.environ.get("MAINBOX_SMARTSCAN_CATALOG", "").strip()
    if env:
        roots.insert(0, os.path.dirname(env))
        names = (os.path.basename(env),) + names
    for root in roots:
        for name in names:
            p = os.path.join(root, name)
            if os.path.exists(p):
                return p
    return ""


def _bridge_config() -> "Config":
    cfg = Config()
    cfg.teach_path = os.path.join(_bridge_data_dir(),
                                  "SmartScan_Embedded_Mainbox_teach.json")
    cat = _bridge_find_catalog()
    if cat:
        cfg.catalog_db = cat
    return cfg


def _bridge_qty_str(qty) -> str:
    if qty is None:
        return ""
    try:
        f = float(qty)
        return str(int(f)) if f == int(f) else f"{f:g}"
    except Exception:
        return str(qty)


def _bridge_row(li, source_path: str = "") -> dict:
    conf = float(getattr(li, "confidence", 0.0) or 0.0)
    if getattr(li, "taught", False):
        status = "Taught"
    elif li.qty is None or conf < 0.6:
        status = "Review"
    else:
        status = "OK"
    notes = []
    if getattr(li, "gauge", None):
        notes.append(str(li.gauge))
    cm = getattr(li, "catalog_match", None)
    if isinstance(cm, dict) and cm.get("part"):
        notes.append(f"catalog:{cm.get('part')}")
    notes.append(f"conf {conf:.2f}")
    return {
        "qty": _bridge_qty_str(li.qty),
        "unit": li.uom or "",
        "description": li.description or li.raw or "",
        "part_number": li.part or "",
        "manufacturer": "",
        "status": status,
        "ai_status": "",
        "review_status": status,
        "source_text": li.raw or "",
        "notes": " | ".join(notes),
        "source_file_path": source_path,
        "source_file_name": os.path.basename(source_path) if source_path else "",
    }


def _bridge_scan_result_payload(results) -> dict:
    rows, warnings, confs = [], [], []
    text_parts = []
    for path, res in results:
        for li in res.line_items:
            rows.append(_bridge_row(li, path))
            confs.append(float(li.confidence or 0.0))
        warnings.extend(res.warnings or [])
        if res.text:
            text_parts.append(f"--- {os.path.basename(path)} ---\n{res.text}")
    avg = round(sum(confs) / len(confs), 2) if confs else ""
    summary = (f"{BRIDGE_TITLE}: {len(rows)} row(s) from {len(results)} file(s)")
    return {"ok": True, "rows": rows, "logs": [], "warnings": warnings[:40],
            "engine_summary": summary, "confidence": avg,
            "merged_text": "\n\n".join(text_parts)}


# ---- in-process fallback API (module import path) --------------------------
class ScanSettings:
    def __init__(self):
        self.mode = "Normal"
        self.enable_cache = False
        self.auto_ai_review_after_scan = True
        self.smart_review_during_scan = True


class SmartScanDB:
    def load_settings(self):
        return ScanSettings()


class _BridgeScanResult:
    def __init__(self, payload):
        self.materials = payload["rows"]
        self.merged_text = payload.get("merged_text", "")
        self.confidence = payload.get("confidence", "")
        self.warnings = payload.get("warnings", [])
        self.engine_summary = payload.get("engine_summary", "")


class ScanEngine:
    def __init__(self, db=None, settings=None, log=None):
        self._log = log or (lambda *_a: None)
        self._cfg = _bridge_config()

    def scan_file(self, path, force_rescan=False):
        try:
            self._log(f"{BRIDGE_TITLE}: scanning {os.path.basename(str(path))}")
        except Exception:
            pass
        res = scan_path(str(path), self._cfg)
        return _BridgeScanResult(_bridge_scan_result_payload([(str(path), res)]))


# ---- themed MaINbox review window ------------------------------------------
class MainboxReviewGUI(OmniScanGUI):
    def __init__(self, files, json_out):
        self._mb_files = list(files)
        self._mb_json_out = json_out
        self._mb_sent = False
        super().__init__()
        self.title(f"{BRIDGE_TITLE} \u2014 MaINbox Review")
        self._mb_apply_theme()
        bar = ttk.Frame(self)
        bar.pack(side="bottom", fill="x", padx=6, pady=(0, 6))
        ttk.Button(bar, text="Done \u2192 Send Back to MaINbox",
                   command=self._mb_send).pack(side="right", padx=4)
        ttk.Label(bar, text="Correct any rows above, then send them back "
                            "to the MaINbox RFQ.").pack(side="left", padx=4)
        self.protocol("WM_DELETE_WINDOW", self._mb_cancel)
        cfg = _bridge_config()
        if getattr(cfg, "catalog_db", ""):
            self.var_catalog.set(cfg.catalog_db)
        self.var_teach.set(cfg.teach_path)
        self._load_teach_store()
        self.after(200, self._mb_autostart)

    def _mb_autostart(self):
        self._add_paths([p for p in self._mb_files if os.path.exists(p)])
        self._start_scan()
        if os.environ.get("MAINBOX_SMARTSCAN_AUTOSEND", "").strip() == "1":
            self._mb_autosend_poll()

    def _mb_autosend_poll(self):
        if self.scanning:
            self.after(400, self._mb_autosend_poll)
        else:
            self._mb_send()

    def _mb_apply_theme(self):
        try:
            theme = json.loads(os.environ.get("MAINBOX_SMARTSCAN_THEME_JSON", "{}"))
        except Exception:
            theme = {}
        if not isinstance(theme, dict) or not theme:
            return
        bg = theme.get("bg") or ""
        fg = theme.get("fg") or ""
        try:
            style = ttk.Style(self)
            if bg:
                self.configure(bg=bg)
                for name in ("TFrame", "TLabelframe", "TLabel", "TCheckbutton"):
                    style.configure(name, background=bg,
                                    foreground=fg or None)
                style.configure("TLabelframe.Label", background=bg,
                                foreground=fg or None)
            if theme.get("entry_bg"):
                style.configure("TEntry", fieldbackground=theme["entry_bg"])
            font_family = theme.get("font_family")
            if font_family:
                size = int(theme.get("font_size") or 10)
                self.option_add("*Font", (font_family, size))
        except Exception:
            pass                                # theming is best-effort only

    def _mb_collect(self):
        flat = []

        def walk(res, path):
            flat.append((path or res.name or "", res))
            for child in (res.children or []):
                walk(child, path or res.name or "")

        for res in self.results:
            walk(res, res.name or "")
        return _bridge_scan_result_payload(flat)

    def _mb_send(self):
        if self._mb_sent:
            return
        payload = self._mb_collect()
        # user edits in the review pane are already applied to the LineItems
        try:
            with open(self._mb_json_out, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=1)
            self._mb_sent = True
        except OSError as exc:
            messagebox.showerror(BRIDGE_TITLE,
                                 f"Could not write results for MaINbox:\n{exc}")
            return
        self.destroy()

    def _mb_cancel(self):
        if not self._mb_sent:
            try:
                with open(self._mb_json_out, "w", encoding="utf-8") as fh:
                    json.dump({"ok": False, "cancelled": True,
                               "merged_text": ""}, fh)
            except OSError:
                pass
        self.destroy()


# ---- CLI dispatch -----------------------------------------------------------
def _bridge_headless(argv):
    files, json_out = [], ""
    it = iter(argv)
    for a in it:
        if a == "--json-out":
            json_out = next(it, "")
        elif not a.startswith("--"):
            files.append(a)
    if not json_out:
        print("missing --json-out", file=sys.stderr)
        return 2
    cfg = _bridge_config()
    results = []
    try:
        for f in files:
            results.append((f, scan_path(f, cfg)))
        payload = _bridge_scan_result_payload(results)
    except Exception as exc:
        payload = {"ok": False, "error": f"{type(exc).__name__}: {exc}",
                   "rows": [], "logs": [], "warnings": [],
                   "engine_summary": "", "confidence": ""}
    with open(json_out, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=1)
    print(f"{BRIDGE_TITLE}: wrote {len(payload.get('rows', []))} row(s)")
    return 0


def _bridge_review(argv):
    files, json_out = [], ""
    it = iter(argv)
    for a in it:
        if a == "--json-out":
            json_out = next(it, "")
        elif not a.startswith("--"):
            files.append(a)
    if not json_out:
        print("missing --json-out", file=sys.stderr)
        return 2
    app = MainboxReviewGUI(files, json_out)
    app.mainloop()
    return 0


if __name__ == "__main__":
    _args = sys.argv[1:]
    if "--mainbox-headless" in _args:
        _args.remove("--mainbox-headless")
        sys.exit(_bridge_headless(_args))
    elif "--mainbox-review" in _args:
        _args.remove("--mainbox-review")
        sys.exit(_bridge_review(_args))
    elif "--selftest" in _args or "--deps" in _args or (
            _args and not _args[0].startswith("--") and "--json" in " ".join(_args)):
        sys.exit(main())
    elif _args and all(not a.startswith("--") for a in _args):
        sys.exit(main())
    else:
        _app = OmniScanGUI()
        _app.mainloop()
