# -*- coding: utf-8 -*-
"""
FastText - nucleo: modello del documento, formato .ftx, Word, conversioni Qt.

Formato .ftx: un file ZIP che contiene
    content.md      testo in Markdown semplice (tabelle con |, immagini con ![](...))
    images/...      le immagini incollate
    manifest.json   versione del formato

Apre anche .docx (Word) in forma semplificata: solo testo, tabelle e immagini,
sempre con un unico font. Esporta in .ftx, .docx, .txt e PDF.
"""

import base64
import hashlib
import io
import json
import os
import re
import zipfile

from PySide6.QtCore import QUrl, QBuffer, QIODevice
from PySide6.QtGui import (
    QColor, QFont, QTextFormat, QImage, QTextBlockFormat, QTextCharFormat, QTextCursor, QTextDocument, QTextFrameFormat,
    QTextImageFormat, QTextLength, QTextListFormat, QTextTable, QTextTableFormat,
)

APP_NAME = "FastText"
APP_VERSION = '1.41'
FORMAT_VERSION = 1
MAX_IMG_WIDTH = 700          # larghezza massima con cui una nuova immagine appare nel testo
OBJ = "￼"               # carattere segnaposto che Qt usa per le immagini
IMG_RE = re.compile(r"(?<!\\)!\[([^\]]*)\]\(([^)\s]+)\)")


# ---------------------------------------------------------------------------
#  Modello del documento
#  documento = lista di blocchi:
#     ("p", riga)                  paragrafo
#     ("h1"/"h2"/"h3", riga)       intestazione (titolo di sezione)
#     ("hr", [])                   linea di separazione
#     ("t", [[cella, ...], ...])   tabella; cella = lista di righe
#  riga = lista di segmenti: str oppure Img
# ---------------------------------------------------------------------------

class Img:
    __slots__ = ("name", "width")

    def __init__(self, name, width=None):
        self.name = name
        self.width = int(width) if width else None

    def __eq__(self, other):
        return isinstance(other, Img) and other.name == self.name and other.width == self.width

    def __repr__(self):
        return f"Img({self.name!r}, {self.width})"


class TableRows(list):
    """Righe di una tabella + larghezze delle colonne (in % della pagina) oppure None = automatiche."""

    def __init__(self, rows=(), widths=None):
        super().__init__(rows)
        self.widths = list(widths) if widths else None


def normalize_widths(widths, ncols):
    """Larghezze valide (interi, somma <= 100) oppure None."""
    if not widths or len(widths) != ncols:
        return None
    w = [max(0.0, float(x)) for x in widths]
    tot = sum(w)
    if tot <= 0:
        return None
    if tot > 100:
        w = [x * 100.0 / tot for x in w]
    return [max(2, int(round(x))) for x in w]


def line_is_empty(line):
    return all(isinstance(s, str) and s == "" for s in line)


def line_plain(line):
    return "".join(s for s in line if isinstance(s, str))


def model_has_rich(model):
    """True se il documento contiene tabelle o immagini."""
    for kind, data in model:
        if kind == "t":
            return True
        if any(isinstance(s, Img) for s in data):
            return True
    return False


def model_image_names(model):
    names = []
    def scan(line):
        for s in line:
            if isinstance(s, Img) and s.name not in names:
                names.append(s.name)
    for kind, data in model:
        if kind != "t":
            scan(data)
        else:
            for row in data:
                for cell in row:
                    for ln in cell:
                        scan(ln)
    return names


# ---------------------------------------------------------------------------
#  Archivio immagini (nome -> bytes). Il nome deriva dal contenuto: niente doppioni.
# ---------------------------------------------------------------------------

class ImageStore:
    def __init__(self):
        self.data = {}

    def clear(self):
        self.data.clear()

    def add_bytes(self, raw, hint_ext=None):
        """Aggiunge un'immagine; ritorna il nome oppure None se non leggibile."""
        img = QImage.fromData(raw)
        if img.isNull():
            raw = _convert_with_pillow(raw)
            if raw is None:
                return None
            img = QImage.fromData(raw)
            if img.isNull():
                return None
        ext = _sniff_ext(raw) or hint_ext or "png"
        if ext not in ("png", "jpg", "gif", "bmp", "webp"):
            raw = qimage_to_bytes(img)
            ext = "png"
        name = "images/" + hashlib.sha1(raw).hexdigest()[:12] + "." + ext
        self.data[name] = raw
        return name

    def add_qimage(self, img):
        if img is None or img.isNull():
            return None
        raw = qimage_to_bytes(img, "PNG")
        # foto grandi senza trasparenza: JPEG per non gonfiare il file
        if len(raw) > 2_500_000 and not img.hasAlphaChannel():
            raw = qimage_to_bytes(img, "JPG", 90)
        return self.add_bytes(raw)

    def qimage(self, name):
        raw = self.data.get(name)
        return QImage.fromData(raw) if raw else QImage()


def _sniff_ext(raw):
    if raw[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if raw[:3] == b"\xff\xd8\xff":
        return "jpg"
    if raw[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if raw[:2] == b"BM":
        return "bmp"
    if raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        return "webp"
    return None


def _convert_with_pillow(raw):
    """Prova a convertire formati che Qt non legge (es. EMF/WMF di Word su Windows)."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(raw))
        out = io.BytesIO()
        im.save(out, "PNG")
        return out.getvalue()
    except Exception:
        return None


def qimage_to_bytes(img, fmt="PNG", quality=-1):
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    img.save(buf, fmt, quality)
    return bytes(buf.data())


# ---------------------------------------------------------------------------
#  Markdown <-> modello
# ---------------------------------------------------------------------------

def _inline_to_md(line, in_cell=False):
    out = []
    for s in line:
        if isinstance(s, Img):
            alt = f"w={s.width}" if s.width else ""
            out.append(f"![{alt}]({s.name})")
        else:
            t = s.replace("![", "\\![")
            if in_cell:
                t = t.replace("|", "\\|")
            out.append(t)
    return "".join(out)


def model_to_md(model):
    lines = []
    items = list(model)
    # il paragrafo vuoto iniziale/finale attorno a una tabella è "strutturale" in Qt: non serve salvarlo
    if len(items) >= 2 and items[0][0] == "p" and line_is_empty(items[0][1]) and items[1][0] == "t":
        items = items[1:]
    if len(items) >= 2 and items[-1][0] == "p" and line_is_empty(items[-1][1]) and items[-2][0] == "t":
        items = items[:-1]
    for kind, data in items:
        if kind == "p":
            md = _inline_to_md(data)
            if md.startswith(("|", "\\", "#")) or _HR_RE.match(md):
                md = "\\" + md
            lines.append(md)
        elif kind in HEADING_KINDS:
            lines.append("#" * int(kind[1]) + " " + _inline_to_md(data))
        elif kind == "hr":
            lines.append("---")
        else:
            ncols = max(len(r) for r in data) if data else 1
            widths = normalize_widths(getattr(data, "widths", None), ncols)
            if widths:
                lines.append("<!-- ftx-larghezze: " + " ".join(str(w) for w in widths) + " -->")
            for ri, row in enumerate(data):
                cells = []
                for ci in range(ncols):
                    cell = row[ci] if ci < len(row) else [[""]]
                    cells.append("<br>".join(_inline_to_md(ln, True) for ln in cell))
                lines.append("| " + " | ".join(cells) + " |")
                if ri == 0:
                    lines.append("|" + "|".join([" --- "] * ncols) + "|")
    return "\n".join(lines) + "\n"


def _md_inline(text, in_cell=False):
    segs = []
    pos = 0
    for m in IMG_RE.finditer(text):
        if m.start() > pos:
            segs.append(text[pos:m.start()])
        alt, src = m.group(1), m.group(2)
        w = re.search(r"w=(\d+)", alt)
        segs.append(Img(src, int(w.group(1)) if w else None))
        pos = m.end()
    if pos < len(text):
        segs.append(text[pos:])
    out = []
    for s in segs:
        if isinstance(s, str):
            s = s.replace("\\![", "![")
            if in_cell:
                s = s.replace("\\|", "|")
        out.append(s)
    return out or [""]


def _split_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    return [c.strip() for c in re.split(r"(?<!\\)\|", s)]


_SEP_CELL = re.compile(r"^:?-{3,}:?$")
_WIDTHS_RE = re.compile(r"^<!--\s*ftx-larghezze:\s*([\d.\s]+?)\s*-->\s*$")
_HR_RE = re.compile(r"^\s*(?:-{3,}|\*{3,}|_{3,})\s*$")
_HEAD_RE = re.compile(r"^(#{1,3})(?: (.*))?$")
HEADING_KINDS = ("h1", "h2", "h3")


def md_to_model(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text.endswith("\n"):
        text = text[:-1]
    src = text.split("\n")
    model = []
    i = 0
    widths = None
    while i < len(src):
        ln = src[i]
        wm = _WIDTHS_RE.match(ln)
        if wm and i + 1 < len(src) and src[i + 1].startswith("|"):
            widths = [float(x) for x in wm.group(1).split()]
            i += 1
            continue
        if ln.startswith("|"):
            rows = []
            while i < len(src) and src[i].startswith("|"):
                cells = _split_row(src[i])
                if not (len(rows) == 1 and all(_SEP_CELL.match(c) for c in cells if c) and any(cells)):
                    rows.append([[_md_inline(part, True) for part in c.split("<br>")] for c in cells])
                i += 1
            ncols = max(len(r) for r in rows)
            for r in rows:
                while len(r) < ncols:
                    r.append([[""]])
            model.append(("t", TableRows(rows, normalize_widths(widths, ncols))))
            widths = None
            continue
        if ln[:1] == "\\" and ln[1:2] in ("|", "\\", "#", "-", "*", "_"):
            model.append(("p", _md_inline(ln[1:])))
            i += 1
            continue
        hm = _HEAD_RE.match(ln)
        if hm:
            model.append((f"h{len(hm.group(1))}", _md_inline(hm.group(2) or "")))
        elif _HR_RE.match(ln):
            model.append(("hr", []))
        else:
            model.append(("p", _md_inline(ln)))
        i += 1
    if not model:
        model.append(("p", [""]))
    return model


def model_to_plain(model):
    """Testo semplice (per .txt): tabelle in stile | a | b |, immagini omesse."""
    out = []
    for kind, data in model:
        if kind == "p":
            out.append(line_plain(data))
        elif kind in HEADING_KINDS:
            t = line_plain(data)
            out.append(t.upper() if kind == "h1" else t)
            if kind in ("h1", "h2"):
                out.append(("=" if kind == "h1" else "-") * max(3, len(t)))
        elif kind == "hr":
            out.append("-" * 40)
        else:
            for row in data:
                out.append("| " + " | ".join(" / ".join(line_plain(l) for l in cell) for cell in row) + " |")
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
#  File .ftx
# ---------------------------------------------------------------------------

def save_ftx(path, model, store):
    md = model_to_md(model)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as z:
        z.writestr("manifest.json", json.dumps({"format": "ftx", "version": FORMAT_VERSION,
                                                "app": f"{APP_NAME} {APP_VERSION}"}, indent=2),
                   compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("content.md", md.encode("utf-8"), compress_type=zipfile.ZIP_DEFLATED)
        for name in model_image_names(model):
            if name in store.data:
                z.writestr(name, store.data[name], compress_type=zipfile.ZIP_STORED)
    os.replace(tmp, path)


def load_ftx(path, store):
    with zipfile.ZipFile(path) as z:
        md = z.read("content.md").decode("utf-8")
        model = md_to_model(md)
        for name in model_image_names(model):
            try:
                raw = z.read(name)
            except KeyError:
                continue
            store.data[name] = raw
    return model


def load_md(path, store):
    with open(path, "rb") as f:
        md = _decode_text(f.read())
    model = md_to_model(md)
    base = os.path.dirname(os.path.abspath(path))
    rename = {}
    for name in model_image_names(model):
        p = name if os.path.isabs(name) else os.path.join(base, name)
        if os.path.isfile(p):
            with open(p, "rb") as f:
                new = store.add_bytes(f.read())
            if new:
                rename[name] = new
    _rename_images(model, rename)
    return model


def _rename_images(model, rename):
    def fix(line):
        for k, s in enumerate(line):
            if isinstance(s, Img):
                line[k] = Img(rename[s.name], s.width) if s.name in rename else s
    for kind, data in model:
        if kind != "t":
            fix(data)
        else:
            for row in data:
                for cell in row:
                    for ln in cell:
                        fix(ln)


def _decode_text(raw):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", "replace")


def load_txt(path):
    with open(path, "rb") as f:
        text = _decode_text(f.read())
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text.endswith("\n"):
        text = text[:-1]
    return [("p", [ln]) for ln in text.split("\n")]


# ---------------------------------------------------------------------------
#  Word (.docx): lettura semplificata e scrittura con un solo font
# ---------------------------------------------------------------------------

def load_docx(path, store):
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(path)
    part = doc.part
    T, TAB, BR, CR = qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")
    BLIP, IMGDATA, EXTENT = qn("a:blip"), "{urn:schemas-microsoft-com:vml}imagedata", qn("wp:extent")
    FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
    P, TBL, SDT, SDTC = qn("w:p"), qn("w:tbl"), qn("w:sdt"), qn("w:sdtContent")
    NUMPR, PPR, DELTEXT = qn("w:numPr"), qn("w:pPr"), qn("w:del")

    def image_from_rid(rid):
        try:
            blob = part.related_parts[rid].blob
        except Exception:
            return None
        return store.add_bytes(blob)

    def para_lines(p_el):
        lines = [[]]
        state = {"w": None}

        def add_text(t):
            cur = lines[-1]
            if cur and isinstance(cur[-1], str):
                cur[-1] += t
            else:
                cur.append(t)

        def walk(el):
            for ch in el:
                tag = ch.tag
                if tag == FALLBACK or tag == DELTEXT:
                    continue
                if tag == T:
                    add_text(ch.text or "")
                elif tag == TAB:
                    add_text("\t")
                elif tag in (BR, CR):
                    if ch.get(qn("w:type")) in (None, "textWrapping"):
                        lines.append([])
                elif tag == EXTENT:
                    try:
                        state["w"] = round(int(ch.get("cx")) / 9525)
                    except (TypeError, ValueError):
                        state["w"] = None
                elif tag == BLIP:
                    name = image_from_rid(ch.get(qn("r:embed")))
                    if name:
                        lines[-1].append(Img(name, state["w"]))
                    state["w"] = None
                elif tag == IMGDATA:
                    name = image_from_rid(ch.get(qn("r:id")))
                    if name:
                        lines[-1].append(Img(name))
                else:
                    walk(ch)

        walk(p_el)
        ppr = p_el.find(PPR)
        if ppr is not None and ppr.find(NUMPR) is not None:
            lines[0].insert(0, "• ")
        return [ln if ln else [""] for ln in lines]

    def table_rows(tbl_el):
        rows = []
        for tr in tbl_el.findall(qn("w:tr")):
            row = []
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                span, vcont = 1, False
                if tcpr is not None:
                    gs = tcpr.find(qn("w:gridSpan"))
                    if gs is not None:
                        try:
                            span = int(gs.get(qn("w:val")))
                        except (TypeError, ValueError):
                            span = 1
                    vm = tcpr.find(qn("w:vMerge"))
                    if vm is not None and vm.get(qn("w:val")) in (None, "continue"):
                        vcont = True
                cell = [[""]] if vcont else cell_lines(tc)
                row.append(cell)
                for _ in range(span - 1):
                    row.append([[""]])
            rows.append(row)
        ncols = max((len(r) for r in rows), default=1)
        for r in rows:
            while len(r) < ncols:
                r.append([[""]])
        widths = None
        grid = tbl_el.find(qn("w:tblGrid"))
        if grid is not None and text_width_twips:
            try:
                cols = [int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol"))]
                if len(cols) == ncols and all(c > 0 for c in cols):
                    widths = normalize_widths([c * 100.0 / text_width_twips for c in cols], ncols)
            except (TypeError, ValueError):
                widths = None
        return TableRows(rows or [[[[""]]]], widths)

    def cell_lines(tc):
        out = []
        for ch in tc:
            if ch.tag == P:
                out.extend(para_lines(ch))
            elif ch.tag == TBL:     # tabella annidata: appiattita in righe di testo
                for row in table_rows(ch):
                    out.append([" | ".join(line_plain(l) for c in row for l in c)])
            elif ch.tag == SDT:
                c = ch.find(SDTC)
                if c is not None:
                    out.extend(cell_lines(c))
        return out or [[""]]

    text_width_twips = 0
    try:
        sec = doc.sections[0]
        text_width_twips = (sec.page_width - sec.left_margin - sec.right_margin) / 635
    except Exception:
        pass
    model = []
    style_names = {}
    for st in doc.styles:
        try:
            style_names[st.style_id] = (st.name or "").lower()
        except Exception:
            pass

    def heading_level(p_el):
        ppr = p_el.find(PPR)
        if ppr is None:
            return 0
        ps = ppr.find(qn("w:pStyle"))
        if ps is not None:
            sid = ps.get(qn("w:val")) or ""
            name = style_names.get(sid, sid.lower())
            m = re.match(r"(?:heading|titolo)\s*(\d)", name) or re.match(r"(?:heading|titolo)\s*(\d)", sid.lower())
            if m:
                return min(3, max(1, int(m.group(1))))
            if name in ("title", "titolo"):
                return 1
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is not None:
            try:
                v = int(ol.get(qn("w:val")))
                if v < 9:
                    return min(3, v + 1)
            except (TypeError, ValueError):
                pass
        return 0

    def is_rule(p_el, lines):
        ppr = p_el.find(PPR)
        if ppr is None or not all(line_is_empty(l) for l in lines):
            return False
        bdr = ppr.find(qn("w:pBdr"))
        return bdr is not None and bdr.find(qn("w:bottom")) is not None

    def body(el):
        for ch in el:
            if ch.tag == P:
                lines = para_lines(ch)
                if is_rule(ch, lines):
                    model.append(("hr", []))
                    continue
                lvl = heading_level(ch)
                for ln in lines:
                    model.append((f"h{lvl}" if lvl else "p", ln))
            elif ch.tag == TBL:
                model.append(("t", table_rows(ch)))
            elif ch.tag == SDT:
                c = ch.find(SDTC)
                if c is not None:
                    body(c)

    body(doc.element.body)
    if not model:
        model.append(("p", [""]))
    return model


def save_docx(path, model, store, font):
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.oxml.ns import qn

    d = Document()
    st = d.styles["Normal"]
    st.font.name = font.family()
    st.font.size = Pt(font.pointSizeF() if font.pointSizeF() > 0 else 11)
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rfonts.set(qn(a), font.family())
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.line_spacing = 1.0
    base = font.pointSizeF() if font.pointSizeF() > 0 else 11
    from docx.shared import RGBColor
    for lvl, extra in ((1, 6), (2, 3), (3, 0)):
        hs = d.styles[f"Heading {lvl}"]
        hs.font.name = font.family()
        hs.font.size = Pt(base + extra)
        hs.font.bold = True
        hs.font.italic = False
        hs.font.color.rgb = RGBColor(0, 0, 0)
        hr_pr = hs.element.get_or_add_rPr()
        rf = hr_pr.find(qn("w:rFonts"))
        if rf is not None:
            for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if rf.get(qn(a)) is not None:
                    del rf.attrib[qn(a)]
            for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                rf.set(qn(a), font.family())
        hs.paragraph_format.space_before = Pt(10 if lvl == 1 else 6)
        hs.paragraph_format.space_after = Pt(3)

    def fill(par, line):
        for s in line:
            if isinstance(s, Img):
                raw = store.data.get(s.name)
                if not raw:
                    continue
                w = s.width or QImage.fromData(raw).width()
                w = min(w, 620)
                try:
                    par.add_run().add_picture(io.BytesIO(raw), width=Inches(w / 96))
                except Exception:
                    png = qimage_to_bytes(QImage.fromData(raw))
                    par.add_run().add_picture(io.BytesIO(png), width=Inches(w / 96))
            elif s:
                par.add_run(s)

    first = True
    for kind, data in model:
        if kind == "p":
            par = d.paragraphs[0] if first and d.paragraphs else d.add_paragraph()
            fill(par, data)
        elif kind in HEADING_KINDS:
            par = d.add_paragraph(style=f"Heading {kind[1]}")
            fill(par, data)
        elif kind == "hr":
            par = d.add_paragraph()
            from docx.oxml import OxmlElement
            ppr = par._p.get_or_add_pPr()
            bdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "808080")):
                bottom.set(qn(k), v)
            bdr.append(bottom)
            ppr.append(bdr)
        else:
            ncols = max(len(r) for r in data)
            t = d.add_table(rows=len(data), cols=ncols)
            t.style = "Table Grid"
            widths = normalize_widths(getattr(data, "widths", None), ncols)
            if widths:
                sec = d.sections[0]
                avail = sec.page_width - sec.left_margin - sec.right_margin
                t.autofit = False
                for ci, w in enumerate(widths):
                    emu = int(avail * w / 100)
                    t.columns[ci].width = emu
                    for cell in t.columns[ci].cells:
                        cell.width = emu
            for ri, row in enumerate(data):
                for ci, cell in enumerate(row):
                    c = t.cell(ri, ci)
                    for li, ln in enumerate(cell):
                        par = c.paragraphs[0] if li == 0 else c.add_paragraph()
                        fill(par, ln)
        first = False
    d.save(path)


# ---------------------------------------------------------------------------
#  QTextDocument <-> modello
# ---------------------------------------------------------------------------

def _block_lines(block, resolve):
    """Righe di un blocco Qt (le interruzioni morbide diventano righe nuove)."""
    lines = [[]]
    tl = block.textList()
    if tl is not None:
        style = tl.format().style()
        if style in (QTextListFormat.ListDecimal, QTextListFormat.ListLowerAlpha,
                     QTextListFormat.ListUpperAlpha, QTextListFormat.ListLowerRoman,
                     QTextListFormat.ListUpperRoman):
            lines[0].append(f"{tl.itemNumber(block) + 1}. ")
        else:
            lines[0].append("• ")
    it = block.begin()
    while not it.atEnd():
        frag = it.fragment()
        it += 1
        if not frag.isValid():
            continue
        fmt = frag.charFormat()
        if fmt.isImageFormat():
            imf = fmt.toImageFormat()
            name = resolve(imf.name())
            if name:
                w = int(imf.width()) if imf.width() > 0 else None
                for _ in range(frag.length()):
                    lines[-1].append(Img(name, w))
            continue
        parts = frag.text().replace(OBJ, "").split(" ")
        for k, part in enumerate(parts):
            if k > 0:
                lines.append([])
            if part:
                cur = lines[-1]
                if cur and isinstance(cur[-1], str):
                    cur[-1] += part
                else:
                    cur.append(part)
    return [ln if ln else [""] for ln in lines]


def _iter_lines(iterator, resolve):
    """Tutte le righe di un frame/cella; le tabelle annidate vengono appiattite."""
    out = []
    it = iterator
    while not it.atEnd():
        fr = it.currentFrame()
        if fr is not None:
            if isinstance(fr, QTextTable):
                for row in _table_rows(fr, resolve):
                    out.append([" | ".join(line_plain(l) for c in row for l in c)])
            else:
                out.extend(_iter_lines(fr.begin(), resolve))
        else:
            b = it.currentBlock()
            if b.isValid():
                out.extend(_block_lines(b, resolve))
        it += 1
    return out


def _table_rows(tbl, resolve):
    rows = []
    for r in range(tbl.rows()):
        row = []
        for c in range(tbl.columns()):
            cell = tbl.cellAt(r, c)
            if cell.row() != r or cell.column() != c:
                row.append([[""]])          # parte di una cella unita
            else:
                row.append(_iter_lines(cell.begin(), resolve) or [[""]])
        rows.append(row)
    return TableRows(rows, table_widths(tbl))


def doc_to_model(doc, resolve):
    model = []
    it = doc.rootFrame().begin()
    while not it.atEnd():
        fr = it.currentFrame()
        if fr is not None:
            if isinstance(fr, QTextTable):
                model.append(("t", _table_rows(fr, resolve)))
            else:
                for ln in _iter_lines(fr.begin(), resolve):
                    model.append(("p", ln))
        else:
            b = it.currentBlock()
            if b.isValid():
                bf = b.blockFormat()
                lines = _block_lines(b, resolve)
                if is_rule_format(bf):
                    for ln in lines:
                        if not line_is_empty(ln):
                            model.append(("p", ln))
                    model.append(("hr", []))
                else:
                    lvl = min(3, bf.headingLevel()) if bf.headingLevel() > 0 else 0
                    for ln in lines:
                        model.append((f"h{lvl}" if lvl else "p", ln))
        it += 1
    return model or [("p", [""])]


# ---- intestazioni e linee di separazione ---------------------------------

HEADING_SIZE = {1: 2, 2: 1, 3: 0}     # passi di ingrandimento rispetto al font del testo


def heading_block_format(level):
    bf = QTextBlockFormat()
    if level:
        bf.setHeadingLevel(level)
        bf.setTopMargin(10 if level == 1 else 6)
        bf.setBottomMargin(3)
    return bf


def heading_char_format(level):
    cf = QTextCharFormat()
    cf.setFontWeight(QFont.Bold if level else QFont.Normal)
    cf.setProperty(QTextFormat.FontSizeAdjustment, HEADING_SIZE.get(level, 0))
    return cf


def rule_block_format():
    bf = QTextBlockFormat()
    bf.setProperty(QTextFormat.BlockTrailingHorizontalRulerWidth, QTextLength(QTextLength.PercentageLength, 100))
    bf.setTopMargin(2)
    bf.setBottomMargin(2)
    return bf


def is_rule_format(bf):
    return bf.hasProperty(QTextFormat.BlockTrailingHorizontalRulerWidth)


def is_dark():
    """True se l'interfaccia usa un tema scuro (Windows in modalità scura o tema scelto)."""
    from PySide6.QtGui import QGuiApplication, QPalette
    return QGuiApplication.palette().color(QPalette.Base).lightness() < 128


def char_format_at(cursor):
    """Formato giusto per scrivere testo nuovo nel blocco del cursore."""
    lvl = cursor.blockFormat().headingLevel()
    return heading_char_format(lvl) if lvl > 0 else QTextCharFormat()


def text_format(cursor):
    """Formato del testo esistente sotto il cursore (mai un formato immagine)."""
    f = cursor.charFormat()
    if f.isImageFormat():
        return char_format_at(cursor)
    return f


def table_widths(tbl):
    """Larghezze delle colonne di una QTextTable in % della pagina, oppure None se automatiche."""
    tf = tbl.format()
    cons = tf.columnWidthConstraints()
    n = tbl.columns()
    if not cons or len(cons) != n:
        return None
    if any(c.type() == QTextLength.VariableLength for c in cons):
        return None
    vals = [c.rawValue() for c in cons]
    tot = sum(vals) or 1.0
    tw = tf.width()
    total = tw.rawValue() if tw.type() == QTextLength.PercentageLength else 100.0
    return normalize_widths([v / tot * total for v in vals], n)


def apply_table_widths(tf, widths, ncols):
    """Imposta sul formato tabella le larghezze (in % della pagina); None = predefinite."""
    widths = normalize_widths(widths, ncols)
    if not widths:
        tf.setWidth(QTextLength(QTextLength.PercentageLength, 100))
        tf.clearColumnWidthConstraints()
        return tf
    tot = sum(widths)
    tf.setWidth(QTextLength(QTextLength.PercentageLength, min(100, tot)))
    tf.setColumnWidthConstraints([QTextLength(QTextLength.PercentageLength, w * 100.0 / tot) for w in widths])
    return tf


def make_table_format():
    tf = QTextTableFormat()
    tf.setBorder(1)
    tf.setBorderStyle(QTextFrameFormat.BorderStyle_Solid)
    tf.setBorderBrush(QColor("#888888"))
    tf.setBorderCollapse(True)
    tf.setCellPadding(4)
    tf.setCellSpacing(0)
    tf.setMargin(4)
    tf.setWidth(QTextLength(QTextLength.PercentageLength, 100))
    return tf




def insert_model(cursor, model, insert_image):
    """Scrive il modello nel documento a partire dal cursore.
    insert_image(cursor, Img) inserisce un'immagine (o non fa nulla)."""
    def put_line(cur, line):
        fmt = char_format_at(cur)
        for s in line:
            if isinstance(s, Img):
                insert_image(cur, s)
            elif s:
                cur.insertText(s, fmt)

    def start_block(bf, cf):
        if need_block:
            cursor.insertBlock(bf, cf)
        else:
            cursor.setBlockFormat(bf)
            cursor.setBlockCharFormat(cf)

    need_block = False
    for kind, data in model:
        if kind == "p":
            if need_block:
                cursor.insertBlock(QTextBlockFormat(), QTextCharFormat())
            put_line(cursor, data)
            need_block = True
        elif kind in HEADING_KINDS:
            lvl = int(kind[1])
            start_block(heading_block_format(lvl), heading_char_format(lvl))
            put_line(cursor, data)
            need_block = True
        elif kind == "hr":
            # Qt disegna male una linea messa subito dopo una tabella: lascia la riga vuota in mezzo
            if not need_block and QTextCursor(cursor.block().previous()).currentTable() is not None:
                need_block = True
            start_block(rule_block_format(), QTextCharFormat())
            need_block = True
        else:
            rows = len(data)
            cols = max(len(r) for r in data) if data else 1
            tf = apply_table_widths(make_table_format(), getattr(data, "widths", None), max(1, cols))
            tbl = cursor.insertTable(max(1, rows), max(1, cols), tf)
            for ri, row in enumerate(data):
                for ci, cell in enumerate(row):
                    cc = tbl.cellAt(ri, ci).firstCursorPosition()
                    for li, ln in enumerate(cell):
                        if li:
                            cc.insertBlock(QTextBlockFormat(), QTextCharFormat())
                        put_line(cc, ln)
            cursor.setPosition(tbl.lastPosition() + 1)
            need_block = False
    if model and model[-1][0] == "hr":        # dopo una linea si continua a scrivere in un blocco normale
        cursor.insertBlock(QTextBlockFormat(), QTextCharFormat())
    return cursor


def build_headless(model, store):
    """Documento Qt senza finestra (per cercare/sostituire nei file non aperti).
    Le immagini restano come riferimenti: posizioni identiche all'editor."""
    doc = QTextDocument()

    def put_img(cur, img):
        if img.name in store.data:
            f = QTextImageFormat()
            f.setName(img.name)
            if img.width:
                f.setWidth(img.width)
            cur.insertImage(f)

    insert_model(QTextCursor(doc), model, put_img)
    return doc


def load_any(path, store):
    """Apre un file supportato; ritorna (modello, tipo)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return load_docx(path, store), "docx"
    if ext == ".md":
        return load_md(path, store), "md"
    if ext == ".ftx" or (ext not in (".txt",) and zipfile.is_zipfile(path)):
        return load_ftx(path, store), "ftx"
    return load_txt(path), "txt"

