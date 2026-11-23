# -*- coding: utf-8 -*-
"""
FastText - un Blocco note con tabelle e immagini, a schede (tab) come Notepad++.

Formato .ftx: un file ZIP che contiene
    content.md      testo in Markdown semplice (tabelle con |, immagini con ![](...))
    images/...      le immagini incollate
    manifest.json   versione del formato

Apre anche .docx (Word) in forma semplificata, .txt e .md.
"""

import base64
import json
import os
import subprocess
import sys
import tempfile
import urllib.request

from PySide6.QtCore import Qt, QUrl, QSettings, QDateTime, QMimeData, Signal, QSize, QRectF, QTimer, QEvent, QByteArray
from PySide6.QtGui import (
    QTextBlockFormat, QAction, QColor, QFont, QFontDatabase, QIcon, QImage, QKeySequence, QPainter, QPen, QPixmap,
    QTextCharFormat, QTextCursor, QTextDocument, QTextFormat, QTextImageFormat, QDesktopServices, QTextOption,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QApplication, QDialog, QDialogButtonBox, QFileDialog, QFontDialog, QFormLayout, QLabel, QMainWindow,
    QMenu, QMessageBox, QSpinBox, QTabWidget, QTextEdit, QToolBar, QVBoxLayout, QWidget, QStyle, QTabBar,
    QToolButton,
)

from ftcore import (
    APP_NAME, APP_VERSION, MAX_IMG_WIDTH, OBJ, Img, ImageStore, doc_to_model, insert_model, line_is_empty,
    line_plain, load_any, make_table_format, model_has_rich, model_image_names, model_to_plain, qimage_to_bytes,
    save_docx, save_ftx, char_format_at, heading_block_format, heading_char_format, rule_block_format,
    is_rule_format, HEADING_KINDS,
)
from ftpanels import SidePanel
import fttextops as ops
from ftsearch import ResultsDock, SearchController, SearchDialog, compile_pattern, find_matches, Opts

IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")
DOC_EXTS = (".ftx", ".docx", ".txt", ".md")
SERVER_NAME = "FastText-" + (os.environ.get("USERNAME") or os.environ.get("USER") or "user")


# ---------------------------------------------------------------------------
#  Editor
# ---------------------------------------------------------------------------

class Editor(QTextEdit):
    openFileRequested = Signal(str)

    def __init__(self, store, parent=None):
        super().__init__(parent)
        self.store = store
        self.setAcceptRichText(True)
        self.setTabChangesFocus(False)
        self.document().setDocumentMargin(10)
        self.reserved = set()            # scorciatoie della finestra che l'editor non deve "mangiare"
        self.marks = []
        self.smart = []
        self.highlight_line = True
        self._smart_timer = QTimer(self)
        self._smart_timer.setSingleShot(True)
        self._smart_timer.setInterval(150)
        self._smart_timer.timeout.connect(self._update_smart)
        self.selectionChanged.connect(self._smart_timer.start)
        self.cursorPositionChanged.connect(self.refresh_extra)

    # ---- evidenziazioni -----------------------------------------------
    def _sel(self, s, e, color, full=False):
        sel = QTextEdit.ExtraSelection()
        c = QTextCursor(self.document())
        c.setPosition(s)
        c.setPosition(e, QTextCursor.KeepAnchor)
        sel.cursor = c
        sel.format.setBackground(QColor(color))
        if full:
            sel.format.setProperty(QTextFormat.FullWidthSelection, True)
        return sel

    def refresh_extra(self):
        sels = []
        if self.highlight_line and not self.textCursor().hasSelection():
            c = self.textCursor()
            line = QTextEdit.ExtraSelection()
            line.format.setBackground(QColor("#eef4fc"))
            line.format.setProperty(QTextFormat.FullWidthSelection, True)
            c.clearSelection()
            line.cursor = c
            sels.append(line)
        self.setExtraSelections(sels + self.smart + self.marks)

    def set_marks(self, spans):
        self.marks = [self._sel(s, e, "#ffe36e") for s, e in spans]
        self.refresh_extra()

    def clear_marks(self):
        self.marks = []
        self.refresh_extra()

    def _update_smart(self):
        """Doppio clic su una parola: evidenzia tutte le altre uguali (come Notepad++)."""
        self.smart = []
        c = self.textCursor()
        word = c.selectedText()
        if word and len(word) < 100 and word.replace("_", "a").isalnum():
            pat = compile_pattern(Opts(word, case=True, word=True))
            ms = find_matches(self.document(), pat)
            if len(ms) > 1:
                self.smart = [self._sel(m.start(), m.end(), "#c8f0c0") for m in ms[:2000]
                              if not (m.start() == c.selectionStart() and m.end() == c.selectionEnd())]
        self.refresh_extra()

    # ---- immagini -----------------------------------------------------
    def register_image(self, name):
        img = self.store.qimage(name)
        if not img.isNull():
            self.document().addResource(QTextDocument.ImageResource, QUrl(name), img)
        return img

    def image_format(self, name, width=None):
        img = self.register_image(name)
        f = QTextImageFormat()
        f.setName(name)
        if not img.isNull():
            w = width or min(img.width(), MAX_IMG_WIDTH)
            f.setWidth(w)
            f.setHeight(w * img.height() / max(1, img.width()))
        elif width:
            f.setWidth(width)
        return f

    def insert_image_name(self, cursor, name, width=None):
        cursor.insertImage(self.image_format(name, width))

    def _put_img(self, cursor, img):
        if img.name in self.store.data:
            self.insert_image_name(cursor, img.name, img.width)

    # ---- documento <-> modello ----------------------------------------
    def insert_model(self, cursor, model):
        return insert_model(cursor, model, self._put_img)

    def set_model(self, model):
        doc = self.document()
        doc.setUndoRedoEnabled(False)
        self.clear()
        self.insert_model(QTextCursor(doc), model)
        doc.setUndoRedoEnabled(True)
        doc.clearUndoRedoStacks()
        self.setTextCursor(QTextCursor(doc))
        doc.setModified(False)
        self.marks, self.smart = [], []
        self.refresh_extra()

    def get_model(self):
        return doc_to_model(self.document(), lambda n: n if n in self.store.data else None)

    # ---- incolla / trascina -------------------------------------------
    def canInsertFromMimeData(self, source):
        return source.hasImage() or source.hasUrls() or super().canInsertFromMimeData(source)

    def _resolve_external(self, name, src_doc):
        if not name:
            return None
        if name in self.store.data:
            return name
        raw = None
        try:
            if name.startswith("data:"):
                head, _, payload = name.partition(",")
                raw = base64.b64decode(payload) if ";base64" in head else urllib.request.unquote(payload).encode()
            else:
                url = QUrl(name)
                local = url.toLocalFile() if url.isLocalFile() else (name if os.path.isfile(name) else "")
                if local and os.path.isfile(local):
                    with open(local, "rb") as f:
                        raw = f.read()
                elif url.scheme() in ("http", "https"):
                    req = urllib.request.Request(name, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=4) as r:
                        raw = r.read()
        except Exception:
            raw = None
        if raw:
            return self.store.add_bytes(raw)
        if src_doc is not None:
            res = src_doc.resource(QTextDocument.ImageResource, QUrl(name))
            if isinstance(res, QImage) and not res.isNull():
                return self.store.add_qimage(res)
            if isinstance(res, QPixmap) and not res.isNull():
                return self.store.add_qimage(res.toImage())
        return None

    def html_to_model(self, html):
        tmp = QTextDocument()
        tmp.setHtml(html)
        model = doc_to_model(tmp, lambda n: self._resolve_external(n, tmp))
        out = []
        for k, item in enumerate(model):
            if (k > 0 and model[k - 1][0] == "t" and item[0] == "p" and line_is_empty(item[1])
                    and k + 1 < len(model) and model[k + 1][0] == "p"):
                continue
            out.append(item)
        return out

    def insertFromMimeData(self, source):
        cur = self.textCursor()
        if source.hasUrls():
            files = [u.toLocalFile() for u in source.urls() if u.isLocalFile()]
            docs = [f for f in files if f.lower().endswith(DOC_EXTS)]
            imgs = [f for f in files if f.lower().endswith(IMAGE_EXTS)]
            if docs and not imgs:
                for d in docs:
                    self.openFileRequested.emit(d)
                return
            if imgs:
                cur.beginEditBlock()
                for f in imgs:
                    with open(f, "rb") as fh:
                        name = self.store.add_bytes(fh.read())
                    if name:
                        self.insert_image_name(cur, name)
                cur.endEditBlock()
                self.setTextCursor(cur)
                return
        if source.hasHtml():
            model = self.html_to_model(source.html())
            has_text = any(kind == "t" or line_plain(d).strip() for kind, d in model)
            if model_has_rich(model) or has_text or not source.hasImage():
                self.paste_model(model)
                return
        if source.hasImage():
            img = source.imageData()
            if isinstance(img, QPixmap):
                img = img.toImage()
            name = self.store.add_qimage(img)
            if name:
                cur.insertImage(self.image_format(name))
                self.setTextCursor(cur)
                return
        if source.hasText():
            cur.insertText(source.text().replace("\r\n", "\n"), char_format_at(cur))
            self.setTextCursor(cur)

    def paste_model(self, model):
        while len(model) > 1 and model[0][0] == "p" and line_is_empty(model[0][1]):
            model = model[1:]
        while len(model) > 1 and model[-1][0] == "p" and line_is_empty(model[-1][1]):
            model = model[:-1]
        cur = self.textCursor()
        cur.beginEditBlock()
        cur.removeSelectedText()
        if model and model[0][0] in HEADING_KINDS + ("hr",) and cur.positionInBlock() > 0 and not cur.currentTable():
            cur.insertBlock(QTextBlockFormat(), QTextCharFormat())
        if cur.currentTable() is not None:     # niente intestazioni/linee dentro le celle
            model = [("p", d) if k in HEADING_KINDS else (k, d) for k, d in model if k != "hr"]
        self.insert_model(cur, model)
        cur.endEditBlock()
        self.setTextCursor(cur)

    def paste_plain(self):
        md = QApplication.clipboard().mimeData()
        if md and md.hasText():
            c = self.textCursor()
            c.insertText(md.text().replace("\r\n", "\n"), char_format_at(c))

    def createMimeDataFromSelection(self):
        md = super().createMimeDataFromSelection()
        out = QMimeData()
        cur = self.textCursor()
        frag = cur.selection()
        html = frag.toHtml()
        for name, raw in self.store.data.items():
            if name in html:
                mime = "image/" + ("jpeg" if name.endswith(".jpg") else name.rsplit(".", 1)[-1])
                html = html.replace(f'src="{name}"', f'src="data:{mime};base64,{base64.b64encode(raw).decode()}"')
        out.setHtml(html)
        out.setText(frag.toPlainText().replace(OBJ, "").replace(" ", "\n").replace(" ", "\n"))
        for fmt in md.formats():
            if fmt not in ("text/html", "text/plain", "text/markdown"):
                out.setData(fmt, md.data(fmt))
        if cur.selectedText() == OBJ and cur.charFormat().isImageFormat():
            out.setImageData(self.store.qimage(cur.charFormat().toImageFormat().name()))
        return out

    # ---- tastiera -----------------------------------------------------
    def event(self, e):
        if e.type() == QEvent.ShortcutOverride and self.reserved:
            seq = QKeySequence(e.keyCombination()).toString()
            if seq in self.reserved:
                e.ignore()
                return True
        return super().event(e)

    def keyPressEvent(self, e):
        cur = self.textCursor()
        tbl = cur.currentTable()
        if tbl is not None and e.key() in (Qt.Key_Tab, Qt.Key_Backtab) and not (e.modifiers() & Qt.ControlModifier):
            cell = tbl.cellAt(cur)
            idx = cell.row() * tbl.columns() + cell.column()
            idx += -1 if e.key() == Qt.Key_Backtab else 1
            if idx < 0:
                return
            if idx >= tbl.rows() * tbl.columns():
                tbl.appendRows(1)
            nc = tbl.cellAt(idx // tbl.columns(), idx % tbl.columns())
            self.setTextCursor(nc.firstCursorPosition())
            return
        bf = cur.blockFormat()
        if (bf.headingLevel() > 0 or is_rule_format(bf)) and not cur.hasSelection():
            if e.key() in (Qt.Key_Return, Qt.Key_Enter) and cur.atBlockEnd() and not (e.modifiers() & Qt.ShiftModifier):
                # dopo un titolo (o una linea) si torna al testo normale
                cur.insertBlock(QTextBlockFormat(), QTextCharFormat())
                self.setTextCursor(cur)
                self.ensureCursorVisible()
                return
            if is_rule_format(bf) and e.text() and e.text().isprintable() and not (e.modifiers() & Qt.ControlModifier):
                cur.movePosition(QTextCursor.EndOfBlock)
                cur.insertBlock(QTextBlockFormat(), QTextCharFormat())
                self.setTextCursor(cur)
        if e.key() == Qt.Key_Escape and (self.marks or self.smart):
            self.marks, self.smart = [], []
            self.refresh_extra()
            return
        super().keyPressEvent(e)

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier:
            win = self.window()
            if hasattr(win, "zoom"):
                win.zoom(1 if e.angleDelta().y() > 0 else -1)
            return
        super().wheelEvent(e)

    # ---- immagini sotto il mouse --------------------------------------
    def image_cursor_at(self, pos):
        base = self.cursorForPosition(pos).position()
        for start in (base, base - 1):
            if start < 0:
                continue
            c = QTextCursor(self.document())
            c.setPosition(start)
            if not c.movePosition(QTextCursor.NextCharacter, QTextCursor.KeepAnchor):
                continue
            if c.charFormat().isImageFormat():
                return c
        return None

    def mouseDoubleClickEvent(self, e):
        c = self.image_cursor_at(e.position().toPoint())
        if c is not None:
            self.open_image_external(c)
            return
        super().mouseDoubleClickEvent(e)

    def open_image_external(self, c):
        name = c.charFormat().toImageFormat().name()
        raw = self.store.data.get(name)
        if not raw:
            return
        path = os.path.join(tempfile.gettempdir(), "FastText_" + os.path.basename(name))
        with open(path, "wb") as f:
            f.write(raw)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def resize_image(self, c, factor):
        f = c.charFormat().toImageFormat()
        img = self.store.qimage(f.name())
        if img.isNull():
            return
        w = max(16, int(img.width() * factor)) if factor else min(img.width(), MAX_IMG_WIDTH)
        f.setWidth(w)
        f.setHeight(w * img.height() / max(1, img.width()))
        c.setCharFormat(f)

    def save_image_as(self, c):
        name = c.charFormat().toImageFormat().name()
        raw = self.store.data.get(name)
        if not raw:
            return
        ext = name.rsplit(".", 1)[-1]
        path, _ = QFileDialog.getSaveFileName(self, "Salva immagine", "immagine." + ext, f"Immagine (*.{ext})")
        if path:
            with open(path, "wb") as fh:
                fh.write(raw)

    def contextMenuEvent(self, e):
        menu = self.createStandardContextMenu(e.pos())
        win = self.window()
        menu.addSeparator()
        menu.addAction("Incolla come testo semplice", self.paste_plain)
        if hasattr(win, "case_menu"):
            menu.addMenu(win.case_menu)
            menu.addMenu(win.lines_menu)
        ic = self.image_cursor_at(e.pos())
        if ic is not None:
            menu.addSeparator()
            sub = menu.addMenu("Immagine")
            for label, fac in (("25%", 0.25), ("50%", 0.5), ("75%", 0.75), ("100% (originale)", 1.0),
                               ("Adatta alla pagina", None)):
                sub.addAction(f"Dimensione {label}", lambda fac=fac, ic=ic: self.resize_image(ic, fac))
            sub.addSeparator()
            sub.addAction("Apri con il visualizzatore", lambda ic=ic: self.open_image_external(ic))
            sub.addAction("Salva immagine come...", lambda ic=ic: self.save_image_as(ic))
        clicked = self.cursorForPosition(e.pos())
        if clicked.currentTable() is not None and not self.textCursor().currentTable():
            self.setTextCursor(clicked)
        if self.textCursor().currentTable() is not None and hasattr(win, "table_menu"):
            menu.addSeparator()
            menu.addMenu(win.table_menu)
        menu.exec(e.globalPos())


# ---------------------------------------------------------------------------
#  Una scheda = un documento
# ---------------------------------------------------------------------------

class DocTab(QWidget):
    _untitled = 0

    def __init__(self, win):
        super().__init__()
        self.store = ImageStore()
        self.editor = Editor(self.store)
        self.path = None
        self.kind = None
        DocTab._untitled += 1
        self.untitled_name = f"Nuovo {DocTab._untitled}"
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.editor)
        self.editor.set_model([("p", [""])])

    def title(self):
        return os.path.basename(self.path) if self.path else self.untitled_name

    def modified(self):
        return self.editor.document().isModified()

    def pristine(self):
        return not self.path and not self.modified() and self.editor.document().isEmpty()


# ---------------------------------------------------------------------------
#  Dialoghi
# ---------------------------------------------------------------------------

class TableDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Inserisci tabella")
        lay = QFormLayout(self)
        self.rows = QSpinBox(); self.rows.setRange(1, 200); self.rows.setValue(3)
        self.cols = QSpinBox(); self.cols.setRange(1, 30); self.cols.setValue(3)
        lay.addRow("Righe:", self.rows)
        lay.addRow("Colonne:", self.cols)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addRow(bb)


def make_icon():
    icon = QIcon()
    for size in (16, 24, 32, 48, 64, 128, 256):
        pm = QPixmap(size, size)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        s = size / 64.0
        p.setPen(QPen(QColor("#3a4a5c"), max(1.0, 2 * s)))
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(QRectF(10 * s, 4 * s, 44 * s, 56 * s), 5 * s, 5 * s)
        p.setPen(QPen(QColor("#7a8794"), max(1.0, 2 * s)))
        for y in (14, 20):
            p.drawLine(int(17 * s), int(y * s), int(47 * s), int(y * s))
        p.setPen(QPen(QColor("#2f7dd1"), max(1.0, 1.6 * s)))
        p.setBrush(Qt.NoBrush)
        p.drawRect(QRectF(17 * s, 27 * s, 30 * s, 12 * s))
        p.drawLine(int(32 * s), int(27 * s), int(32 * s), int(39 * s))
        p.drawLine(int(17 * s), int(33 * s), int(47 * s), int(33 * s))
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#45a35a"))
        p.drawRect(QRectF(17 * s, 43 * s, 30 * s, 11 * s))
        p.setBrush(QColor("#f2c14e"))
        p.drawEllipse(QRectF(20 * s, 45 * s, 5 * s, 5 * s))
        p.end()
        icon.addPixmap(pm)
    return icon


# ---------------------------------------------------------------------------
#  Finestra principale
# ---------------------------------------------------------------------------

FILTER_OPEN = ("Tutti i documenti (*.ftx *.docx *.txt *.md);;FastText (*.ftx);;Word (*.docx);;"
               "Testo (*.txt);;Markdown (*.md);;Tutti i file (*)")
FILTER_SAVE = "FastText (*.ftx);;Word semplificato (*.docx);;Testo semplice (*.txt)"
KIND_LABEL = {"ftx": "FastText", "docx": "Word (semplificato)", "txt": "Testo", "md": "Markdown", None: "Nuovo"}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("FastText", "FastText")
        self.zoom_delta = 0
        self.closed_paths = []
        self.last_opts = None
        self.setWindowIcon(make_icon())

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.setElideMode(Qt.ElideMiddle)
        self.tabs.setUsesScrollButtons(True)
        self.tabs.tabCloseRequested.connect(lambda i: self.close_tab(self.tabs.widget(i)))
        self.tabs.currentChanged.connect(self.on_tab_changed)
        bar = self.tabs.tabBar()
        bar.setContextMenuPolicy(Qt.CustomContextMenu)
        bar.customContextMenuRequested.connect(self.tab_context_menu)
        bar.installEventFilter(self)
        bar.tabMoved.connect(lambda *_: self.side.refresh_files())
        self.setCentralWidget(self.tabs)

        self.side = SidePanel(self)
        self.addDockWidget(Qt.LeftDockWidgetArea, self.side)
        self.results = ResultsDock(self)
        self.results.hitActivated.connect(self.goto_hit)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.results)
        self.results.hide()
        self.search = SearchController(self)
        self.search_dlg = None

        f = QFont()
        saved = self.settings.value("font", "")
        if not (saved and f.fromString(saved)):
            f = QFont("Consolas", 11)
            if not QFontDatabase.hasFamily("Consolas"):
                f = QFontDatabase.systemFont(QFontDatabase.FixedFont)
                f.setPointSize(11)
        self.base_font = f
        self.wrap = self.settings.value("wrap", True, type=bool)
        self.show_ws = self.settings.value("show_ws", False, type=bool)
        self.hl_line = self.settings.value("hl_line", True, type=bool)

        self.status = QLabel()
        self.status_kind = QLabel()
        self.statusBar().addPermanentWidget(self.status_kind)
        self.statusBar().addPermanentWidget(self.status)
        self.build_actions()
        self.resize(self.settings.value("size", QSize(1100, 760)))
        geo = self.settings.value("geometry")
        if isinstance(geo, QByteArray):
            self.restoreGeometry(geo)
        st = self.settings.value("winstate")
        if isinstance(st, QByteArray):
            self.restoreState(st)
            self.results.hide()

    # ---- schede -------------------------------------------------------
    @property
    def current(self):
        t = self.tabs.currentWidget()
        if t is None:
            t = self.new_tab()
        return t

    @property
    def editor(self):
        return self.current.editor

    def all_tabs(self):
        return [self.tabs.widget(i) for i in range(self.tabs.count())]

    def tab_for_path(self, path):
        norm = os.path.normcase(os.path.abspath(path))
        for t in self.all_tabs():
            if t.path and os.path.normcase(os.path.abspath(t.path)) == norm:
                return t
        return None

    def new_tab(self):
        t = DocTab(self)
        ed = t.editor
        ed.openFileRequested.connect(self.open_path)
        ed.document().modificationChanged.connect(lambda _m, t=t: self.update_tab(t))
        ed.cursorPositionChanged.connect(self.update_status)
        ed.selectionChanged.connect(self.update_status)
        ed.document().contentsChanged.connect(lambda t=t: self.side.schedule_outline() if t is self.tabs.currentWidget() else None)
        ed.reserved = self.reserved_keys
        self.style_editor(ed)
        i = self.tabs.addTab(t, t.title())
        self.tabs.setCurrentIndex(i)
        self.update_tab(t)
        ed.setFocus()
        return t

    def update_tab(self, t):
        i = self.tabs.indexOf(t)
        if i < 0:
            return
        self.tabs.setTabText(i, ("● " if t.modified() else "") + t.title())
        self.tabs.setTabToolTip(i, t.path or t.title())
        self.tabs.setTabIcon(i, QIcon())
        if t is self.tabs.currentWidget():
            self.update_title()
        self.side.refresh_files()

    def on_tab_changed(self, _i):
        if self.tabs.count() == 0:
            return
        self.update_title()
        self.update_status()
        self.side.refresh_files()
        self.side.refresh_outline()
        self.editor.setFocus()

    def close_tab(self, t, ask=True):
        if t is None:
            return False
        if ask and not self.maybe_save(t):
            return False
        if t.path:
            self.closed_paths.append(t.path)
        self.tabs.removeTab(self.tabs.indexOf(t))
        t.deleteLater()
        if self.tabs.count() == 0:
            self.new_tab()
        self.side.refresh_files()
        return True

    def close_others(self, keep, only_right=False):
        tabs = self.all_tabs()
        start = tabs.index(keep) + 1 if only_right else 0
        for t in tabs[start:]:
            if t is not keep:
                if not self.close_tab(t):
                    return

    def close_all(self):
        for t in self.all_tabs():
            if not self.close_tab(t):
                return

    def reopen_closed(self):
        while self.closed_paths:
            p = self.closed_paths.pop()
            if os.path.isfile(p) and not self.tab_for_path(p):
                self.open_path(p)
                return

    def cycle_tab(self, step):
        n = self.tabs.count()
        if n > 1:
            self.tabs.setCurrentIndex((self.tabs.currentIndex() + step) % n)

    def tab_context_menu(self, pos):
        bar = self.tabs.tabBar()
        i = bar.tabAt(pos)
        if i >= 0:
            self.show_tab_menu(self.tabs.widget(i), bar.mapToGlobal(pos))

    def show_tab_menu(self, t, global_pos):
        m = QMenu(self)
        m.addAction("Chiudi", lambda: self.close_tab(t))
        m.addAction("Chiudi tutti gli altri", lambda: self.close_others(t))
        m.addAction("Chiudi tutti a destra", lambda: self.close_others(t, True))
        m.addAction("Chiudi tutti", self.close_all)
        m.addSeparator()
        a = m.addAction("Salva", lambda: self.save(t))
        m.addAction("Salva con nome...", lambda: self.save_as(t))
        m.addSeparator()
        for act in (m.addAction("Copia percorso completo", lambda: QApplication.clipboard().setText(t.path or "")),
                    m.addAction("Copia nome file", lambda: QApplication.clipboard().setText(t.title())),
                    m.addAction("Apri cartella contenente", lambda: self.show_in_folder(t.path)),
                    m.addAction("Ricarica dal disco", lambda: self.reload_tab(t))):
            act.setEnabled(bool(t.path))
        m.exec(global_pos)

    def eventFilter(self, obj, e):
        if obj is self.tabs.tabBar():
            if e.type() == QEvent.MouseButtonRelease and e.button() == Qt.MiddleButton:
                i = obj.tabAt(e.position().toPoint())
                if i >= 0:
                    self.close_tab(self.tabs.widget(i))
                    return True
            if e.type() == QEvent.MouseButtonDblClick and obj.tabAt(e.position().toPoint()) < 0:
                self.new_tab()
                return True
        return super().eventFilter(obj, e)

    def show_in_folder(self, path):
        if not path:
            return
        if sys.platform == "win32":
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        else:
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))

    # ---- azioni / menu ------------------------------------------------
    def act(self, menu, text, slot, shortcut=None, icon=None, widget_only=False):
        a = QAction(text, self)
        if shortcut:
            a.setShortcuts([QKeySequence(s) for s in shortcut] if isinstance(shortcut, list) else [QKeySequence(shortcut)])
        if icon is not None:
            a.setIcon(self.style().standardIcon(icon))
        a.triggered.connect(lambda _=False: slot())
        if widget_only:
            a.setShortcutContext(Qt.WidgetShortcut)
        menu.addAction(a)
        self.addAction(a)
        return a

    def build_actions(self):
        mb = self.menuBar()
        S = QStyle.StandardPixmap
        ed = lambda: self.editor

        m = mb.addMenu("&File")
        a_new = self.act(m, "Nuovo", self.new_tab, QKeySequence.New, S.SP_FileIcon)
        a_open = self.act(m, "Apri...", self.open_dialog, QKeySequence.Open, S.SP_DialogOpenButton)
        self.recent_menu = m.addMenu("File recenti")
        self.recent_menu.aboutToShow.connect(self.fill_recent)
        self.act(m, "Riapri la scheda chiusa", self.reopen_closed, "Ctrl+Shift+T")
        self.act(m, "Ricarica dal disco", lambda: self.reload_tab(self.current))
        m.addSeparator()
        a_save = self.act(m, "Salva", lambda: self.save(), QKeySequence.Save, S.SP_DialogSaveButton)
        self.act(m, "Salva con nome...", lambda: self.save_as(), "Ctrl+Alt+S")
        a_saveall = self.act(m, "Salva tutto", self.save_all, "Ctrl+Shift+S")
        m.addSeparator()
        self.act(m, "Chiudi scheda", lambda: self.close_tab(self.current), ["Ctrl+W", "Ctrl+F4"])
        self.act(m, "Chiudi tutte le schede", self.close_all, "Ctrl+Shift+W")
        m.addSeparator()
        self.act(m, "Esporta in Word (.docx)...", lambda: self.export("docx"))
        self.act(m, "Esporta in PDF...", lambda: self.export("pdf"))
        self.act(m, "Stampa...", self.print_doc, QKeySequence.Print)
        m.addSeparator()
        self.act(m, "Esci", self.close, "Alt+F4")

        m = mb.addMenu("&Modifica")
        self.act(m, "Annulla", lambda: ed().undo(), QKeySequence.Undo, widget_only=True)
        self.act(m, "Ripeti", lambda: ed().redo(), QKeySequence.Redo, widget_only=True)
        m.addSeparator()
        self.act(m, "Taglia", lambda: ed().cut(), QKeySequence.Cut, widget_only=True)
        self.act(m, "Copia", lambda: ed().copy(), QKeySequence.Copy, widget_only=True)
        self.act(m, "Incolla", lambda: ed().paste(), QKeySequence.Paste, widget_only=True)
        self.act(m, "Incolla come testo semplice", lambda: ed().paste_plain(), "Ctrl+Shift+V")
        self.act(m, "Seleziona tutto", lambda: ed().selectAll(), QKeySequence.SelectAll, widget_only=True)
        m.addSeparator()
        self.case_menu = m.addMenu("Maiuscole / minuscole")
        cm = self.case_menu
        self.act(cm, "MAIUSCOLO", lambda: self.case("upper"), "Ctrl+Shift+U")
        self.act(cm, "minuscolo", lambda: self.case("lower"), "Ctrl+U")
        self.act(cm, "Iniziali Maiuscole", lambda: self.case("proper"), "Alt+U")
        self.act(cm, "Iniziali Maiuscole (mantieni il resto)", lambda: self.case("proper_blend"), "Alt+Shift+U")
        self.act(cm, "Prima lettera della frase maiuscola", lambda: self.case("sentence"), "Ctrl+Alt+U")
        self.act(cm, "Prima lettera della frase (mantieni il resto)", lambda: self.case("sentence_blend"), "Ctrl+Alt+Shift+U")
        self.act(cm, "iNVERTI mAIUSCOLE", lambda: self.case("invert"), "Ctrl+Alt+I")
        self.act(cm, "cAsUaLe", lambda: self.case("random"), "Ctrl+Alt+R")
        self.lines_menu = m.addMenu("Operazioni sulle righe")
        lm = self.lines_menu
        self.act(lm, "Duplica riga", lambda: self.line_op(ops.duplicate_lines), "Ctrl+D")
        self.act(lm, "Elimina riga", lambda: self.line_op(ops.delete_lines), "Ctrl+Shift+L")
        self.act(lm, "Sposta su", lambda: self.line_op(ops.move_lines, True), "Ctrl+Shift+Up")
        self.act(lm, "Sposta giù", lambda: self.line_op(ops.move_lines, False), "Ctrl+Shift+Down")
        self.act(lm, "Unisci righe", lambda: self.line_op(ops.join_lines), "Ctrl+J")
        lm.addSeparator()
        self.act(lm, "Ordina righe A → Z", lambda: self.line_op(ops.sort_lines, False))
        self.act(lm, "Ordina righe Z → A", lambda: self.line_op(ops.sort_lines, True))
        self.act(lm, "Inverti ordine delle righe", lambda: self.line_op(ops.reverse_lines))
        self.act(lm, "Rimuovi righe duplicate", lambda: self.line_op(ops.remove_duplicates, msg="{} righe duplicate rimosse"))
        self.act(lm, "Rimuovi righe vuote", lambda: self.line_op(ops.remove_empty, msg="{} righe vuote rimosse"))
        lm.addSeparator()
        self.act(lm, "Togli spazi a fine riga", lambda: self.line_op(ops.trim, "trailing"))
        self.act(lm, "Togli spazi a inizio riga", lambda: self.line_op(ops.trim, "leading"))
        self.act(lm, "Togli spazi a inizio e fine riga", lambda: self.line_op(ops.trim, "both"))
        self.act(lm, "Tabulazioni → spazi", lambda: self.line_op(ops.tabs_to_spaces))
        m.addSeparator()
        self.act(m, "Data/ora", self.insert_datetime, "F5")

        m = mb.addMenu("&Cerca")
        self.act(m, "Trova...", lambda: self.open_search("find"), QKeySequence.Find)
        self.act(m, "Trova successivo", lambda: self.quick_find(False), "F3")
        self.act(m, "Trova precedente", lambda: self.quick_find(True), "Shift+F3")
        self.act(m, "Trova la parola selezionata", self.find_selected, "Ctrl+F3")
        self.act(m, "Sostituisci...", lambda: self.open_search("replace"), "Ctrl+H")
        self.act(m, "Trova nei file...", lambda: self.open_search("files"), "Ctrl+Shift+F")
        m.addSeparator()
        self.act(m, "Cancella evidenziazioni", self.clear_marks)
        self.act(m, "Mostra/nascondi risultati", self.toggle_results, "F7")
        self.act(m, "Vai alla riga...", self.goto_line, "Ctrl+G")

        m = mb.addMenu("&Inserisci")
        self.heading_menu = m.addMenu("Intestazione")
        hm = self.heading_menu
        self.act(hm, "Titolo 1 (grande)", lambda: self.set_heading(1), "Ctrl+1")
        self.act(hm, "Titolo 2 (medio)", lambda: self.set_heading(2), "Ctrl+2")
        self.act(hm, "Titolo 3 (piccolo)", lambda: self.set_heading(3), "Ctrl+3")
        self.act(hm, "Testo normale", lambda: self.set_heading(0))
        hm.addSeparator()
        self.act(hm, "Linea di separazione", self.insert_rule, "Ctrl+Shift+R")
        m.addSeparator()
        a_tbl = self.act(m, "Tabella...", self.insert_table, "Ctrl+T")
        a_img = self.act(m, "Immagine da file...", self.insert_image_file, "Ctrl+Shift+I")

        self.table_menu = mb.addMenu("&Tabella")
        tm = self.table_menu
        self.act(tm, "Inserisci riga sopra", lambda: self.table_op("row_above"))
        self.act(tm, "Inserisci riga sotto", lambda: self.table_op("row_below"), "Ctrl+Return")
        self.act(tm, "Inserisci colonna a sinistra", lambda: self.table_op("col_left"))
        self.act(tm, "Inserisci colonna a destra", lambda: self.table_op("col_right"))
        tm.addSeparator()
        self.act(tm, "Elimina riga", lambda: self.table_op("del_row"))
        self.act(tm, "Elimina colonna", lambda: self.table_op("del_col"))
        self.act(tm, "Elimina tabella", lambda: self.table_op("del_table"))

        m = mb.addMenu("&Visualizza")
        self.a_wrap = self.act(m, "A capo automatico", lambda: self.toggle_view("wrap"))
        self.a_ws = self.act(m, "Mostra spazi e tabulazioni", lambda: self.toggle_view("show_ws"))
        self.a_hl = self.act(m, "Evidenzia la riga corrente", lambda: self.toggle_view("hl_line"))
        for a, v in ((self.a_wrap, self.wrap), (self.a_ws, self.show_ws), (self.a_hl, self.hl_line)):
            a.setCheckable(True)
            a.setChecked(v)
        self.act(m, "Font...", self.choose_font)
        m.addSeparator()
        self.act(m, "Zoom avanti", lambda: self.zoom(1), ["Ctrl++", "Ctrl+="])
        self.act(m, "Zoom indietro", lambda: self.zoom(-1), "Ctrl+-")
        self.act(m, "Zoom predefinito", lambda: self.zoom(0), "Ctrl+0")
        m.addSeparator()
        self.act(m, "Scheda successiva", lambda: self.cycle_tab(1), ["Ctrl+Tab", "Ctrl+PgDown"])
        self.act(m, "Scheda precedente", lambda: self.cycle_tab(-1), ["Ctrl+Shift+Tab", "Ctrl+PgUp"])
        self.act(m, "Statistiche del documento...", self.show_stats)
        m.addSeparator()
        a_side = self.side.toggleViewAction()
        a_side.setText("Pannello file aperti e struttura")
        a_side.setShortcut(QKeySequence("F8"))
        m.addAction(a_side)
        self.addAction(a_side)

        m = mb.addMenu("&?")
        self.act(m, "Informazioni su FastText", self.about)

        tb = QToolBar("Strumenti")
        tb.setObjectName("MainToolbar")
        tb.setMovable(False)
        tb.setIconSize(QSize(18, 18))
        tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(tb)
        for a in (a_new, a_open, a_save, a_saveall):
            tb.addAction(a)
        a_saveall.setIcon(self.style().standardIcon(S.SP_DialogSaveButton))
        tb.addSeparator()
        hb = QToolButton()
        hb.setText("Intestazione")
        hb.setToolTip("Titoli per dividere il testo in sezioni (Ctrl+1, Ctrl+2, Ctrl+3)")
        hb.setIcon(self.style().standardIcon(S.SP_ArrowRight))
        hb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        hb.setPopupMode(QToolButton.InstantPopup)
        hb.setMenu(self.heading_menu)
        tb.addWidget(hb)
        a_tbl.setIcon(self.style().standardIcon(S.SP_FileDialogDetailedView))
        a_img.setIcon(self.style().standardIcon(S.SP_DirIcon))
        tb.addAction(a_tbl)
        tb.addAction(a_img)
        tb.addSeparator()
        a_find = QAction(self.style().standardIcon(S.SP_FileDialogContentsView), "Trova", self)
        a_find.triggered.connect(lambda: self.open_search("find"))
        tb.addAction(a_find)

        # scorciatoie che l'editor deve lasciar passare alla finestra
        self.reserved_keys = set()
        for a in self.actions():
            if a.shortcutContext() != Qt.WidgetShortcut:
                for k in a.shortcuts():
                    self.reserved_keys.add(k.toString())

    # ---- stato --------------------------------------------------------
    def update_title(self):
        t = self.tabs.currentWidget()
        if t is None:
            return
        star = "*" if t.modified() else ""
        where = t.path or t.title()
        extra = "  [Word, vista semplificata]" if t.kind == "docx" else ""
        self.setWindowTitle(f"{star}{where}{extra} - {APP_NAME}")
        self.status_kind.setText(KIND_LABEL.get(t.kind, "") + "   ")

    def update_status(self):
        t = self.tabs.currentWidget()
        if t is None:
            return
        c = t.editor.textCursor()
        tbl = c.currentTable()
        sel = len(c.selectedText())
        sel_txt = f"   Sel: {sel}" if sel else ""
        if tbl is not None:
            cell = tbl.cellAt(c)
            self.status.setText(f"Tabella: riga {cell.row() + 1}, colonna {cell.column() + 1}{sel_txt}   ")
        else:
            self.status.setText(f"Riga {c.blockNumber() + 1}, Col {c.positionInBlock() + 1}{sel_txt}   ")

    def style_editor(self, ed):
        f = QFont(self.base_font)
        size = max(6.0, (self.base_font.pointSizeF() if self.base_font.pointSizeF() > 0 else 11) + self.zoom_delta)
        f.setPointSizeF(size)
        ed.setFont(f)
        ed.document().setDefaultFont(f)
        ed.setTabStopDistance(ed.fontMetrics().horizontalAdvance(" ") * 4)
        ed.setLineWrapMode(QTextEdit.WidgetWidth if self.wrap else QTextEdit.NoWrap)
        opt = ed.document().defaultTextOption()
        flags = opt.flags()
        if self.show_ws:
            flags |= QTextOption.ShowTabsAndSpaces
        else:
            flags &= ~QTextOption.ShowTabsAndSpaces
        opt.setFlags(flags)
        ed.document().setDefaultTextOption(opt)
        ed.highlight_line = self.hl_line
        ed.refresh_extra()

    def restyle_all(self):
        for t in self.all_tabs():
            self.style_editor(t.editor)
        self.results.set_font(QFont(self.base_font.family(), 10))

    def toggle_view(self, what):
        val = not getattr(self, what)
        setattr(self, what, val)
        self.settings.setValue(what, val)
        self.restyle_all()

    def zoom(self, step):
        self.zoom_delta = 0 if step == 0 else self.zoom_delta + step
        self.restyle_all()

    def choose_font(self):
        ok, f = QFontDialog.getFont(self.base_font, self, "Font (unico per tutto il testo)")
        if isinstance(ok, QFont):
            ok, f = f, ok
        if ok:
            self.base_font = f
            self.zoom_delta = 0
            self.restyle_all()
            self.settings.setValue("font", f.toString())

    # ---- file ---------------------------------------------------------
    def maybe_save(self, t):
        if not t.modified():
            return True
        self.tabs.setCurrentWidget(t)
        r = QMessageBox.question(self, APP_NAME, f"Salvare le modifiche a \"{t.title()}\"?",
                                 QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        if r == QMessageBox.Save:
            return self.save(t)
        return r == QMessageBox.Discard

    def add_recent(self, path):
        rec = [p for p in (self.settings.value("recent", []) or []) if os.path.normcase(p) != os.path.normcase(path)]
        if isinstance(rec, str):
            rec = [rec]
        rec.insert(0, path)
        self.settings.setValue("recent", rec[:15])

    def fill_recent(self):
        self.recent_menu.clear()
        rec = self.settings.value("recent", []) or []
        if isinstance(rec, str):
            rec = [rec]
        for i, p in enumerate(rec):
            a = self.recent_menu.addAction(f"&{(i + 1) % 10}  {p}" if i < 10 else p, lambda p=p: self.open_path(p))
            a.setEnabled(os.path.isfile(p))
        if not rec:
            self.recent_menu.addAction("(nessuno)").setEnabled(False)
        else:
            self.recent_menu.addSeparator()
            self.recent_menu.addAction("Svuota elenco", lambda: self.settings.setValue("recent", []))

    def open_dialog(self):
        cur = self.tabs.currentWidget()
        start = os.path.dirname(cur.path) if cur is not None and cur.path else self.settings.value("lastdir", "")
        paths, _ = QFileDialog.getOpenFileNames(self, "Apri", start, FILTER_OPEN)
        for p in paths:
            self.open_path(p)

    def open_path(self, path, select=True):
        path = os.path.abspath(path)
        existing = self.tab_for_path(path)
        if existing is not None:
            if select:
                self.tabs.setCurrentWidget(existing)
            return existing
        store = ImageStore()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            model, kind = load_any(path, store)
        except Exception as ex:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, APP_NAME, f"Impossibile aprire il file:\n{path}\n\n{ex}")
            return None
        cur = self.tabs.currentWidget()
        t = cur if (cur is not None and cur.pristine()) else self.new_tab()
        t.store.data = store.data
        t.editor.set_model(model)
        QApplication.restoreOverrideCursor()
        t.path, t.kind = path, kind
        self.settings.setValue("lastdir", os.path.dirname(path))
        self.add_recent(path)
        self.update_tab(t)
        if select:
            self.tabs.setCurrentWidget(t)
        self.update_title()
        self.update_status()
        return t

    def reload_tab(self, t):
        if not t.path or not os.path.isfile(t.path):
            return
        if t.modified():
            r = QMessageBox.question(self, APP_NAME, f"Ricaricare \"{t.title()}\" dal disco?\nLe modifiche non salvate andranno perse.")
            if r != QMessageBox.Yes:
                return
        store = ImageStore()
        try:
            model, kind = load_any(t.path, store)
        except Exception as ex:
            QMessageBox.critical(self, APP_NAME, f"Impossibile ricaricare:\n{ex}")
            return
        t.store.data = store.data
        t.editor.set_model(model)
        t.kind = kind
        self.update_tab(t)

    def save(self, t=None):
        t = t or self.current
        if not t.path or t.kind == "md":
            return self.save_as(t)
        model = t.editor.get_model()
        if t.kind == "docx":
            self.tabs.setCurrentWidget(t)
            box = QMessageBox(self)
            box.setWindowTitle(APP_NAME)
            box.setIcon(QMessageBox.Question)
            box.setText(f"\"{t.title()}\" è un documento Word.\n\nSovrascrivendolo, la formattazione originale "
                        "(font, colori, stili, intestazioni) andrà persa: resteranno solo testo, "
                        "tabelle e immagini.")
            b_ftx = box.addButton("Salva come .ftx", QMessageBox.AcceptRole)
            b_docx = box.addButton("Sovrascrivi il .docx", QMessageBox.DestructiveRole)
            box.addButton("Annulla", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() == b_ftx:
                return self.save_as(t, os.path.splitext(t.path)[0] + ".ftx")
            if box.clickedButton() != b_docx:
                return False
        if t.kind == "txt" and model_has_rich(model):
            self.tabs.setCurrentWidget(t)
            box = QMessageBox(self)
            box.setWindowTitle(APP_NAME)
            box.setIcon(QMessageBox.Question)
            box.setText(f"\"{t.title()}\": un file .txt non può contenere immagini né tabelle vere.\n\n"
                        "Vuoi salvarlo come .ftx?")
            b_ftx = box.addButton("Salva come .ftx", QMessageBox.AcceptRole)
            b_txt = box.addButton("Salva .txt (senza immagini)", QMessageBox.DestructiveRole)
            box.addButton("Annulla", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() == b_ftx:
                return self.save_as(t, os.path.splitext(t.path)[0] + ".ftx")
            if box.clickedButton() != b_txt:
                return False
        return self.write_file(t, t.path, t.kind, model)

    def save_as(self, t=None, suggested=None):
        t = t or self.current
        self.tabs.setCurrentWidget(t)
        if not suggested:
            base = os.path.splitext(t.path)[0] if t.path else os.path.join(
                self.settings.value("lastdir", "") or os.path.expanduser("~"), t.title())
            suggested = base + ".ftx"
        path, flt = QFileDialog.getSaveFileName(self, "Salva con nome", suggested, FILTER_SAVE)
        if not path:
            return False
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".ftx", ".docx", ".txt"):
            ext = ".docx" if "docx" in flt else ".txt" if "txt" in flt else ".ftx"
            path += ext
        other = self.tab_for_path(path)
        if other is not None and other is not t:
            QMessageBox.warning(self, APP_NAME, "Quel file è già aperto in un'altra scheda: chiudila prima.")
            return False
        kind = ext[1:]
        if self.write_file(t, path, kind, t.editor.get_model()):
            t.path, t.kind = path, kind
            self.settings.setValue("lastdir", os.path.dirname(path))
            self.add_recent(path)
            self.update_tab(t)
            return True
        return False

    def save_all(self):
        for t in self.all_tabs():
            if t.modified():
                if not self.save(t):
                    return

    def write_file(self, t, path, kind, model):
        try:
            if kind == "ftx":
                save_ftx(path, model, t.store)
            elif kind == "docx":
                save_docx(path, model, t.store, self.base_font)
            else:
                with open(path, "w", encoding="utf-8", newline="\r\n") as f:
                    f.write(model_to_plain(model))
        except Exception as ex:
            QMessageBox.critical(self, APP_NAME, f"Errore durante il salvataggio:\n{ex}")
            return False
        t.editor.document().setModified(False)
        self.update_tab(t)
        self.statusBar().showMessage(f"Salvato: {path}", 3000)
        return True

    def export(self, what):
        t = self.current
        base = os.path.splitext(t.path)[0] if t.path else t.title()
        if what == "docx":
            path, _ = QFileDialog.getSaveFileName(self, "Esporta in Word", base + ".docx", "Word (*.docx)")
            if path:
                try:
                    save_docx(path, t.editor.get_model(), t.store, self.base_font)
                    self.statusBar().showMessage(f"Esportato: {path}", 3000)
                except Exception as ex:
                    QMessageBox.critical(self, APP_NAME, f"Errore:\n{ex}")
        else:
            path, _ = QFileDialog.getSaveFileName(self, "Esporta in PDF", base + ".pdf", "PDF (*.pdf)")
            if path:
                from PySide6.QtPrintSupport import QPrinter
                pr = QPrinter(QPrinter.HighResolution)
                pr.setOutputFormat(QPrinter.PdfFormat)
                pr.setOutputFileName(path)
                self.print_document(pr)
                self.statusBar().showMessage(f"Esportato: {path}", 3000)

    def print_document(self, printer, t=None):
        t = t or self.current
        doc = t.editor.document().clone()
        doc.setDefaultFont(self.base_font)
        for name in model_image_names(t.editor.get_model()):
            doc.addResource(QTextDocument.ImageResource, QUrl(name), t.store.qimage(name))
        doc.print_(printer)

    def print_doc(self):
        from PySide6.QtPrintSupport import QPrinter, QPrintDialog
        pr = QPrinter(QPrinter.HighResolution)
        if QPrintDialog(pr, self).exec() == QDialog.Accepted:
            self.print_document(pr)

    # ---- ricerca ------------------------------------------------------
    def open_search(self, page):
        if self.search_dlg is None:
            self.search_dlg = SearchDialog(self)
        self.search_dlg.show_page(page)

    def quick_find(self, backward):
        if self.last_opts is None or not self.last_opts.text:
            self.open_search("find")
            return
        try:
            msg = self.search.do("prev" if backward else "next", self.last_opts)
        except ValueError as e:
            msg = str(e)
        if msg:
            self.statusBar().showMessage(msg, 3000)

    def find_selected(self):
        c = self.editor.textCursor()
        if not c.hasSelection():
            c.select(QTextCursor.WordUnderCursor)
            self.editor.setTextCursor(c)
        word = c.selectedText()
        if not word or " " in word or OBJ in word:
            return
        base = self.last_opts or Opts()
        self.last_opts = Opts(word, base.repl, "normal", base.case, base.word, True)
        self.quick_find(False)

    def clear_marks(self):
        for t in self.all_tabs():
            t.editor.clear_marks()

    def toggle_results(self):
        self.results.setVisible(not self.results.isVisible())

    def goto_hit(self, hit):
        t = hit.get("tab")
        if t is None or t not in self.all_tabs():
            t = self.open_path(hit["path"]) if hit.get("path") else None
        if t is None:
            return
        self.tabs.setCurrentWidget(t)
        doc = t.editor.document()
        b = doc.findBlockByNumber(hit["line"])
        if not b.isValid():
            return
        s = b.position() + min(hit["col"], b.length() - 1)
        e = min(s + hit["len"], doc.characterCount() - 1)
        c = QTextCursor(doc)
        c.setPosition(s)
        c.setPosition(e, QTextCursor.KeepAnchor)
        t.editor.setTextCursor(c)
        t.editor.ensureCursorVisible()
        t.editor.setFocus()

    def goto_line(self):
        from PySide6.QtWidgets import QInputDialog
        doc = self.editor.document()
        n, ok = QInputDialog.getInt(self, "Vai alla riga", f"Riga (1 - {doc.blockCount()}):",
                                    self.editor.textCursor().blockNumber() + 1, 1, doc.blockCount())
        if ok:
            c = QTextCursor(doc.findBlockByNumber(n - 1))
            self.editor.setTextCursor(c)
            self.editor.ensureCursorVisible()

    # ---- testo --------------------------------------------------------
    def case(self, mode):
        ops.convert_case(self.editor, mode)

    def line_op(self, fn, *args, msg=None):
        try:
            r = fn(self.editor, *args)
        except ops.LineOpError as e:
            self.statusBar().showMessage(str(e), 5000)
            return
        if msg and r is not None:
            self.statusBar().showMessage(msg.format(r), 3000)

    def show_stats(self):
        st = ops.text_stats(self.editor.document())
        t = self.current
        sel = len(self.editor.textCursor().selectedText())
        QMessageBox.information(self, "Statistiche", (
            f"<b>{t.title()}</b><br>{t.path or ''}<br><br>"
            f"Righe: {st['righe']}<br>Parole: {st['parole']}<br>"
            f"Caratteri (senza a capo): {st['caratteri']}<br>Caratteri senza spazi: {st['caratteri_senza_spazi']}<br>"
            f"Immagini: {st['immagini']}<br>Selezione: {sel} caratteri"))

    # ---- inserimenti --------------------------------------------------
    def insert_table(self):
        d = TableDialog(self)
        if d.exec() != QDialog.Accepted:
            return
        ed = self.editor
        c = ed.textCursor()
        c.beginEditBlock()
        if c.currentTable() is not None:
            c.setPosition(c.currentTable().lastPosition() + 1)
        tbl = c.insertTable(d.rows.value(), d.cols.value(), make_table_format())
        c.endEditBlock()
        ed.setTextCursor(tbl.cellAt(0, 0).firstCursorPosition())
        ed.setFocus()

    def insert_image_file(self):
        t = self.current
        paths, _ = QFileDialog.getOpenFileNames(self, "Inserisci immagine", self.settings.value("lastdir", ""),
                                                "Immagini (*.png *.jpg *.jpeg *.gif *.bmp *.webp)")
        c = t.editor.textCursor()
        for p in paths:
            with open(p, "rb") as f:
                name = t.store.add_bytes(f.read())
            if name:
                t.editor.insert_image_name(c, name)
            else:
                QMessageBox.warning(self, APP_NAME, f"Formato immagine non supportato:\n{p}")
        t.editor.setFocus()

    def insert_datetime(self):
        c = self.editor.textCursor()
        c.insertText(QDateTime.currentDateTime().toString("HH:mm dd/MM/yyyy"), char_format_at(c))

    def set_heading(self, level):
        ed = self.editor
        c = ed.textCursor()
        if c.currentTable() is not None:
            self.statusBar().showMessage("Le intestazioni si mettono fuori dalle tabelle", 4000)
            return
        doc = ed.document()
        b1 = doc.findBlock(c.selectionStart())
        b2 = doc.findBlock(c.selectionEnd())
        cur_lvl = b1.blockFormat().headingLevel()
        new = 0 if (level == 0 or level == cur_lvl) else level     # stesso tasto = torna testo normale
        pos, anchor = c.position(), c.anchor()
        edit = QTextCursor(doc)
        edit.beginEditBlock()
        b = b1
        while b.isValid():
            bc = QTextCursor(b)
            if bc.currentTable() is None and not is_rule_format(b.blockFormat()):
                bc.setBlockFormat(heading_block_format(new))
                bc.setBlockCharFormat(heading_char_format(new))
                bc.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
                if bc.hasSelection():
                    bc.mergeCharFormat(heading_char_format(new))
            if b.blockNumber() >= b2.blockNumber():
                break
            b = b.next()
        edit.endEditBlock()
        nc = QTextCursor(doc)
        nc.setPosition(anchor)
        nc.setPosition(pos, QTextCursor.KeepAnchor)
        ed.setTextCursor(nc)
        self.side.refresh_outline()

    def insert_rule(self):
        ed = self.editor
        c = ed.textCursor()
        if c.currentTable() is not None:
            self.statusBar().showMessage("La linea di separazione si mette fuori dalle tabelle", 4000)
            return
        c.beginEditBlock()
        c.clearSelection()
        bf = c.blockFormat()
        if not c.block().text() and bf.headingLevel() == 0 and not is_rule_format(bf):
            c.setBlockFormat(rule_block_format())
        else:
            c.movePosition(QTextCursor.EndOfBlock)
            c.insertBlock(rule_block_format(), QTextCharFormat())
        c.insertBlock(QTextBlockFormat(), QTextCharFormat())
        c.endEditBlock()
        ed.setTextCursor(c)
        ed.setFocus()

    def table_op(self, op):
        ed = self.editor
        c = ed.textCursor()
        tbl = c.currentTable()
        if tbl is None:
            self.statusBar().showMessage("Metti il cursore dentro una tabella", 3000)
            return
        cell = tbl.cellAt(c)
        r, col = cell.row(), cell.column()
        c.beginEditBlock()
        if op == "row_above":
            tbl.insertRows(r, 1)
        elif op == "row_below":
            tbl.insertRows(r + 1, 1)
        elif op == "col_left":
            tbl.insertColumns(col, 1)
        elif op == "col_right":
            tbl.insertColumns(col + 1, 1)
        elif op == "del_row":
            tbl.removeRows(r, 1)
        elif op == "del_col":
            tbl.removeColumns(col, 1)
        elif op == "del_table":
            tbl.removeRows(0, tbl.rows())
        c.endEditBlock()
        if op == "row_below":
            ed.setTextCursor(tbl.cellAt(r + 1, col).firstCursorPosition())

    def about(self):
        QMessageBox.about(self, APP_NAME,
                          f"<b>{APP_NAME} {APP_VERSION}</b><br><br>"
                          "Un Blocco note con tabelle e immagini, a schede.<br>"
                          "Un solo font, nessuna formattazione: solo testo, tabelle e immagini.<br><br>"
                          "Formato .ftx = ZIP con <i>content.md</i> + cartella <i>images</i>.<br>"
                          "Apre anche i file Word (.docx) in vista semplificata.")

    # ---- sessione -----------------------------------------------------
    def restore_session(self):
        paths = self.settings.value("session", []) or []
        if isinstance(paths, str):
            paths = [paths]
        cur = self.settings.value("session_current", "")
        for p in paths:
            if os.path.isfile(p):
                self.open_path(p, select=False)
        if cur and self.tab_for_path(cur):
            self.tabs.setCurrentWidget(self.tab_for_path(cur))

    def closeEvent(self, e):
        for t in self.all_tabs():
            if not self.maybe_save(t):
                e.ignore()
                return
        self.settings.setValue("session", [t.path for t in self.all_tabs() if t.path])
        cur = self.tabs.currentWidget()
        self.settings.setValue("session_current", cur.path if cur is not None and cur.path else "")
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("winstate", self.saveState())
        e.accept()

    def bring_to_front(self):
        if self.isMinimized():
            self.showNormal()
        self.show()
        self.raise_()
        self.activateWindow()


# ---------------------------------------------------------------------------
#  Avvio (una sola finestra: i file aperti dopo vanno in nuove schede)
# ---------------------------------------------------------------------------

def send_to_running(paths):
    sock = QLocalSocket()
    sock.connectToServer(SERVER_NAME)
    if not sock.waitForConnected(300):
        return False
    sock.write(json.dumps(paths).encode("utf-8"))
    sock.flush()
    sock.waitForBytesWritten(1000)
    sock.disconnectFromServer()
    return True


def start_server(win):
    server = QLocalServer(win)
    QLocalServer.removeServer(SERVER_NAME)
    server.listen(SERVER_NAME)

    def on_conn():
        sock = server.nextPendingConnection()
        buf = bytearray()

        def read():
            buf.extend(bytes(sock.readAll()))

        def done():
            read()
            try:
                paths = json.loads(buf.decode("utf-8"))
            except Exception:
                paths = []
            for p in paths:
                if os.path.isfile(p):
                    win.open_path(p)
            win.bring_to_front()
            sock.deleteLater()

        sock.readyRead.connect(read)
        sock.disconnected.connect(done)

    server.newConnection.connect(on_conn)
    return server


def main():
    QApplication.setApplicationName(APP_NAME)
    app = QApplication(sys.argv)
    args = [os.path.abspath(a) for a in sys.argv[1:] if not a.startswith("-")]
    if "--new-window" not in sys.argv and send_to_running(args):
        return
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("FastText.Editor")
        except Exception:
            pass
    app.setWindowIcon(make_icon())
    w = MainWindow()
    w._server = start_server(w)
    if w.settings.value("restore_session", True, type=bool):
        w.restore_session()
    for a in args:
        if os.path.isfile(a):
            w.open_path(a)
    if w.tabs.count() == 0:
        w.new_tab()
    w.restyle_all()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
