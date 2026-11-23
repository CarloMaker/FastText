# -*- coding: utf-8 -*-
"""
FastText - ricerca avanzata stile Notepad++.

Modalità: normale, estesa (\\n \\t \\xHH ...), espressione regolare.
Opzioni: parola intera, maiuscole/minuscole, ricomincia dall'inizio, solo nella selezione.
Trova tutti (documento / tutti i tab), sostituisci, evidenzia, trova e sostituisci nei file.
"""

import fnmatch
import html
import os
import re

from PySide6.QtCore import Qt, Signal, QSettings, QSize, QRectF
from PySide6.QtGui import QTextCursor, QTextCharFormat, QTextDocument, QAbstractTextDocumentLayout, QPalette, QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDockWidget, QFileDialog, QFormLayout, QGridLayout,
    QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMenu, QMessageBox, QProgressDialog, QPushButton,
    QRadioButton, QStyle, QStyledItemDelegate, QStyleOptionViewItem, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget, QButtonGroup, QSizePolicy,
)

from ftcore import OBJ, ImageStore, text_format, is_dark, build_headless, doc_to_model, load_any, model_to_plain, save_ftx

FRAME_CHARS = "\ufdd0\ufdd1"
_TRANS = {ord(c): "\n" for c in "\u2029\u2028\ufdd0\ufdd1"}
MAX_HITS = 20000


# ---------------------------------------------------------------------------
#  Motore
# ---------------------------------------------------------------------------

class Opts:
    def __init__(self, text="", repl="", mode="normal", case=False, word=False, wrap=True,
                 dotall=False, insel=False):
        self.text, self.repl, self.mode = text, repl, mode
        self.case, self.word, self.wrap, self.dotall, self.insel = case, word, wrap, dotall, insel


def unescape_extended(s):
    """Modalità estesa: \\n \\r \\t \\0 \\\\ \\xHH \\uHHHH \\dNNN."""
    out, i = [], 0
    while i < len(s):
        ch = s[i]
        if ch == "\\" and i + 1 < len(s):
            n = s[i + 1]
            simple = {"n": "\n", "r": "\r", "t": "\t", "0": "\0", "\\": "\\"}
            if n in simple:
                out.append(simple[n]); i += 2; continue
            if n == "x" and re.fullmatch(r"[0-9a-fA-F]{2}", s[i + 2:i + 4]):
                out.append(chr(int(s[i + 2:i + 4], 16))); i += 4; continue
            if n == "u" and re.fullmatch(r"[0-9a-fA-F]{4}", s[i + 2:i + 6]):
                out.append(chr(int(s[i + 2:i + 6], 16))); i += 6; continue
            if n == "d" and re.fullmatch(r"[0-9]{3}", s[i + 2:i + 5]):
                out.append(chr(int(s[i + 2:i + 5]))); i += 5; continue
        out.append(ch)
        i += 1
    # nel documento le righe sono separate solo da \n
    return "".join(out).replace("\r\n", "\n").replace("\r", "\n")


def compile_pattern(o):
    if not o.text:
        raise ValueError("Scrivi cosa cercare.")
    if o.mode == "regex":
        body = o.text
    else:
        body = re.escape(unescape_extended(o.text) if o.mode == "extended" else o.text)
    if o.word:
        body = r"(?<!\w)(?:" + body + r")(?!\w)"
    flags = re.MULTILINE
    if not o.case:
        flags |= re.IGNORECASE
    if o.mode == "regex" and o.dotall:
        flags |= re.DOTALL
    try:
        return re.compile(body, flags)
    except re.error as e:
        raise ValueError(f"Espressione regolare non valida: {e}")


def replacement(o, m):
    if o.mode == "regex":
        tpl = re.sub(r"\$(\d+)", r"\\g<\1>", o.repl)        # $1 come in Notepad++
        tpl = re.sub(r"\\0(?![0-9])", r"\\g<0>", tpl)
        try:
            return m.expand(tpl)
        except (re.error, IndexError) as e:
            raise ValueError(f"Testo di sostituzione non valido: {e}")
    if o.mode == "extended":
        return unescape_extended(o.repl)
    return o.repl


def find_matches(doc, pat, lo=0, hi=None, nonempty=True):
    """Corrispondenze nel documento. Le posizioni coincidono con quelle del QTextDocument."""
    raw = doc.toRawText()
    view = raw.translate(_TRANS)
    if hi is None:
        hi = len(view)
    out = []
    for m in pat.finditer(view, lo, hi):
        s, e = m.span()
        if nonempty and s == e:
            continue
        if any(c in FRAME_CHARS for c in raw[s:e]):     # non attraversare i bordi delle tabelle
            continue
        out.append(m)
        if len(out) >= MAX_HITS:
            break
    return out


def find_next(doc, pat, cursor, backward, wrap, rng=None):
    lo, hi = rng if rng else (0, None)
    ms = find_matches(doc, pat, lo, hi)
    if not ms:
        return None
    s0, e0 = cursor.selectionStart(), cursor.selectionEnd()
    if backward:
        cand = [m for m in ms if m.end() <= s0 and not (m.start() == s0)]
        if cand:
            return cand[-1], False
        return (ms[-1], True) if wrap else None
    cand = [m for m in ms if m.start() >= e0 and not (m.start() == s0 and m.end() == e0)]
    if cand:
        return cand[0], False
    return (ms[0], True) if wrap else None


def replace_all_in_doc(doc, o, pat, lo=0, hi=None):
    ms = [m for m in find_matches(doc, pat, lo, hi, nonempty=False) if OBJ not in m.group(0)]
    reps = [(m.start(), m.end(), replacement(o, m)) for m in ms]
    if not reps:
        return 0
    cur = QTextCursor(doc)
    cur.beginEditBlock()
    for s, e, t in reversed(reps):
        cur.setPosition(s)
        cur.setPosition(e, QTextCursor.KeepAnchor)
        if t:
            cur.insertText(t, text_format(cur))
        else:
            cur.removeSelectedText()
    cur.endEditBlock()
    return len(reps)


def hits_from_matches(doc, ms):
    hits = []
    for m in ms:
        b = doc.findBlock(m.start())
        col = m.start() - b.position()
        hits.append({"line": b.blockNumber(), "col": col, "len": m.end() - m.start(),
                     "text": b.text()})
    return hits


def hit_html(text, col, ln):
    """Riga del risultato con la parte trovata evidenziata."""
    start = max(0, col - 80)
    end = min(len(text), col + ln + 120)
    pre, mid, post = text[start:col], text[col:col + ln], text[col + ln:end]
    f = lambda s: html.escape(s.replace(OBJ, "▣"))
    return (("…" if start else "") + f(pre) + '<span style="background:#ffe36e;color:#000000;font-weight:bold">'
            + f(mid) + "</span>" + f(post) + ("…" if end < len(text) else ""))


# ---------------------------------------------------------------------------
#  Pannello risultati (in basso, come Notepad++)
# ---------------------------------------------------------------------------

HTML_ROLE = Qt.UserRole + 1
HIT_ROLE = Qt.UserRole


class HtmlDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        h = index.data(HTML_ROLE)
        if not h:
            return super().paint(painter, option, index)
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.CE_ItemViewItem, opt, painter, opt.widget)
        doc = QTextDocument()
        doc.setDefaultFont(opt.font)
        doc.setDocumentMargin(0)
        doc.setHtml(h)
        rect = style.subElementRect(QStyle.SE_ItemViewItemText, opt, opt.widget)
        painter.save()
        painter.translate(rect.left(), rect.top() + max(0, (rect.height() - doc.size().height()) / 2))
        painter.setClipRect(QRectF(0, 0, rect.width(), rect.height()))
        ctx = QAbstractTextDocumentLayout.PaintContext()
        if opt.state & QStyle.State_Selected:
            ctx.palette.setColor(QPalette.Text, opt.palette.color(QPalette.HighlightedText))
        doc.documentLayout().draw(painter, ctx)
        painter.restore()


class ResultsDock(QDockWidget):
    hitActivated = Signal(dict)

    def __init__(self, parent=None):
        super().__init__("Risultati della ricerca", parent)
        self.setObjectName("ResultsDock")
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setItemDelegate(HtmlDelegate(self.tree))
        self.tree.itemActivated.connect(self._activated)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._menu)
        self.setWidget(self.tree)

    def set_font(self, font):
        self.tree.setFont(font)

    def _activated(self, item, _col):
        hit = item.data(0, HIT_ROLE)
        if hit:
            self.hitActivated.emit(hit)

    def _menu(self, pos):
        m = QMenu(self)
        m.addAction("Copia righe selezionate", self.copy_selected)
        m.addSeparator()
        m.addAction("Espandi tutto", self.tree.expandAll)
        m.addAction("Comprimi tutto", self.tree.collapseAll)
        m.addSeparator()
        m.addAction("Cancella tutti i risultati", self.tree.clear)
        m.exec(self.tree.viewport().mapToGlobal(pos))

    def copy_selected(self):
        lines = []
        for it in self.tree.selectedItems():
            hit = it.data(0, HIT_ROLE)
            lines.append(hit["text"].replace(OBJ, "") if hit else it.text(0))
        QApplication.clipboard().setText("\n".join(lines))

    def add_search(self, title, groups):
        """groups: lista di (etichetta file, percorso, tab, [hit])"""
        for i in range(self.tree.topLevelItemCount()):
            self.tree.topLevelItem(i).setExpanded(False)
        top = QTreeWidgetItem([title])
        f = top.font(0); f.setBold(True); top.setFont(0, f)
        self.tree.insertTopLevelItem(0, top)
        for label, path, tab, hits in groups:
            g = QTreeWidgetItem([f"{label}  ({len(hits)})"])
            g.setToolTip(0, path or label)
            top.addChild(g)
            for h in hits:
                h = dict(h, path=path, tab=tab)
                it = QTreeWidgetItem([f"Riga {h['line'] + 1}: {h['text']}"])
                it.setData(0, HIT_ROLE, h)
                it.setData(0, HTML_ROLE, f'<span style="color:{"#7fb2ff" if is_dark() else "#2b6cb0"}">Riga {h["line"] + 1}:</span>&nbsp; '
                                         + hit_html(h["text"], h["col"], h["len"]))
                g.addChild(it)
            g.setExpanded(True)
        top.setExpanded(True)
        self.tree.scrollToTop()
        self.show()
        self.raise_()


# ---------------------------------------------------------------------------
#  Finestra di ricerca
# ---------------------------------------------------------------------------

class SearchDialog(QDialog):
    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.settings = QSettings("FastText", "FastText")
        self.setWindowTitle("Trova e sostituisci")
        self.setModal(False)

        self.cb_find = self._combo("hist_find")
        self.cb_repl = self._combo("hist_repl")
        form = QFormLayout()
        form.addRow("Trova:", self.cb_find)
        form.addRow("Sostituisci con:", self.cb_repl)

        self.chk_word = QCheckBox("Solo parole intere")
        self.chk_case = QCheckBox("Maiuscole/minuscole")
        self.chk_wrap = QCheckBox("Ricomincia dall'inizio")
        self.chk_insel = QCheckBox("Solo nella selezione")
        opt_box = QGroupBox("Opzioni")
        ol = QVBoxLayout(opt_box)
        for w in (self.chk_word, self.chk_case, self.chk_wrap, self.chk_insel):
            ol.addWidget(w)

        self.rb_normal = QRadioButton("Normale")
        self.rb_ext = QRadioButton("Estesa (\\n, \\t, \\0, \\x..)")
        self.rb_regex = QRadioButton("Espressione regolare")
        self.chk_dotall = QCheckBox(". trova anche l'a capo")
        self.mode_group = QButtonGroup(self)
        mode_box = QGroupBox("Modalità di ricerca")
        ml = QVBoxLayout(mode_box)
        for rb in (self.rb_normal, self.rb_ext, self.rb_regex):
            self.mode_group.addButton(rb)
            ml.addWidget(rb)
        ml.addWidget(self.chk_dotall)
        self.rb_regex.toggled.connect(self.chk_dotall.setEnabled)

        opts_row = QHBoxLayout()
        opts_row.addWidget(opt_box)
        opts_row.addWidget(mode_box)

        # cerca nei file
        self.files_box = QGroupBox("Cerca nei file di una cartella")
        self.files_box.setCheckable(True)
        self.files_inner = QWidget()
        fl = QGridLayout(self.files_inner)
        fl.setContentsMargins(0, 0, 0, 0)
        self.le_dir = QLineEdit()
        b_browse = QPushButton("...")
        b_browse.setFixedWidth(32)
        b_browse.clicked.connect(self._browse)
        self.le_filter = QLineEdit()
        self.chk_sub = QCheckBox("Anche nelle sottocartelle")
        self.btn_ffind = QPushButton("Trova nei file")
        self.btn_frepl = QPushButton("Sostituisci nei file")
        fl.addWidget(QLabel("Cartella:"), 0, 0)
        fl.addWidget(self.le_dir, 0, 1)
        fl.addWidget(b_browse, 0, 2)
        fl.addWidget(QLabel("Filtri:"), 1, 0)
        fl.addWidget(self.le_filter, 1, 1, 1, 2)
        fl.addWidget(self.chk_sub, 2, 1)
        hb = QHBoxLayout()
        hb.addWidget(self.btn_ffind)
        hb.addWidget(self.btn_frepl)
        hb.addStretch()
        fl.addLayout(hb, 3, 1, 1, 2)
        fbl = QVBoxLayout(self.files_box)
        fbl.addWidget(self.files_inner)
        self.files_box.toggled.connect(self.files_inner.setVisible)

        left = QVBoxLayout()
        left.addLayout(form)
        left.addLayout(opts_row)
        left.addWidget(self.files_box)
        left.addStretch()

        def btn(text, slot, default=False):
            b = QPushButton(text)
            b.clicked.connect(slot)
            b.setAutoDefault(default)
            b.setDefault(default)
            b.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            right.addWidget(b)
            return b

        right = QVBoxLayout()
        btn("Trova successivo", lambda: self.run("next"), True)
        btn("Trova precedente", lambda: self.run("prev"))
        btn("Conta", lambda: self.run("count"))
        btn("Trova tutti nel documento", lambda: self.run("all_doc"))
        btn("Trova tutti nei tab aperti", lambda: self.run("all_tabs"))
        right.addSpacing(10)
        btn("Sostituisci", lambda: self.run("replace"))
        btn("Sostituisci tutti", lambda: self.run("replace_all"))
        btn("Sostituisci tutti nei tab aperti", lambda: self.run("replace_tabs"))
        right.addSpacing(10)
        btn("Evidenzia tutti", lambda: self.run("mark"))
        btn("Cancella evidenziazioni", lambda: self.win.clear_marks())
        right.addStretch()
        btn("Chiudi", self.close)
        self.btn_ffind.clicked.connect(lambda: self.run("files_find"))
        self.btn_frepl.clicked.connect(lambda: self.run("files_replace"))

        top = QHBoxLayout()
        top.addLayout(left, 1)
        top.addLayout(right)
        self.status = QLabel(" ")
        self.status.setStyleSheet("font-weight:bold")
        main = QVBoxLayout(self)
        main.addLayout(top)
        main.addWidget(self.status)

        self._load_settings()
        self.files_inner.setVisible(self.files_box.isChecked())
        self.chk_dotall.setEnabled(self.rb_regex.isChecked())

    # ---- impostazioni -------------------------------------------------
    def _combo(self, key):
        cb = QComboBox()
        cb.setEditable(True)
        cb.setInsertPolicy(QComboBox.NoInsert)
        cb.setMinimumWidth(320)
        cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        hist = self.settings.value(key, []) or []
        if isinstance(hist, str):
            hist = [hist]
        cb.addItems(hist)
        cb.setCurrentText("")
        return cb

    def _load_settings(self):
        s = self.settings
        b = lambda k, d: s.value(k, d, type=bool)
        self.chk_word.setChecked(b("s_word", False))
        self.chk_case.setChecked(b("s_case", False))
        self.chk_wrap.setChecked(b("s_wrap", True))
        self.chk_dotall.setChecked(b("s_dotall", False))
        mode = s.value("s_mode", "normal")
        {"normal": self.rb_normal, "extended": self.rb_ext, "regex": self.rb_regex}.get(mode, self.rb_normal).setChecked(True)
        self.le_dir.setText(s.value("s_dir", ""))
        self.le_filter.setText(s.value("s_filter", "*.ftx *.txt *.md *.docx"))
        self.chk_sub.setChecked(b("s_sub", True))
        self.files_box.setChecked(b("s_files", False))

    def _save_settings(self, o):
        s = self.settings
        s.setValue("s_word", o.word); s.setValue("s_case", o.case); s.setValue("s_wrap", o.wrap)
        s.setValue("s_dotall", o.dotall); s.setValue("s_mode", o.mode)
        s.setValue("s_dir", self.le_dir.text()); s.setValue("s_filter", self.le_filter.text())
        s.setValue("s_sub", self.chk_sub.isChecked()); s.setValue("s_files", self.files_box.isChecked())
        for cb, key in ((self.cb_find, "hist_find"), (self.cb_repl, "hist_repl")):
            t = cb.currentText()
            if key == "hist_repl" and not t:
                continue
            if t:
                idx = cb.findText(t, Qt.MatchExactly | Qt.MatchCaseSensitive)
                if idx >= 0:
                    cb.removeItem(idx)
                cb.insertItem(0, t)
                cb.setCurrentIndex(0)
                while cb.count() > 20:
                    cb.removeItem(cb.count() - 1)
            s.setValue(key, [cb.itemText(i) for i in range(cb.count())])

    def opts(self):
        mode = "regex" if self.rb_regex.isChecked() else "extended" if self.rb_ext.isChecked() else "normal"
        return Opts(self.cb_find.currentText(), self.cb_repl.currentText(), mode, self.chk_case.isChecked(),
                    self.chk_word.isChecked(), self.chk_wrap.isChecked(), self.chk_dotall.isChecked(),
                    self.chk_insel.isChecked())

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "Cartella in cui cercare", self.le_dir.text())
        if d:
            self.le_dir.setText(d)

    # ---- apertura -----------------------------------------------------
    def show_page(self, page="find"):
        ed = self.win.editor
        sel = ed.textCursor().selectedText()
        if sel and "\u2029" not in sel and OBJ not in sel and len(sel) < 300:
            self.cb_find.setCurrentText(sel)
            self.chk_insel.setChecked(False)
        elif sel and "\u2029" in sel:
            self.chk_insel.setChecked(True)
        if page == "files":
            self.files_box.setChecked(True)
            if not self.le_dir.text() and self.win.current.path:
                self.le_dir.setText(os.path.dirname(self.win.current.path))
        self.show()
        self.raise_()
        self.activateWindow()
        target = self.cb_repl if page == "replace" else self.cb_find
        target.setFocus()
        target.lineEdit().selectAll()

    def run(self, action):
        o = self.opts()
        self._save_settings(o)
        self.win.last_opts = o
        try:
            msg = self.win.search.do(action, o, self)
        except ValueError as e:
            msg = str(e)
            QApplication.beep()
        self.status.setText(msg or " ")

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.close()
            return
        super().keyPressEvent(e)


# ---------------------------------------------------------------------------
#  Controllore: esegue le azioni sui tab / file
# ---------------------------------------------------------------------------

class SearchController:
    def __init__(self, win):
        self.win = win

    def _range(self, o, ed):
        c = ed.textCursor()
        if o.insel and c.hasSelection():
            return c.selectionStart(), c.selectionEnd()
        return None

    def do(self, action, o, dlg=None):
        win = self.win
        pat = compile_pattern(o)
        tab = win.current
        ed = tab.editor
        doc = ed.document()

        if action in ("next", "prev"):
            res = find_next(doc, pat, ed.textCursor(), action == "prev", o.wrap, self._range(o, ed))
            if not res:
                QApplication.beep()
                return f"Non trovato: \"{o.text}\""
            m, wrapped = res
            self.select(ed, m.start(), m.end())
            if wrapped:
                return "Trovato (ricominciato dall'" + ("inizio" if action == "next" else "fine") + ")"
            return ""

        if action == "count":
            r = self._range(o, ed) or (0, None)
            n = len(find_matches(doc, pat, *r))
            return f"{n} occorrenz{'a' if n == 1 else 'e'}"

        if action == "replace":
            c = ed.textCursor()
            if c.hasSelection():
                for m in find_matches(doc, pat, c.selectionStart(), c.selectionEnd()):
                    if m.start() == c.selectionStart() and m.end() == c.selectionEnd():
                        c.insertText(replacement(o, m), text_format(c))
                        ed.setTextCursor(c)
                        break
            res = find_next(doc, pat, ed.textCursor(), False, o.wrap)
            if res:
                self.select(ed, res[0].start(), res[0].end())
                return ""
            return "Nessun'altra occorrenza"

        if action == "replace_all":
            r = self._range(o, ed) or (0, None)
            n = replace_all_in_doc(doc, o, pat, *r)
            return f"{n} sostituzion{'e' if n == 1 else 'i'}"

        if action == "replace_tabs":
            tot, nd = 0, 0
            for t in win.all_tabs():
                n = replace_all_in_doc(t.editor.document(), o, pat)
                tot += n
                nd += 1 if n else 0
            return f"{tot} sostituzioni in {nd} document{'o' if nd == 1 else 'i'}"

        if action in ("all_doc", "all_tabs"):
            tabs = [tab] if action == "all_doc" else win.all_tabs()
            groups, tot = [], 0
            for t in tabs:
                d = t.editor.document()
                r = self._range(o, t.editor) if t is tab and action == "all_doc" else None
                ms = find_matches(d, pat, *(r or (0, None)))
                if ms:
                    groups.append((t.title(), t.path, t, hits_from_matches(d, ms)))
                    tot += len(ms)
            win.results.add_search(f"Cerca \"{o.text}\"  ({tot} risultati in {len(groups)} document{'o' if len(groups) == 1 else 'i'})", groups)
            return f"{tot} risultati"

        if action == "mark":
            r = self._range(o, ed) or (0, None)
            ms = find_matches(doc, pat, *r)
            ed.set_marks([(m.start(), m.end()) for m in ms])
            return f"{len(ms)} occorrenze evidenziate"

        if action in ("files_find", "files_replace"):
            return self.in_files(o, pat, dlg, action == "files_replace")
        return ""

    def select(self, ed, s, e):
        c = QTextCursor(ed.document())
        c.setPosition(s)
        c.setPosition(e, QTextCursor.KeepAnchor)
        ed.setTextCursor(c)
        ed.ensureCursorVisible()

    # ---- nei file -----------------------------------------------------
    def _list_files(self, folder, filters, sub):
        pats = [p for p in re.split(r"[;\s]+", filters) if p] or ["*"]
        out = []
        for root, dirs, files in os.walk(folder):
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            for f in sorted(files):
                if any(fnmatch.fnmatch(f.lower(), p.lower()) for p in pats):
                    out.append(os.path.join(root, f))
            if not sub:
                break
        return out

    def in_files(self, o, pat, dlg, replace):
        folder = dlg.le_dir.text().strip() if dlg else ""
        if not folder or not os.path.isdir(folder):
            raise ValueError("Scegli una cartella valida.")
        files = self._list_files(folder, dlg.le_filter.text(), dlg.chk_sub.isChecked())
        if not files:
            return "Nessun file corrisponde ai filtri."
        win = self.win
        if replace:
            editable = [f for f in files if f.lower().endswith((".ftx", ".txt"))]
            skipped = len(files) - len(editable)
            note = f"\n\n{skipped} file .docx/.md verranno saltati (non si modificano)." if skipped else ""
            r = QMessageBox.question(dlg, "Sostituisci nei file",
                                     f"Sostituire \"{o.text}\" con \"{o.repl}\" in {len(editable)} file "
                                     f".ftx/.txt di\n{folder}?\n\nI file non aperti vengono salvati subito "
                                     f"e l'operazione NON si può annullare.{note}")
            if r != QMessageBox.Yes:
                return "Annullato."
            files = editable
        prog = QProgressDialog("Ricerca nei file...", "Interrompi", 0, len(files), dlg)
        prog.setWindowModality(Qt.WindowModal)
        prog.setMinimumDuration(300)
        groups, tot, nfiles, errors = [], 0, 0, []
        for i, path in enumerate(files):
            prog.setValue(i)
            prog.setLabelText(os.path.basename(path))
            QApplication.processEvents()
            if prog.wasCanceled():
                break
            try:
                if os.path.getsize(path) > 50_000_000:
                    continue
                open_tab = win.tab_for_path(path)
                if open_tab is not None:
                    doc = open_tab.editor.document()
                    if replace:
                        n = replace_all_in_doc(doc, o, pat)
                        tot += n; nfiles += 1 if n else 0
                        continue
                    ms = find_matches(doc, pat)
                    if ms:
                        groups.append((path, path, open_tab, hits_from_matches(doc, ms)))
                        tot += len(ms)
                    continue
                store = ImageStore()
                model, kind = load_any(path, store)
                doc = build_headless(model, store)
                if replace:
                    n = replace_all_in_doc(doc, o, pat)
                    if n:
                        new = doc_to_model(doc, lambda nm: nm if nm in store.data else None)
                        if kind == "ftx":
                            save_ftx(path, new, store)
                        else:
                            with open(path, "w", encoding="utf-8", newline="\r\n") as fh:
                                fh.write(model_to_plain(new))
                        tot += n; nfiles += 1
                    continue
                ms = find_matches(doc, pat)
                if ms:
                    groups.append((path, path, None, hits_from_matches(doc, ms)))
                    tot += len(ms)
            except Exception as ex:
                errors.append(f"{os.path.basename(path)}: {ex}")
        prog.setValue(len(files))
        err = f" ({len(errors)} file non leggibili)" if errors else ""
        if replace:
            return f"{tot} sostituzioni in {nfiles} file{err}"
        win.results.add_search(f"Cerca \"{o.text}\" in {folder}  ({tot} risultati in {len(groups)} file su {len(files)}){err}", groups)
        return f"{tot} risultati in {len(groups)} file{err}"
