# -*- coding: utf-8 -*-
"""
FastText - operazioni sul testo stile Notepad++:
maiuscole/minuscole e operazioni sulle righe (ordina, duplica, sposta, elimina, ...).
Le immagini nelle righe vengono conservate.
"""

import random
import re

from PySide6.QtGui import QTextCursor, QTextCharFormat

from ftcore import OBJ, text_format

SPECIAL = OBJ + "  ﷐﷑"


# ---------------------------------------------------------------------------
#  Maiuscole / minuscole
# ---------------------------------------------------------------------------

def _up(ch):
    u = ch.upper()
    return u if len(u) == 1 else ch


def _lo(ch):
    l = ch.lower()
    return l if len(l) == 1 else ch


def transform_case(text, mode):
    out = []
    if mode == "upper":
        return "".join(_up(c) for c in text)
    if mode == "lower":
        return "".join(_lo(c) for c in text)
    if mode == "invert":
        return "".join(_lo(c) if c.isupper() else _up(c) for c in text)
    if mode == "random":
        return "".join(random.choice((_up, _lo))(c) for c in text)
    if mode in ("proper", "proper_blend"):
        in_word = False
        for c in text:
            if c.isalnum():
                if not in_word:
                    out.append(_up(c))
                else:
                    out.append(_lo(c) if mode == "proper" else c)
                in_word = True
            else:
                out.append(c)
                in_word = False
        return "".join(out)
    if mode in ("sentence", "sentence_blend"):
        start = True
        for c in text:
            if c.isalpha():
                out.append(_up(c) if start else (_lo(c) if mode == "sentence" else c))
                start = False
            else:
                out.append(c)
                if c in ".!?\n  ﷐﷑":
                    start = True
        return "".join(out)
    return text


def convert_case(editor, mode):
    c = editor.textCursor()
    if not c.hasSelection():
        c.select(QTextCursor.WordUnderCursor)
    s, e = c.selectionStart(), c.selectionEnd()
    if s == e:
        return
    doc = editor.document()
    old = doc.toRawText()[s:e]
    new = transform_case(old, mode)
    if new == old:
        return
    cur = QTextCursor(doc)
    cur.beginEditBlock()
    n = len(old)
    # sostituisce solo i tratti cambiati, dalla fine all'inizio (immagini intatte)
    runs = []
    k = 0
    while k < n:
        if old[k] != new[k] and old[k] not in SPECIAL:
            j = k
            while j < n and old[j] != new[j] and old[j] not in SPECIAL:
                j += 1
            runs.append((k, j))
            k = j
        else:
            k += 1
    for a, b in reversed(runs):
        cur.setPosition(s + a)
        cur.setPosition(s + b, QTextCursor.KeepAnchor)
        cur.insertText(new[a:b], text_format(cur))
    cur.endEditBlock()
    sel = QTextCursor(doc)
    sel.setPosition(s)
    sel.setPosition(e, QTextCursor.KeepAnchor)
    editor.setTextCursor(sel)


# ---------------------------------------------------------------------------
#  Righe
# ---------------------------------------------------------------------------

def _container(block):
    c = QTextCursor(block)
    t = c.currentTable()
    if t is None:
        return None
    cell = t.cellAt(c)
    return (t.firstPosition(), cell.row(), cell.column())


def _same(a, b):
    return a.isValid() and b.isValid() and _container(a) == _container(b)


def selected_blocks(editor, whole_if_no_sel=False):
    """Righe selezionate (tutte nello stesso 'contenitore': testo normale o una cella).
    Senza selezione: la riga corrente, oppure tutto il tratto di testo se whole_if_no_sel."""
    c = editor.textCursor()
    doc = editor.document()
    b1 = doc.findBlock(c.selectionStart())
    b2 = doc.findBlock(c.selectionEnd())
    if c.hasSelection() and b2.blockNumber() > b1.blockNumber() and c.selectionEnd() == b2.position():
        b2 = b2.previous()
    if not c.hasSelection() and whole_if_no_sel:
        while _same(b1, b1.previous()):
            b1 = b1.previous()
        while _same(b2, b2.next()):
            b2 = b2.next()
    cont = _container(b1)
    blocks, b = [], b1
    while True:
        if not b.isValid() or _container(b) != cont:
            return None
        blocks.append(b)
        if b.blockNumber() >= b2.blockNumber():
            break
        b = b.next()
    return blocks


def _frag(block):
    """Contenuto di una riga + il suo formato (intestazione, linea...)."""
    c = QTextCursor(block)
    c.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
    return (c.selection(), block.blockFormat(), block.charFormat())


def _rewrite(editor, blocks, frags, select_count=None):
    """Sostituisce le righe 'blocks' con i frammenti 'frags' (una riga ciascuno)."""
    doc = editor.document()
    s = blocks[0].position()
    e = blocks[-1].position() + blocks[-1].length() - 1
    c = QTextCursor(doc)
    c.beginEditBlock()
    c.setPosition(s)
    c.setPosition(e, QTextCursor.KeepAnchor)
    c.removeSelectedText()
    ends = []
    for i, (f, bf, cf) in enumerate(frags):
        if i:
            c.insertBlock(bf, cf)
        else:
            c.setBlockFormat(bf)
            c.setBlockCharFormat(cf)
        if not f.isEmpty():
            c.insertFragment(f)
        ends.append(c.position())
    c.endEditBlock()
    sel = QTextCursor(doc)
    sel.setPosition(s)
    if frags:
        k = len(frags) if select_count is None else select_count
        sel.setPosition(ends[k - 1] if k else s, QTextCursor.KeepAnchor)
    editor.setTextCursor(sel)


def _natural_key(text):
    t = text.replace(OBJ, "").casefold()
    return [(0, int(p), "") if p.isdigit() else (1, 0, p) for p in re.split(r"(\d+)", t) if p != ""]


class LineOpError(Exception):
    pass


def _need(editor, whole=False):
    blocks = selected_blocks(editor, whole)
    if not blocks:
        raise LineOpError("Seleziona righe che stanno tutte fuori da una tabella oppure nella stessa cella.")
    return blocks


def sort_lines(editor, reverse=False):
    blocks = _need(editor, True)
    order = sorted(blocks, key=lambda b: _natural_key(b.text()), reverse=reverse)
    _rewrite(editor, blocks, [_frag(b) for b in order])


def reverse_lines(editor):
    blocks = _need(editor, True)
    _rewrite(editor, blocks, [_frag(b) for b in reversed(blocks)])


def remove_duplicates(editor):
    blocks = _need(editor, True)
    seen, keep = set(), []
    for b in blocks:
        if b.text() not in seen:
            seen.add(b.text())
            keep.append(b)
    if len(keep) != len(blocks):
        _rewrite(editor, blocks, [_frag(b) for b in keep])
    return len(blocks) - len(keep)


def remove_empty(editor):
    blocks = _need(editor, True)
    keep = [b for b in blocks if b.text().strip()]
    if not keep:
        keep = blocks[:1]
    if len(keep) != len(blocks):
        _rewrite(editor, blocks, [_frag(b) for b in keep])
    return len(blocks) - len(keep)


def trim(editor, where="trailing"):
    blocks = _need(editor, True)
    doc = editor.document()
    c = QTextCursor(doc)
    c.beginEditBlock()
    for b in reversed(blocks):
        t = b.text()
        if where in ("trailing", "both"):
            n = len(t) - len(t.rstrip(" \t "))
            if n:
                c.setPosition(b.position() + len(t) - n)
                c.setPosition(b.position() + len(t), QTextCursor.KeepAnchor)
                c.removeSelectedText()
        if where in ("leading", "both"):
            n = len(t) - len(t.lstrip(" \t "))
            if n:
                c.setPosition(b.position())
                c.setPosition(b.position() + n, QTextCursor.KeepAnchor)
                c.removeSelectedText()
    c.endEditBlock()


def join_lines(editor):
    blocks = _need(editor)
    if len(blocks) < 2:
        nxt = blocks[0].next()
        if not _same(blocks[0], nxt):
            return
        blocks.append(nxt)
    doc = editor.document()
    c = QTextCursor(doc)
    c.beginEditBlock()
    for b in reversed(blocks[1:]):
        t = b.text()
        lead = len(t) - len(t.lstrip(" \t"))
        prev = b.previous().text()
        c.setPosition(b.position() - 1)
        c.setPosition(b.position() + lead, QTextCursor.KeepAnchor)
        c.insertText("" if (not prev or prev.endswith((" ", "\t")) or not t.strip()) else " ", text_format(c))
    c.endEditBlock()


def duplicate_lines(editor):
    blocks = _need(editor)
    doc = editor.document()
    c = QTextCursor(doc)
    c.beginEditBlock()
    c.setPosition(blocks[-1].position() + blocks[-1].length() - 1)
    frags = [_frag(b) for b in blocks]
    for f, bf, cf in frags:
        c.insertBlock(bf, cf)
        if not f.isEmpty():
            c.insertFragment(f)
    c.endEditBlock()


def delete_lines(editor):
    blocks = _need(editor)
    doc = editor.document()
    first, last = blocks[0], blocks[-1]
    c = QTextCursor(doc)
    nxt, prv = last.next(), first.previous()
    if _same(last, nxt):
        c.setPosition(first.position())
        c.setPosition(nxt.position(), QTextCursor.KeepAnchor)
    elif _same(first, prv):
        c.setPosition(prv.position() + prv.length() - 1)
        c.setPosition(last.position() + last.length() - 1, QTextCursor.KeepAnchor)
    else:
        c.setPosition(first.position())
        c.setPosition(last.position() + last.length() - 1, QTextCursor.KeepAnchor)
    c.removeSelectedText()
    editor.setTextCursor(c)


def move_lines(editor, up=True):
    blocks = _need(editor)
    if up:
        nb = blocks[0].previous()
        if not _same(blocks[0], nb):
            return
        _rewrite(editor, [nb] + blocks, [_frag(b) for b in blocks] + [_frag(nb)], len(blocks))
    else:
        nb = blocks[-1].next()
        if not _same(blocks[-1], nb):
            return
        order = [nb] + blocks
        doc = editor.document()
        _rewrite(editor, blocks + [nb], [_frag(b) for b in order])
        # seleziona le righe spostate (ora dalla seconda in poi)
        first = doc.findBlock(editor.textCursor().selectionStart()).next()
        last = first
        for _ in range(len(blocks) - 1):
            last = last.next()
        sel = QTextCursor(doc)
        sel.setPosition(first.position())
        sel.setPosition(last.position() + last.length() - 1, QTextCursor.KeepAnchor)
        editor.setTextCursor(sel)


def tabs_to_spaces(editor, width=4):
    blocks = _need(editor, True)
    c = QTextCursor(editor.document())
    c.beginEditBlock()
    for b in reversed(blocks):
        t = b.text()
        for i in reversed([k for k, ch in enumerate(t) if ch == "\t"]):
            c.setPosition(b.position() + i)
            c.setPosition(b.position() + i + 1, QTextCursor.KeepAnchor)
            c.insertText(" " * width, text_format(c))
    c.endEditBlock()


def text_stats(doc):
    raw = doc.toRawText()
    plain = raw.replace(OBJ, "")
    text = re.sub(r"[  ﷐﷑]", "\n", plain)
    words = len(re.findall(r"\w+", text))
    chars = len(re.sub(r"\n", "", text))
    chars_ns = len(re.sub(r"\s", "", text))
    return {"righe": doc.blockCount(), "parole": words, "caratteri": chars, "caratteri_senza_spazi": chars_ns,
            "immagini": raw.count(OBJ)}
