# -*- coding: utf-8 -*-
"""
FastText - pannello laterale: elenco dei file aperti e struttura (intestazioni) del documento.
"""

import os

from PySide6.QtCore import Qt, QTimer, QEvent
from PySide6.QtGui import QTextCursor, QColor, QBrush, QFont
from PySide6.QtWidgets import (
    QDockWidget, QLabel, QLineEdit, QSplitter, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
    QAbstractItemView, QHeaderView,
)

TAB_ROLE = Qt.UserRole
BLOCK_ROLE = Qt.UserRole + 1


def _section_label(text):
    lab = QLabel(text)
    lab.setStyleSheet("font-weight:bold; padding:4px 4px 2px 4px;")
    return lab


class SidePanel(QDockWidget):
    def __init__(self, win):
        super().__init__("File aperti", win)
        self.win = win
        self.setObjectName("SidePanel")
        self.setFeatures(QDockWidget.DockWidgetClosable | QDockWidget.DockWidgetMovable)

        # --- file aperti ---
        top = QWidget()
        tv = QVBoxLayout(top)
        tv.setContentsMargins(0, 0, 0, 0)
        tv.setSpacing(2)
        self.filter = QLineEdit()
        self.filter.setPlaceholderText("Filtra i file...")
        self.filter.setClearButtonEnabled(True)
        self.filter.textChanged.connect(self.refresh_files)
        self.files = QTreeWidget()
        self.files.setHeaderLabels(["Nome", "Cartella"])
        self.files.setRootIsDecorated(False)
        self.files.setUniformRowHeights(True)
        self.files.setSelectionMode(QAbstractItemView.SingleSelection)
        self.files.header().setSectionResizeMode(0, QHeaderView.Interactive)
        self.files.header().setStretchLastSection(True)
        self.files.setColumnWidth(0, 150)
        self.files.itemClicked.connect(self._file_clicked)
        self.files.itemActivated.connect(self._file_clicked)
        self.files.setContextMenuPolicy(Qt.CustomContextMenu)
        self.files.customContextMenuRequested.connect(self._file_menu)
        self.files.viewport().installEventFilter(self)
        tv.addWidget(self.filter)
        tv.addWidget(self.files)

        # --- struttura del documento ---
        bottom = QWidget()
        bv = QVBoxLayout(bottom)
        bv.setContentsMargins(0, 0, 0, 0)
        bv.setSpacing(2)
        self.outline = QTreeWidget()
        self.outline.setHeaderHidden(True)
        self.outline.setUniformRowHeights(True)
        self.outline.itemClicked.connect(self._outline_clicked)
        self.outline.itemActivated.connect(self._outline_clicked)
        bv.addWidget(_section_label("Struttura del documento"))
        bv.addWidget(self.outline)

        split = QSplitter(Qt.Vertical)
        split.addWidget(top)
        split.addWidget(bottom)
        split.setStretchFactor(0, 3)
        split.setStretchFactor(1, 2)
        split.setChildrenCollapsible(False)
        self.setWidget(split)

        self._outline_timer = QTimer(self)
        self._outline_timer.setSingleShot(True)
        self._outline_timer.setInterval(400)
        self._outline_timer.timeout.connect(self.refresh_outline)

    # ---- file aperti --------------------------------------------------
    def refresh_files(self):
        if not hasattr(self.win, "tabs"):
            return
        flt = self.filter.text().strip().lower()
        cur = self.win.tabs.currentWidget()
        tabs = self.win.all_tabs()
        self.setWindowTitle(f"File aperti ({len(tabs)})")
        self.files.blockSignals(True)
        self.files.clear()
        for t in tabs:
            folder = os.path.dirname(t.path) if t.path else "(non salvato)"
            it = QTreeWidgetItem([("● " if t.modified() else "") + t.title(), folder])
            it.setData(0, TAB_ROLE, t)
            it.setToolTip(0, t.path or t.title())
            it.setToolTip(1, t.path or "")
            it.setForeground(1, QBrush(QColor("#7a8794")))
            if t is cur:
                f = it.font(0)
                f.setBold(True)
                it.setFont(0, f)
            self.files.addTopLevelItem(it)
            if flt and flt not in t.title().lower() and flt not in (t.path or "").lower():
                it.setHidden(True)
            if t is cur:
                self.files.setCurrentItem(it)
        self.files.blockSignals(False)

    def _file_clicked(self, item, _col=0):
        t = item.data(0, TAB_ROLE)
        if t is not None and t in self.win.all_tabs():
            self.win.tabs.setCurrentWidget(t)
            self.win.editor.setFocus()

    def _file_menu(self, pos):
        item = self.files.itemAt(pos)
        if item is None:
            return
        t = item.data(0, TAB_ROLE)
        if t is not None:
            self.win.show_tab_menu(t, self.files.viewport().mapToGlobal(pos))

    def eventFilter(self, obj, e):
        if obj is self.files.viewport():
            if e.type() == QEvent.MouseButtonRelease and e.button() == Qt.MiddleButton:
                item = self.files.itemAt(e.position().toPoint())
                if item is not None:
                    self.win.close_tab(item.data(0, TAB_ROLE))
                    return True
            if e.type() == QEvent.MouseButtonDblClick and self.files.itemAt(e.position().toPoint()) is None:
                self.win.new_tab()
                return True
        return super().eventFilter(obj, e)

    # ---- struttura (intestazioni) -------------------------------------
    def schedule_outline(self):
        self._outline_timer.start()

    def refresh_outline(self):
        if not hasattr(self.win, "tabs") or self.win.tabs.currentWidget() is None:
            return
        doc = self.win.tabs.currentWidget().editor.document()
        self.outline.clear()
        parents = {0: None}
        b = doc.begin()
        count = 0
        while b.isValid():
            lvl = b.blockFormat().headingLevel()
            if lvl > 0:
                lvl = min(lvl, 3)
                text = b.text().replace("￼", "").strip() or "(intestazione vuota)"
                it = QTreeWidgetItem([text])
                it.setData(0, BLOCK_ROLE, b.blockNumber())
                if lvl == 1:
                    f = it.font(0)
                    f.setBold(True)
                    it.setFont(0, f)
                parent = None
                for k in range(lvl - 1, 0, -1):
                    if parents.get(k) is not None:
                        parent = parents[k]
                        break
                (parent.addChild(it) if parent else self.outline.addTopLevelItem(it))
                parents[lvl] = it
                for k in range(lvl + 1, 4):
                    parents[k] = None
                count += 1
            b = b.next()
        if not count:
            it = QTreeWidgetItem(["Nessuna intestazione. Usa Ctrl+1, Ctrl+2, Ctrl+3"])
            it.setFlags(Qt.NoItemFlags)
            self.outline.addTopLevelItem(it)
        self.outline.expandAll()

    def _outline_clicked(self, item, _col=0):
        n = item.data(0, BLOCK_ROLE)
        if n is None:
            return
        ed = self.win.editor
        b = ed.document().findBlockByNumber(n)
        if not b.isValid():
            return
        end = QTextCursor(ed.document())
        end.movePosition(QTextCursor.End)
        ed.setTextCursor(end)               # così l'intestazione finisce in cima alla finestra
        ed.setTextCursor(QTextCursor(b))
        ed.ensureCursorVisible()
        ed.setFocus()
