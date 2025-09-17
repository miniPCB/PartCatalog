#!/usr/bin/env python3
"""
Parts Catalog Tool — PyQt5 (Dark-only, Focused Editor v7)
- Single Save (toolbar + Ctrl+S) for files & folders
- Metadata tab has subtabs:
    • Introduction (Title, Part Number, Notes, etc.)
    • Revision History (global table)
    • Variant Details (list of items)
- Editor tabs: Metadata, Netlist (free text), Partlist (free text),
               Circuit Description (free text), Circuit Theory (free text),
               Design Analysis (free text), Review (raw Markdown)
- Autosave + dirty indicator:
    • Autosaves every N seconds *and* when switching tree selection
    • Live toolbar countdown: “Autosave in: XXs” (counts only when there are unsaved edits)
    • “●” on active tab + “•” in window title when unsaved
- Folder metadata panel fills the right pane when a folder is selected
- Drag & Drop tree with MOVE support (files and folders), resizable splitter
- Tree shows Name + Description (files: Title field from Markdown; folders: TITLE from folder JSON)
- Folder metadata JSON (<Folder>/<Folder>.json) with fields: TITLE, DESCRIPTION (UPPERCASE), Summary, Owner, Tags, Created, Last Updated
- Archive: zips folder containing this script; archive saved inside as YYYYMMDD_HHMMSS.zip
- App/window icon uses an emoji (🗂️)

Note: Global/cross-file search features were removed per request.
"""

import sys
import shutil
import os
import datetime
import json
import re
import tempfile
import subprocess
import platform
from pathlib import Path

from PyQt5.QtCore import Qt, QSortFilterProxyModel, QModelIndex, QTimer
from PyQt5.QtGui import QKeySequence, QIcon, QPixmap, QPainter, QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFileSystemModel, QTreeView, QToolBar, QAction, QFileDialog,
    QInputDialog, QMessageBox, QLabel, QAbstractItemView,
    QLineEdit, QHeaderView, QPushButton, QSpacerItem, QSizePolicy,
    QGroupBox, QSplitter, QTabWidget, QTextEdit, QTableWidget, QTableWidgetItem,
    QStyleFactory, QStyledItemDelegate, QFormLayout
)

APP_TITLE = "Parts Catalog Tool"
DEFAULT_CATALOG_DIR = Path.cwd() / "catalog"

# Metadata fields (Interface -> Notes)
FIELD_ORDER = [
    ("Title", "Click or tap here to enter text."),
    ("Part Number", "PN-XXX"),
    ("Notes", "General notes / interfaces / context"),
    ("Type of Design", "Specialized / Fundamental"),
    ("Level of Novelty", "Inventive / Common"),
    ("Performance Feedback", "Exists / Does Not Exist"),
    ("Verified Performance", "Yes / No"),
]

REV_HEADERS = ["Rev", "Date", "Description", "By"]
PARTLIST_HEADERS = ["Ref Des", "Type", "Value/Description"]
VARIANT_HEADERS = ["Item"]

SECTION_TITLES = {
    "cd": "Circuit Description",
    "ct": "Circuit Theory",
    "da": "Design Analysis",
}

# --- Helpers / constants ------------------------------------------------------

def today_iso():
    return datetime.date.today().isoformat()

def now_stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def folder_meta_path(folder: Path) -> Path:
    return folder / f"{folder.name}.json"

MD_ROW_RE = re.compile(r'^\|\s*(?P<field>[^|]+?)\s*\|\s*(?P<value>[^|]*?)\s*\|$')

def _divider_for(headers):
    return "| " + " | ".join("-" * len(h) for h in headers) + " |"

# Recognize markdown table divider cells like ---  :---  ---:
DIVIDER_CELL_RE = re.compile(r'^:?-{2,}:?$')

def _is_md_divider_line(line: str) -> bool:
    s = line.strip()
    if not (s.startswith("|") and s.endswith("|")):
        return False
    cells = [c.strip() for c in s.strip().strip("|").split("|")]
    if not cells:
        return False
    return all(DIVIDER_CELL_RE.fullmatch(c) for c in cells)

def _is_divider_cells(cells: list) -> bool:
    return bool(cells) and all(DIVIDER_CELL_RE.fullmatch((c or "").strip()) for c in cells)

NEW_ENTRY_TEMPLATE = f"""# Circuit Metadata

**Last Updated:** {today_iso()}

## Introduction

| Field                  | Value                     |
| ---------------------- | ------------------------- |
| Title                  | Click or tap here to enter text. |
| Part Number            | PN-XXX                    |
| Notes                  | General notes / interfaces / context |
| Type of Design         | Specialized / Fundamental |
| Level of Novelty       | Inventive / Common        |
| Performance Feedback   | Exists / Does Not Exist   |
| Verified Performance   | Yes / No                  |

## Revision History

| {' | '.join(REV_HEADERS)} |
{_divider_for(REV_HEADERS)}
| - | {today_iso()} | initial release |  |

## Variant Details

- (none)

## Netlist

(paste or type your netlist here)

## Partlist

| {" | ".join(PARTLIST_HEADERS)} |
{_divider_for(PARTLIST_HEADERS)}

## {SECTION_TITLES['cd']}

(brief textual description of the circuit, blocks, signals, constraints…)

## {SECTION_TITLES['ct']}

(theory of operation, formulas, transfer functions, stability, noise…)

## {SECTION_TITLES['da']}

(design tradeoffs, margins, component selection rationale, SOA, derating…)
""".strip() + "\n"

# --- Render an emoji into a QIcon --------------------------------------------

def make_emoji_icon(emoji: str, px: int = 256) -> QIcon:
    pm = QPixmap(px, px)
    pm.fill(Qt.transparent)
    painter = QPainter(pm)
    try:
        f = QFont("Segoe UI Emoji", int(px * 0.66))
        f.setStyleStrategy(QFont.PreferAntialias)
        painter.setFont(f)
        rect = pm.rect()
        painter.drawText(rect, Qt.AlignCenter, emoji)
    finally:
        painter.end()
    return QIcon(pm)

# --- Proxy model: filter + description + DnD forwarding -----------------------

class DescProxyModel(QSortFilterProxyModel):
    """
    Directories + .md files only.
    Column 0 = Name
    Column 1 = Description (file: Title field; folder: <folder>/<folder>.json 'TITLE')
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._desc_cache = {}

    def filterAcceptsRow(self, source_row, source_parent):
        sm = self.sourceModel()
        idx = sm.index(source_row, 0, source_parent)
        if not idx.isValid():
            return False
        if sm.isDir(idx):
            return True
        name = sm.fileName(idx)
        return name.lower().endswith(".md")

    def flags(self, index):
        base = super().flags(index)
        if not index.isValid():
            return Qt.ItemIsDropEnabled
        sidx = self.mapToSource(index)
        sm = self.sourceModel()
        if sm.isDir(sidx):
            return base | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled
        else:
            return (base | Qt.ItemIsDragEnabled) & ~Qt.ItemIsDropEnabled

    def supportedDropActions(self):
        return Qt.MoveAction

    def canDropMimeData(self, data, action, row, column, parent):
        if action != Qt.MoveAction:
            return False
        sp = self.mapToSource(parent)
        sm = self.sourceModel()
        return (not sp.isValid()) or sm.isDir(sp)

    def dropMimeData(self, data, action, row, column, parent):
        if action != Qt.MoveAction:
            return False
        smodel = self.sourceModel()
        src_parent = self.mapToSource(parent)
        if src_parent.isValid() and not smodel.isDir(src_parent):
            src_parent = src_parent.parent()
        return smodel.dropMimeData(data, action, row, column, src_parent)

    def columnCount(self, parent):
        return max(2, super().columnCount(parent))

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if index.column() == 0:
            return super().data(index, role)
        if index.column() == 1 and role in (Qt.DisplayRole, Qt.ToolTipRole):
            sidx = self.mapToSource(index.sibling(index.row(), 0))
            spath = Path(self.sourceModel().filePath(sidx))
            key = str(spath)
            cached = self._desc_cache.get(key)
            if cached is None:
                if spath.is_dir():
                    cached = self._read_folder_title(spath)
                elif spath.is_file() and spath.suffix.lower() == ".md":
                    cached = self._read_title_from_md(spath)
                else:
                    cached = ""
                self._desc_cache[key] = cached
            return cached
        if index.column() >= 2 and role == Qt.DisplayRole:
            return ""
        return super().data(index, role)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return ["Name", "Description"][section] if section in (0, 1) else super().headerData(section, orientation, role)
        return super().headerData(section, orientation, role)

    def _read_title_from_md(self, path: Path) -> str:
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return ""
        for line in text.splitlines():
            s = line.strip()
            if not s.startswith("|"):
                continue
            parts = [p.strip() for p in s.strip("|").split("|")]
            if len(parts) >= 2 and parts[0].lower() == "title":
                return parts[1]
        return ""

    def _read_folder_title(self, folder: Path) -> str:
        try:
            meta_p = folder_meta_path(folder)
            if not meta_p.exists():
                return ""
            with meta_p.open("r", encoding="utf-8") as f:
                meta = json.load(f)
            return meta.get("TITLE") or meta.get("title") or meta.get("description", "")
        except Exception:
            return ""

    def refresh_desc(self, path: Path):
        key = str(path)
        self._desc_cache.pop(key, None)
        sm = self.sourceModel()
        sidx = sm.index(str(path))
        if sidx.isValid():
            pidx = self.mapFromSource(sidx)
            if pidx.isValid():
                left = pidx.sibling(pidx.row(), 1)
                self.dataChanged.emit(left, left, [Qt.DisplayRole, Qt.ToolTipRole])

# --- Windows dark titlebar helper --------------------------------------------
if platform.system() == "Windows":
    import ctypes
    from ctypes import wintypes

    def _set_win_dark_titlebar(hwnd: int, enabled: bool = True):
        try:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19
            attribute = wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE)
            pv = ctypes.c_int(1 if enabled else 0)
            dwm = ctypes.WinDLL("dwmapi")
            res = dwm.DwmSetWindowAttribute(
                wintypes.HWND(hwnd),
                attribute,
                ctypes.byref(pv),
                ctypes.sizeof(pv)
            )
            if res != 0:
                attribute = wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1)
                dwm.DwmSetWindowAttribute(
                    wintypes.HWND(hwnd),
                    attribute,
                    ctypes.byref(pv),
                    ctypes.sizeof(pv)
                )
        except Exception:
            pass

    def apply_windows_dark_titlebar(widget):
        try:
            hwnd = int(widget.winId())
            _set_win_dark_titlebar(hwnd, True)
        except Exception:
            pass
else:
    def apply_windows_dark_titlebar(widget):
        pass

# --- Custom delegate to make inline editor slim/legible -----------------------

class SlimLineEditDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setFrame(False)
        editor.setStyleSheet(
            "QLineEdit {"
            "  background-color: #2A2D31;"
            "  color: #E6E6E6;"
            "  border: 1px solid #3A3F44;"
            "  padding: 2px 4px;"
            "  selection-background-color: #3B4252;"
            "}"
        )
        return editor

# --- Main Window --------------------------------------------------------------

class CatalogWindow(QMainWindow):
    def __init__(self, catalog_root: Path, app_icon: QIcon):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(app_icon)
        self.resize(1280, 900)

        self.catalog_root = catalog_root
        self.current_path: Path | None = None
        self.current_folder: Path | None = None

        # Dirty / autosave
        self.dirty = False
        self.suppress_dirty = False
        self.review_dirty = False

        self.autosave_interval_ms = 30000  # 30s
        self.autosave_interval_s = self.autosave_interval_ms // 1000
        self.autosave_remaining_s = self.autosave_interval_s

        self.base_tab_titles = []  # will be set after tabs are built

        # Toolbar
        tb = QToolBar("Main", self)
        tb.setMovable(False)
        self.addToolBar(tb)

        act_new_entry = QAction("🧩 New Entry", self);  act_new_entry.triggered.connect(self.create_new_entry);  tb.addAction(act_new_entry)
        act_new_folder = QAction("🗂️ New Folder", self); act_new_folder.triggered.connect(self.create_new_folder); tb.addAction(act_new_folder)
        act_rename     = QAction("✏️ Rename", self);     act_rename.triggered.connect(self.rename_item);           tb.addAction(act_rename)
        act_delete     = QAction("🗑️ Delete", self);     act_delete.triggered.connect(self.delete_item);           tb.addAction(act_delete)

        tb.addSeparator()

        act_archive = QAction("📦 Archive", self)
        act_archive.setToolTip("Zip the folder containing this script; save the zip inside that folder as YYYYMMDD_HHMMSS.zip")
        act_archive.triggered.connect(self.archive_script_folder)
        tb.addAction(act_archive)

        act_open_loc = QAction("📂 Open Location", self)
        act_open_loc.setToolTip("Open the selected folder (or the folder containing the selected file) in your file manager")
        act_open_loc.triggered.connect(self.open_file_location)
        tb.addAction(act_open_loc)

        tb.addSeparator()

        self.act_save = QAction("💾 Save (Ctrl+S)", self)
        self.act_save.setShortcut(QKeySequence.Save)
        self.act_save.triggered.connect(self.save_from_form)
        tb.addAction(self.act_save)

        # Right-aligned autosave countdown label
        spacer = QWidget(self); spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred); tb.addWidget(spacer)
        self.autosave_label = QLabel("Autosave in: —", self)
        self.autosave_label.setStyleSheet("color:#A0A6AD; padding: 0 6px;")
        tb.addWidget(self.autosave_label)

        # File system model / proxy
        self.fs_model = QFileSystemModel(self)
        self.fs_model.setReadOnly(False)
        self.fs_model.setRootPath(str(self.catalog_root))
        self.fs_model.setNameFilters(["*.md"])
        self.fs_model.setNameFilterDisables(False)
        self.fs_model.fileRenamed.connect(self.on_fs_file_renamed)

        self.proxy = DescProxyModel(self)
        self.proxy.setSourceModel(self.fs_model)

        # Left: Tree
        self.tree = QTreeView(self)
        self.tree.setModel(self.proxy)
        self.tree.setItemDelegate(SlimLineEditDelegate(self.tree))
        self.tree.setRootIndex(self.proxy.mapFromSource(self.fs_model.index(str(self.catalog_root))))
        self.tree.setHeaderHidden(False)
        self.tree.setSortingEnabled(True)
        self.tree.sortByColumn(0, Qt.AscendingOrder)
        for col in range(2, self.proxy.columnCount(self.tree.rootIndex())):
            self.tree.setColumnHidden(col, True)
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)

        # Drag & Drop (move)
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDropIndicatorShown(True)
        self.tree.setDefaultDropAction(Qt.MoveAction)
        self.tree.setDragDropMode(QAbstractItemView.DragDrop)
        self.tree.setDragDropOverwriteMode(False)

        self.tree.selectionModel().selectionChanged.connect(self.on_tree_selection)

        # Right container (path label + folder panel + tabs)
        right_container = QWidget(self); right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0,0,0,0); right_layout.setSpacing(8)

        self.path_label = QLabel("", self)
        right_layout.addWidget(self.path_label)

        # Folder metadata panel
        self.folder_panel = QGroupBox("Folder Metadata", self)
        fp_layout = QFormLayout(self.folder_panel)
        self.folder_title = QLineEdit(self.folder_panel); self.folder_title.setPlaceholderText("TITLE (UPPERCASE)")
        self.folder_desc  = QLineEdit(self.folder_panel); self.folder_desc.setPlaceholderText("DESCRIPTION (UPPERCASE)")
        self.folder_summary = QTextEdit(self.folder_panel); self.folder_summary.setPlaceholderText("Summary / notes for this folder…")
        self.folder_owner = QLineEdit(self.folder_panel); self.folder_owner.setPlaceholderText("Owner name or team")
        self.folder_tags  = QLineEdit(self.folder_panel); self.folder_tags.setPlaceholderText("Comma-separated tags (e.g., power, digital, hv)")
        self.folder_created = QLineEdit(self.folder_panel); self.folder_created.setReadOnly(True)
        self.folder_updated = QLineEdit(self.folder_panel); self.folder_updated.setReadOnly(True)
        fp_layout.addRow("TITLE:", self.folder_title)
        fp_layout.addRow("DESCRIPTION:", self.folder_desc)
        fp_layout.addRow("Summary:", self.folder_summary)
        fp_layout.addRow("Owner:", self.folder_owner)
        fp_layout.addRow("Tags:", self.folder_tags)
        fp_layout.addRow("Created:", self.folder_created)
        fp_layout.addRow("Last Updated:", self.folder_updated)
        right_layout.addWidget(self.folder_panel, 1)

        # Editor tabs
        self.tabs = QTabWidget(self)

        # -------- Metadata tab with subtabs: Introduction / Revision History / Variant Details
        meta_tab = QWidget(self); meta_v = QVBoxLayout(meta_tab); meta_v.setContentsMargins(0,0,0,0); meta_v.setSpacing(8)

        self.meta_inner = QTabWidget(meta_tab)

        # Introduction
        intro_tab = QWidget(self); intro_v = QVBoxLayout(intro_tab); intro_v.setContentsMargins(0,0,0,0)
        self.fields_group = QGroupBox("Introduction", intro_tab)
        self.fields_form = QFormLayout(self.fields_group)
        self.field_widgets: dict[str, QLineEdit] = {}
        for label, placeholder in FIELD_ORDER:
            le = QLineEdit(self.fields_group); le.setPlaceholderText(placeholder)
            self.field_widgets[label] = le
            self.fields_form.addRow(label + ":", le)
        intro_v.addWidget(self.fields_group)
        self.meta_inner.addTab(intro_tab, "Introduction")

        # Revision History
        rev_tab = QWidget(self); rev_v = QVBoxLayout(rev_tab)
        self.rev_table = QTableWidget(0, len(REV_HEADERS), rev_tab)
        self.rev_table.setHorizontalHeaderLabels(REV_HEADERS)
        for c, h in enumerate(REV_HEADERS):
            mode = QHeaderView.Stretch if h == "Description" else QHeaderView.ResizeToContents
            self.rev_table.horizontalHeader().setSectionResizeMode(c, mode)
        self.rev_table.verticalHeader().setVisible(False)
        rev_buttons = QHBoxLayout()
        self.btn_rev_add = QPushButton("Add Row"); self.btn_rev_del = QPushButton("Remove Selected")
        self.btn_rev_add.clicked.connect(self.add_rev_row)
        self.btn_rev_del.clicked.connect(self.remove_rev_row)
        rev_buttons.addWidget(self.btn_rev_add); rev_buttons.addWidget(self.btn_rev_del)
        rev_buttons.addItem(QSpacerItem(40,20,QSizePolicy.Expanding,QSizePolicy.Minimum))
        rev_v.addWidget(self.rev_table)
        rev_v.addLayout(rev_buttons)
        self.meta_inner.addTab(rev_tab, "Revision History")

        # Variant Details (list)
        var_tab = QWidget(self); var_v = QVBoxLayout(var_tab)
        self.variant_table = QTableWidget(0, 1, var_tab)
        self.variant_table.setHorizontalHeaderLabels(VARIANT_HEADERS)
        self.variant_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.variant_table.verticalHeader().setVisible(False)
        var_buttons = QHBoxLayout()
        self.btn_var_add = QPushButton("Add Item"); self.btn_var_del = QPushButton("Remove Selected")
        self.btn_var_add.clicked.connect(self.add_variant_row)
        self.btn_var_del.clicked.connect(self.remove_variant_row)
        var_buttons.addWidget(self.btn_var_add); var_buttons.addWidget(self.btn_var_del)
        var_buttons.addItem(QSpacerItem(40,20,QSizePolicy.Expanding,QSizePolicy.Minimum))
        var_v.addWidget(self.variant_table)
        var_v.addLayout(var_buttons)
        self.meta_inner.addTab(var_tab, "Variant Details")

        meta_v.addWidget(self.meta_inner)
        self.tabs.addTab(meta_tab, "Metadata")

        # Netlist (free text)
        net_tab = QWidget(self); net_v = QVBoxLayout(net_tab)
        self.netlist_edit = QTextEdit(net_tab)
        self.netlist_edit.setPlaceholderText("(paste or type your netlist here)")
        net_v.addWidget(self.netlist_edit)
        self.tabs.addTab(net_tab, "Netlist")

        # Partlist (free text)
        pl_tab = QWidget(self); pl_v = QVBoxLayout(pl_tab)
        self.partlist_edit = QTextEdit(pl_tab)
        self.partlist_edit.setPlaceholderText(
            "(paste or type your partlist table here)\n"
            f"| {' | '.join(PARTLIST_HEADERS)} |\n{_divider_for(PARTLIST_HEADERS)}"
        )
        pl_v.addWidget(self.partlist_edit)
        self.tabs.addTab(pl_tab, "Partlist")

        # Circuit Description / Theory / Design Analysis (each simple text editor)
        self.cd_edit = QTextEdit(self); self.cd_edit.setPlaceholderText("(Click or type to add content…)")
        cd_tab = QWidget(self); cd_v = QVBoxLayout(cd_tab); cd_v.addWidget(self.cd_edit)
        self.tabs.addTab(cd_tab, SECTION_TITLES["cd"])

        self.ct_edit = QTextEdit(self); self.ct_edit.setPlaceholderText("(Click or type to add content…)")
        ct_tab = QWidget(self); ct_v = QVBoxLayout(ct_tab); ct_v.addWidget(self.ct_edit)
        self.tabs.addTab(ct_tab, SECTION_TITLES["ct"])

        self.da_edit = QTextEdit(self); self.da_edit.setPlaceholderText("(Click or type to add content…)")
        da_tab = QWidget(self); da_v = QVBoxLayout(da_tab); da_v.addWidget(self.da_edit)
        self.tabs.addTab(da_tab, SECTION_TITLES["da"])

        # Review (raw markdown)
        review_tab = QWidget(self); review_v = QVBoxLayout(review_tab)
        self.review_edit = QTextEdit(review_tab)
        self.review_edit.setPlaceholderText("Raw Markdown view. Edits here will be saved verbatim.")
        self.review_edit.textChanged.connect(self._on_review_changed)
        review_v.addWidget(self.review_edit)
        self.tabs.addTab(review_tab, "Review")

        right_layout.addWidget(self.tabs, 1)

        # After building tabs: set base titles, wire dirty + countdown autosave
        self.base_tab_titles = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        self.tabs.currentChanged.connect(self.update_dirty_indicator)

        # Countdown timer (1s ticks) handles autosave + label update
        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self._tick_autosave_countdown)
        self.countdown_timer.start(1000)

        self._wire_dirty_signals()

        # Splitter
        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(self.tree)
        splitter.addWidget(right_container)
        splitter.setStretchFactor(0, 1); splitter.setStretchFactor(1, 2); splitter.setSizes([360, 980])

        central = QWidget(self); outer = QHBoxLayout(central); outer.setContentsMargins(8,8,8,8); outer.setSpacing(8)
        outer.addWidget(splitter); self.setCentralWidget(central)

        # Dark stylesheet
        self.apply_dark_styles()
        self.show_file_ui(False)

    # ---------- Import / Export -------------------------------------------------

    def import_catalog(self):
        dlg = QFileDialog(self, "Import Catalog")
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setNameFilters(["Markdown (*.md)", "All Files (*)"])
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self._apply_dark(dlg)
        if dlg.exec_() != QFileDialog.Accepted:
            return
        src_path = Path(dlg.selectedFiles()[0])

        try:
            text = src_path.read_text(encoding="utf-8")
        except Exception as e:
            self.error("Import Catalog", f"Failed to read file:\n{e}")
            return

        target = self.catalog_root
        if not self.ask_yes_no("Import Target", f"Import into this folder?\n{target}\n\nChoose 'No' to pick a different folder."):
            picked = self.select_directory("Select Import Target Folder")
            if not picked:
                return
            target = picked

        overwrite_all = self.ask_yes_no("Overwrite Policy", "If files/folders already exist, overwrite them?")
        lines = text.splitlines()
        i = 0
        n = len(lines)
        current_rel = None
        errors = []

        def ensure_dir(rel_path: str) -> Path:
            p = target / ("" if rel_path == "." else rel_path)
            p.mkdir(parents=True, exist_ok=True)
            return p

        while i < n:
            ln = lines[i].strip()
            if ln.startswith("## Folder:"):
                current_rel = ln.split(":", 1)[1].strip()
                ensure_dir(current_rel)
                i += 1
                continue

            if current_rel:
                if ln == "```json meta":
                    block = []
                    i += 1
                    while i < n and lines[i].strip() != "```":
                        block.append(lines[i])
                        i += 1
                    if i < n and lines[i].strip() == "```":
                        i += 1
                    folder_p = ensure_dir(current_rel)
                    meta_p = folder_meta_path(folder_p)
                    try:
                        meta = json.loads("\n".join(block))
                    except Exception:
                        meta = {"TITLE": "", "DESCRIPTION": "", "Summary": "", "Owner": "", "Tags": [], "Created": today_iso(), "Last Updated": today_iso()}
                    if isinstance(meta.get("Tags"), str):
                        meta["Tags"] = [t.strip() for t in meta["Tags"].split(",") if t.strip()]
                    try:
                        if meta_p.exists() and not overwrite_all:
                            pass
                        else:
                            meta_p.write_text(json.dumps(meta, indent=2), encoding="utf-8")
                    except Exception as e:
                        errors.append(f"{meta_p}: {e}")
                    continue

                if ln.startswith("```markdown"):
                    m = re.search(r'entry\s*=\s*"([^"]+)"', ln)
                    entry_name = m.group(1) if m else None
                    block = []
                    i += 1
                    while i < n and lines[i].strip() != "```":
                        block.append(lines[i])
                        i += 1
                    if i < n and lines[i].strip() == "```":
                        i += 1
                    if entry_name:
                        folder_p = ensure_dir(current_rel)
                        file_p = folder_p / entry_name
                        try:
                            if file_p.exists() and not overwrite_all:
                                pass
                            else:
                                file_p.write_text("\n".join(block), encoding="utf-8")
                        except Exception as e:
                            errors.append(f"{file_p}: {e}")
                    continue

            i += 1

        self.fs_model.refresh(self.fs_model.index(str(self.catalog_root)))
        if errors:
            self.warn("Import Catalog", "Import completed with some errors:\n" + "\n".join(errors[:20]) + ("\n..." if len(errors) > 20 else ""))
        else:
            self.info("Import Catalog", "Import completed successfully.")

    def export_catalog(self):
        root = self.catalog_root.resolve()
        dlg = QFileDialog(self, "Save Catalog Export")
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["Markdown (*.md)"])
        dlg.selectFile(f"catalog_export_{now_stamp()}.md")
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self._apply_dark(dlg)
        if dlg.exec_() != QFileDialog.Accepted:
            return
        out_path = Path(dlg.selectedFiles()[0])
        if out_path.suffix.lower() != ".md":
            out_path = out_path.with_suffix(".md")

        lines = []
        lines.append("# Parts Catalog Export")
        lines.append("Version: 1")
        lines.append(f"Exported: {datetime.datetime.now().isoformat(timespec='seconds')}")
        lines.append("Root: .")
        lines.append("")

        for folder, dirnames, filenames in os.walk(root):
            folder_p = Path(folder)
            rel = "." if folder_p == root else str(folder_p.relative_to(root)).replace("\\", "/")

            lines.append(f"## Folder: {rel}")
            lines.append("")

            meta_p = folder_meta_path(folder_p)
            if meta_p.exists():
                try:
                    meta = json.loads(meta_p.read_text(encoding="utf-8"))
                except Exception:
                    meta = {"TITLE": "", "DESCRIPTION": "", "Summary": "", "Owner": "", "Tags": [], "Created": today_iso(), "Last Updated": today_iso()}
            else:
                meta = {"TITLE": "", "DESCRIPTION": "", "Summary": "", "Owner": "", "Tags": [], "Created": today_iso(), "Last Updated": today_iso()}

            lines.append("```json meta")
            try:
                lines.append(json.dumps(meta, indent=2))
            except Exception:
                lines.append(json.dumps({"_error": "could not encode meta as json"}, indent=2))
            lines.append("```")
            lines.append("")

            for name in sorted(fn for fn in filenames if fn.lower().endswith(".md")):
                fpath = folder_p / name
                try:
                    content = fpath.read_text(encoding="utf-8")
                except Exception as e:
                    content = f"<!-- ERROR READING FILE: {e} -->\n"
                safe_name = name.replace('"', '\\"')
                lines.append(f"```markdown entry=\"{safe_name}\"")
                lines.append(content.rstrip("\n"))
                lines.append("```")
                lines.append("")

        try:
            out_path.write_text("\n".join(lines), encoding="utf-8")
        except Exception as e:
            self.error("Export Catalog", f"Failed to write export file:\n{e}")
            return

        self.info("Export Catalog", f"Exported to:\n{out_path}")

    # ---------- Tree rename handling -------------------------------------------

    def on_fs_file_renamed(self, dir_path_str: str, old_name: str, new_name: str):
        try:
            dir_path = Path(dir_path_str)
            old_path = dir_path / old_name
            new_path = dir_path / new_name

            if new_path.is_dir():
                old_meta = new_path / f"{old_name}.json"
                new_meta = new_path / f"{new_name}.json"

                if old_meta.exists():
                    if new_meta.exists():
                        try: old_meta.unlink()
                        except Exception: pass
                    else:
                        try: old_meta.rename(new_meta)
                        except Exception:
                            try:
                                data = old_meta.read_text(encoding="utf-8")
                                new_meta.write_text(data, encoding="utf-8")
                                old_meta.unlink()
                            except Exception:
                                pass

                self.proxy.refresh_desc(new_path)

                if self.current_folder and (self.current_folder == old_path or self.current_folder.name == old_name and self.current_folder.parent == dir_path):
                    self.current_folder = new_path
                    self.path_label.setText(f"Folder: {new_path}")

            else:
                if new_path.suffix.lower() == ".md":
                    self.proxy.refresh_desc(new_path)
                if self.current_path and (self.current_path == old_path or (self.current_path.name == old_name and self.current_path.parent == dir_path)):
                    self.current_path = new_path
                    self.path_label.setText(f"File: {new_path}")

        except Exception:
            pass

    # --- Styling ----------------------------------------------------------------
    def apply_dark_styles(self):
        self.setStyleSheet("""
            QWidget { background-color: #202225; color: #E6E6E6; }
            QToolBar { background: #1B1E20; spacing: 6px; border: 0px; }
            QToolButton { color: #E6E6E6; }
            QLabel { color: #E6E6E6; }
            QGroupBox {
                border: 1px solid #3A3F44; border-radius: 6px; margin-top: 12px; padding-top: 8px;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #CFCFCF; }
            QLineEdit, QTextEdit {
                background-color: #2A2D31; color: #E6E6E6; border: 1px solid #3A3F44; border-radius: 6px; padding: 6px;
            }
            QTreeView QLineEdit {
                background-color: #2A2D31; color: #E6E6E6; border: 1px solid #3A3F44; padding: 2px 4px;
            }
            QTreeView {
                background-color: #1E2124; alternate-background-color: #24272B; border: 1px solid #3A3F44;
            }
            QTreeView::item:selected { background: #3B4252; color: #E6E6E6; }
            QHeaderView::section {
                background-color: #2A2D31; color: #E6E6E6; border: 0px; padding: 6px; font-weight: 600;
            }
            QTabBar::tab {
                background: #2A2D31; color: #E6E6E6; padding: 8px 12px; margin-right: 2px;
                border-top-left-radius: 6px; border-top-right-radius: 6px;
            }
            QTabBar::tab:selected { background: #3A3F44; }
            QTabBar::tab:hover { background: #34383D; }

            QMessageBox, QInputDialog, QFileDialog { background-color: #202225; color: #E6E6E6; }
            QMessageBox QPushButton, QInputDialog QPushButton, QFileDialog QPushButton {
                background-color: #2F343A; color: #E6E6E6; border: 1px solid #444;
                border-radius: 6px; padding: 6px 12px;
            }
            QMessageBox QPushButton:hover, QInputDialog QPushButton:hover, QFileDialog QPushButton:hover { background-color: #3A4047; }
        """)

    # --- Dialog helpers ---------------------------------------------------------
    def _apply_dark(self, dlg):
        try:
            apply_windows_dark_titlebar(dlg)
        except Exception:
            pass

    def select_directory(self, title: str) -> Path | None:
        dlg = QFileDialog(self, title)
        dlg.setFileMode(QFileDialog.Directory)
        dlg.setOption(QFileDialog.ShowDirsOnly, True)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        self._apply_dark(dlg)
        if dlg.exec_() == QFileDialog.Accepted:
            return Path(dlg.selectedFiles()[0])
        return None

    def ask_text(self, title: str, label: str, default: str = "") -> tuple[str, bool]:
        dlg = QInputDialog(self)
        dlg.setWindowTitle(title)
        dlg.setLabelText(label)
        dlg.setTextValue(default)
        self._apply_dark(dlg)
        ok = dlg.exec_() == dlg.Accepted
        return (dlg.textValue(), ok)

    def ask_yes_no(self, title: str, text: str) -> bool:
        mb = QMessageBox(self)
        mb.setWindowTitle(title)
        mb.setText(text)
        mb.setIcon(QMessageBox.Question)
        mb.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        self._apply_dark(mb)
        return mb.exec_() == QMessageBox.Yes

    def info(self, title: str, text: str):
        mb = QMessageBox(self)
        mb.setWindowTitle(title)
        mb.setText(text)
        mb.setIcon(QMessageBox.Information)
        mb.setStandardButtons(QMessageBox.Ok)
        self._apply_dark(mb)
        mb.exec_()

    def warn(self, title: str, text: str):
        mb = QMessageBox(self)
        mb.setWindowTitle(title)
        mb.setText(text)
        mb.setIcon(QMessageBox.Warning)
        self._apply_dark(mb)
        mb.setStandardButtons(QMessageBox.Ok)
        mb.exec_()

    def error(self, title: str, text: str):
        mb = QMessageBox(self)
        mb.setWindowTitle(title)
        mb.setText(text)
        mb.setIcon(QMessageBox.Critical)
        mb.setStandardButtons(QMessageBox.Ok)
        self._apply_dark(mb)
        mb.exec_()

    # --- Dirty / autosave helpers ----------------------------------------------
    def _wire_dirty_signals(self):
        # Text editors
        for ed in [self.netlist_edit, self.partlist_edit, self.cd_edit, self.ct_edit, self.da_edit]:
            ed.textChanged.connect(self._mark_dirty)
        # Review editor (also has review_dirty flag)
        self.review_edit.textChanged.connect(self._mark_dirty)
        # Introduction fields
        for le in self.field_widgets.values():
            le.textChanged.connect(self._mark_dirty)
        # Metadata tables
        self.rev_table.itemChanged.connect(self._mark_dirty)
        self.variant_table.itemChanged.connect(self._mark_dirty)
        # Folder panel fields
        self.folder_title.textChanged.connect(self._mark_dirty)
        self.folder_desc.textChanged.connect(self._mark_dirty)
        self.folder_summary.textChanged.connect(self._mark_dirty)
        self.folder_owner.textChanged.connect(self._mark_dirty)
        self.folder_tags.textChanged.connect(self._mark_dirty)

    def _strip_dot(self, s: str) -> str:
        return s.lstrip("● ").strip()

    def _mark_dirty(self, *args):
        if self.suppress_dirty:
            return
        if not self.dirty:
            self.dirty = True
            # start/refresh countdown when first becoming dirty
            self.autosave_remaining_s = self.autosave_interval_s
            self.autosave_label.setText(f"Autosave in: {self.autosave_remaining_s}s")
            self.update_dirty_indicator()

    def update_dirty_indicator(self):
        # Ensure base titles match current tabs
        if len(self.base_tab_titles) != self.tabs.count():
            self.base_tab_titles = [self._strip_dot(self.tabs.tabText(i)) for i in range(self.tabs.count())]
        for i, base in enumerate(self.base_tab_titles):
            want = ("● " + base) if (self.dirty and i == self.tabs.currentIndex()) else base
            if self.tabs.tabText(i) != want:
                self.tabs.setTabText(i, want)
        self.setWindowTitle(APP_TITLE + (" •" if self.dirty else ""))

    def _clear_dirty(self):
        if self.dirty:
            self.dirty = False
            self.update_dirty_indicator()

    def _reset_autosave_countdown(self):
        self.autosave_remaining_s = self.autosave_interval_s
        self.autosave_label.setText("Autosave in: —" if not self.dirty else f"Autosave in: {self.autosave_remaining_s}s")

    def _tick_autosave_countdown(self):
        # Only count when there's something to save
        if not self.current_path and not self.current_folder:
            self.autosave_label.setText("Autosave in: —")
            self.autosave_remaining_s = self.autosave_interval_s
            return
        if not self.dirty:
            # Idle / nothing to save
            self.autosave_label.setText("Autosave in: —")
            self.autosave_remaining_s = self.autosave_interval_s
            return

        # Countdown visible
        if self.autosave_remaining_s <= 0:
            self.save_from_form(silent=True)  # this clears dirty if save succeeds
            self._reset_autosave_countdown()
        else:
            self.autosave_label.setText(f"Autosave in: {self.autosave_remaining_s}s")
            self.autosave_remaining_s -= 1

    # --- UI helpers -------------------------------------------------------------
    def _on_review_changed(self):
        self.review_dirty = True
        self._mark_dirty()

    def show_file_ui(self, file_selected: bool):
        self.tabs.setVisible(file_selected)
        self.folder_panel.setVisible(not file_selected)

    def selected_source_index(self) -> QModelIndex | None:
        sel = self.tree.selectionModel().selectedIndexes()
        if not sel: return None
        idx = sel[0]
        if idx.column() != 0:
            idx = self.proxy.index(idx.row(), 0, idx.parent())
        return self.proxy.mapToSource(idx)

    def selected_path(self) -> Path | None:
        sidx = self.selected_source_index()
        if not sidx or not sidx.isValid(): return None
        return Path(self.fs_model.filePath(sidx))

    def is_markdown(self, path: Path) -> bool:
        return path.is_file() and path.suffix.lower() == ".md"

    # --- Metadata helpers -------------------------------------------------------
    def add_rev_row(self):
        r = self.rev_table.rowCount()
        self.rev_table.insertRow(r)
        defaults = ["", today_iso(), "", ""]
        if r == 0:
            defaults = ["-", today_iso(), "initial release", ""]
        for c, val in enumerate(defaults):
            self.rev_table.setItem(r, c, QTableWidgetItem(val))

    def remove_rev_row(self):
        r = self.rev_table.currentRow()
        if r >= 0:
            self.rev_table.removeRow(r)

    def add_variant_row(self):
        r = self.variant_table.rowCount()
        self.variant_table.insertRow(r)
        self.variant_table.setItem(r, 0, QTableWidgetItem(""))

    def remove_variant_row(self):
        r = self.variant_table.currentRow()
        if r >= 0:
            self.variant_table.removeRow(r)

    # --- Selection / load -------------------------------------------------------
    def on_tree_selection(self, *_):
        # Autosave before switching away
        if self.dirty:
            self.save_from_form(silent=True)
            self._reset_autosave_countdown()

        path = self.selected_path()
        if not path: return
        if path.is_dir():
            self.current_path = None
            self.current_folder = path
            self.path_label.setText(f"Folder: {path}")
            self.load_folder_meta(path)
            self.show_file_ui(False)
            return
        if self.is_markdown(path):
            self.current_folder = None
            self.load_file(path)
            self.show_file_ui(True)

    def load_folder_meta(self, folder: Path):
        self.suppress_dirty = True
        meta = self.read_folder_meta(folder)
        self.folder_title.setText(meta.get("TITLE", ""))
        self.folder_desc.setText(meta.get("DESCRIPTION", ""))
        self.folder_summary.setPlainText(meta.get("Summary", ""))
        self.folder_owner.setText(meta.get("Owner", ""))
        tags_val = meta.get("Tags", "")
        if isinstance(tags_val, list):
            tags_val = ", ".join(tags_val)
        self.folder_tags.setText(tags_val)
        self.folder_created.setText(meta.get("Created", ""))
        self.folder_updated.setText(meta.get("Last Updated", ""))
        self.suppress_dirty = False
        self._clear_dirty()
        self._reset_autosave_countdown()

    def load_file(self, path: Path):
        try:
            text = path.read_text(encoding="utf-8")
        except Exception as e:
            self.error("Error", f"Failed to read file:\n{e}"); return

        self.current_path = path
        self.path_label.setText(f"File: {path}")

        (fields, rev_rows, variant_items, netlist, partlist_text,
         cd, ct, da) = self.parse_markdown(text)

        self.suppress_dirty = True

        # Intro fields
        for key, _ in FIELD_ORDER:
            self.field_widgets[key].setText(fields.get(key, ""))

        # Revision table
        self.rev_table.setRowCount(0)
        for row in rev_rows:
            rr = (row + [""] * len(REV_HEADERS))[:len(REV_HEADERS)]
            r = self.rev_table.rowCount()
            self.rev_table.insertRow(r)
            for c, val in enumerate(rr):
                self.rev_table.setItem(r, c, QTableWidgetItem(val))

        # Variant Details
        self.variant_table.setRowCount(0)
        for item in variant_items:
            r = self.variant_table.rowCount()
            self.variant_table.insertRow(r)
            self.variant_table.setItem(r, 0, QTableWidgetItem(item))

        # Sections
        self.netlist_edit.setPlainText(netlist)
        self.partlist_edit.setPlainText(partlist_text)
        self.cd_edit.setPlainText(cd)
        self.ct_edit.setPlainText(ct)
        self.da_edit.setPlainText(da)

        # Raw review
        self.review_edit.blockSignals(True)
        self.review_edit.setPlainText(text)
        self.review_edit.blockSignals(False)
        self.review_dirty = False

        self.suppress_dirty = False
        self._clear_dirty()
        self._reset_autosave_countdown()

        self.proxy.refresh_desc(path)

    # --- Parse / Build markdown -------------------------------------------------
    def _find_section(self, lines, title: str):
        want = f"## {title}".lower()
        for i, ln in enumerate(lines):
            if ln.strip().lower() == want:
                return i
        return None

    def _read_section_text(self, lines, start_idx: int) -> str:
        j = start_idx + 1
        out = []
        n = len(lines)
        while j < n and not lines[j].startswith("## "):
            out.append(lines[j])
            j += 1
        while out and not out[0].strip(): out.pop(0)
        while out and not out[-1].strip(): out.pop()
        return "\n".join(out)

    def _parse_table_at(self, lines, start_idx: int):
        """Parse a markdown table starting at/after start_idx (skips any divider rows)."""
        rows = []
        if start_idx is None:
            return rows
        j = start_idx + 1
        n = len(lines)

        # find header row
        while j < n and not lines[j].strip().startswith("|"):
            if lines[j].strip().startswith("## "):
                return rows
            j += 1
        if j >= n:
            return rows

        # skip one or more divider lines following the header
        j += 1
        while j < n and lines[j].strip().startswith("|") and _is_md_divider_line(lines[j]):
            j += 1

        # consume body (skip any stray divider lines in body)
        while j < n and lines[j].strip().startswith("|"):
            if _is_md_divider_line(lines[j]):
                j += 1
                continue
            raw = lines[j].strip().strip("|")
            cells = [c.strip() for c in raw.split("|")]
            rows.append(cells)
            j += 1
        return rows

    def _parse_bulleted_list(self, lines, start_idx: int):
        items = []
        if start_idx is None: return items
        j = start_idx + 1
        n = len(lines)
        while j < n and not lines[j].startswith("## "):
            s = lines[j].strip()
            if s.startswith("- "):
                items.append(s[2:].strip())
            elif s.startswith("* "):
                items.append(s[2:].strip())
            j += 1
        items = [it for it in items if it.lower() != "(none)"]
        return items

    def parse_markdown(self, text: str):
        lines = [ln.rstrip("\n") for ln in text.splitlines()]
        fields = {k: "" for k, _ in FIELD_ORDER}
        rev_rows = []
        variant_items = []
        netlist = ""
        partlist_text = ""
        cd = ct = da = ""

        # Introduction table
        i = 0; n = len(lines)
        # collect all fields (for optional migration if you add later)
        all_fields = {}
        while i < n:
            if lines[i].strip().lower().startswith("| field") and "| value" in lines[i].lower():
                i += 2
                while i < n and lines[i].strip().startswith("|"):
                    m = MD_ROW_RE.match(lines[i].strip())
                    if m:
                        field = m.group("field").strip()
                        value = m.group("value").strip()
                        all_fields[field] = value
                        if field in fields: fields[field] = value
                    i += 1
                break
            i += 1

        # Global Revision History
        rhx = self._find_section(lines, "Revision History")
        if rhx is not None:
            rev_rows = self._parse_table_at(lines, rhx)

        # Variant Details
        vdx = self._find_section(lines, "Variant Details")
        if vdx is not None:
            variant_items = self._parse_bulleted_list(lines, vdx)

        # Netlist
        nix = self._find_section(lines, "Netlist")
        if nix is not None:
            netlist = self._read_section_text(lines, nix)

        # Partlist
        pix = self._find_section(lines, "Partlist")
        if pix is not None:
            tmp = self._read_section_text(lines, pix)
            partlist_text = tmp if tmp.strip() else f"| {' | '.join(PARTLIST_HEADERS)} |\n{_divider_for(PARTLIST_HEADERS)}"

        # Narrative sections
        cdx = self._find_section(lines, SECTION_TITLES["cd"])
        if cdx is not None: cd = self._read_section_text(lines, cdx)
        ctx = self._find_section(lines, SECTION_TITLES["ct"])
        if ctx is not None: ct = self._read_section_text(lines, ctx)
        dax = self._find_section(lines, SECTION_TITLES["da"])
        if dax is not None: da = self._read_section_text(lines, dax)

        return fields, rev_rows, variant_items, netlist, partlist_text, cd, ct, da

    def build_markdown(self, fields: dict, rev_rows: list, variant_items: list,
                       netlist: str, partlist_text: str,
                       cd: str, ct: str, da: str) -> str:
        # Introduction table
        intro_lines = [
            "| Field                  | Value                     |",
            "| ---------------------- | ------------------------- |",
        ]
        for key, _ in FIELD_ORDER:
            intro_lines.append(f"| {key:<22} | {fields.get(key,'').strip()} |")

        # Build table helper (skip divider-like rows)
        def build_table(headers, rows):
            out = ["| " + " | ".join(headers) + " |",
                   _divider_for(headers)]
            for r in rows:
                if _is_divider_cells([x.strip() for x in r]):
                    continue
                rr = (r + [""] * len(headers))[:len(headers)]
                out.append("| " + " | ".join(rr) + " |")
            return "\n".join(out)

        # Variant list to markdown
        def build_variant_list(items):
            items = [it.strip() for it in items if it.strip()]
            if not items:
                return "- (none)"
            return "\n".join(f"- {it}" for it in items)

        # Partlist default if empty
        ptxt = (partlist_text or "").strip()
        if not ptxt:
            ptxt = f"| {' | '.join(PARTLIST_HEADERS)} |\n{_divider_for(PARTLIST_HEADERS)}"

        def block(h, body):
            body = (body or "").strip()
            return [f"## {h}", "", (body if body else "(Click or type to add content…)")]

        out = [
            "# Circuit Metadata",
            "",
            f"**Last Updated:** {today_iso()}",
            "",
            "## Introduction", "",
            "\n".join(intro_lines), "",
            "## Revision History", "",
            build_table(REV_HEADERS, rev_rows), "",
            "## Variant Details", "",
            build_variant_list(variant_items), "",
            *block("Netlist", netlist), "",
            "## Partlist", "",
            ptxt, "",
            *block(SECTION_TITLES["cd"], cd), "",
            *block(SECTION_TITLES["ct"], ct), "",
            *block(SECTION_TITLES["da"], da), "",
        ]
        return "\n".join(out).rstrip() + "\n"

    # --- Save -------------------------------------------------------------------
    def save_from_form(self, silent: bool = False):
        """
        If a file is selected -> save Markdown.
        If a folder is selected -> save folder JSON.
        """
        # Save file
        if self.current_path and self.is_markdown(self.current_path):
            # Review tab takes precedence when dirty or active
            if self.review_dirty or self._strip_dot(self.tabs.tabText(self.tabs.currentIndex())) == "Review":
                raw = self.review_edit.toPlainText()
                try:
                    self.current_path.write_text(raw, encoding="utf-8")
                except Exception as e:
                    self.error("Error", f"Failed to save file:\n{e}"); return
                if not silent:
                    self.load_file(self.current_path)  # refresh all views for manual saves
                self.review_dirty = False
                self._clear_dirty()
                self._reset_autosave_countdown()
                # if not silent:
                    # self.info("Saved", "Catalog entry saved from raw Markdown.")
                return

            # Save from structured tabs
            fields = {k: w.text().strip() for k, w in self.field_widgets.items()}

            # Revision rows
            rev_rows = []
            for r in range(self.rev_table.rowCount()):
                row = []
                for c in range(len(REV_HEADERS)):
                    it = self.rev_table.item(r, c)
                    row.append(it.text().strip() if it else "")
                rev_rows.append(row)

            # Variant items
            variant_items = []
            for r in range(self.variant_table.rowCount()):
                it = self.variant_table.item(r, 0)
                txt = it.text().strip() if it else ""
                if txt:
                    variant_items.append(txt)

            netlist = self.netlist_edit.toPlainText()
            partlist_text = self.partlist_edit.toPlainText()
            cd = self.cd_edit.toPlainText()
            ct = self.ct_edit.toPlainText()
            da = self.da_edit.toPlainText()

            text = self.build_markdown(
                fields, rev_rows, variant_items,
                netlist, partlist_text, cd, ct, da
            )

            try:
                self.current_path.write_text(text, encoding="utf-8")
            except Exception as e:
                self.error("Error", f"Failed to save file:\n{e}"); return

            # Update raw view silently
            self.review_edit.blockSignals(True)
            self.review_edit.setPlainText(text)
            self.review_edit.blockSignals(False)
            self.review_dirty = False
            self._clear_dirty()
            self._reset_autosave_countdown()

            self.proxy.refresh_desc(self.current_path)
            # if not silent:
                # self.info("Saved", "Catalog entry saved.")
            return

        # Save folder metadata
        if self.current_folder and self.current_folder.exists() and self.current_folder.is_dir():
            meta_p = folder_meta_path(self.current_folder)
            created = today_iso()
            if meta_p.exists():
                try:
                    old = json.loads(meta_p.read_text(encoding="utf-8"))
                    created = old.get("Created", created)
                except Exception:
                    pass

            raw_tags = (self.folder_tags.text() or "").strip()
            tags_list = [t.strip() for t in raw_tags.split(",") if t.strip()]

            meta = ({
                "TITLE": (self.folder_title.text() or "").upper(),
                "DESCRIPTION": (self.folder_desc.text() or "").upper(),
                "Summary": self.folder_summary.toPlainText().strip(),
                "Owner": (self.folder_owner.text() or "").strip(),
                "Tags": tags_list,
                "Created": created,
                "Last Updated": today_iso(),
            })
            try:
                meta_p.write_text(json.dumps(meta, indent=2), encoding="utf-8")
            except Exception as e:
                self.error("Error", f"Failed to save folder metadata:\n{e}"); return

            self.proxy.refresh_desc(self.current_folder)
            self.folder_created.setText(meta["Created"])
            self.folder_updated.setText(meta["Last Updated"])
            self._clear_dirty()
            self._reset_autosave_countdown()
            if not silent:
                self.info("Saved", "Folder metadata saved.")
            return

        if not silent:
            self.info("Save", "Select a folder or a Markdown file to save.")

    # --- Folder metadata / archive / file ops ----------------------------------
    def read_folder_meta(self, folder: Path) -> dict:
        meta_p = folder_meta_path(folder)
        if not meta_p.exists():
            meta = {
                "TITLE": "",
                "DESCRIPTION": "",
                "Summary": "",
                "Owner": "",
                "Tags": [],
                "Created": today_iso(),
                "Last Updated": today_iso(),
            }
            try: meta_p.write_text(json.dumps(meta, indent=2), encoding="utf-8")
            except Exception: pass
            return meta
        try:
            meta = json.loads(meta_p.read_text(encoding="utf-8"))
            meta.setdefault("TITLE", meta.get("title", "").upper() if meta.get("title") else meta.get("TITLE",""))
            meta.setdefault("DESCRIPTION", meta.get("description", "").upper() if meta.get("description") else meta.get("DESCRIPTION",""))
            meta.setdefault("Summary", meta.get("Summary", ""))
            meta.setdefault("Owner", meta.get("Owner", ""))
            if "Tags" in meta and isinstance(meta["Tags"], str):
                meta["Tags"] = [t.strip() for t in meta["Tags"].split(",") if t.strip()]
            meta.setdefault("Tags", meta.get("Tags", []))
            meta.setdefault("Created", meta.get("Created", today_iso()))
            meta.setdefault("Last Updated", meta.get("Last Updated", today_iso()))
            return meta
        except Exception:
            return {
                "TITLE": "",
                "DESCRIPTION": "",
                "Summary": "",
                "Owner": "",
                "Tags": [],
                "Created": today_iso(),
                "Last Updated": today_iso(),
            }

    def archive_script_folder(self):
        try:
            if getattr(sys, 'frozen', False):
                script_dir = Path(sys.executable).resolve().parent
            else:
                script_dir = Path(__file__).resolve().parent
        except Exception:
            self.error("Archive", "Could not determine script directory."); return

        ts = now_stamp()
        temp_base = Path(tempfile.gettempdir()) / ts
        try:
            shutil.make_archive(str(temp_base), 'zip', root_dir=str(script_dir.parent), base_dir=script_dir.name)
        except Exception as e:
            self.error("Archive", f"Failed to create archive:\n{e}"); return

        temp_zip = Path(str(temp_base) + ".zip")
        if not temp_zip.exists():
            self.error("Archive", "Archive creation failed (file missing)."); return

        dest_zip = script_dir / f"{ts}.zip"
        try:
            if dest_zip.exists():
                dest_zip = script_dir / f"{ts}_1.zip"
            shutil.move(str(temp_zip), str(dest_zip))
        except Exception as e:
            self.error("Archive", f"Failed to move archive into folder:\n{e}"); return

        self.info("Archive", f"Created: {dest_zip}")

    def open_file_location(self):
        path = self.selected_path()
        if not path:
            self.info("Open Location", "Select a folder or file first.")
            return

        try:
            if platform.system() == "Windows":
                if path.is_file():
                    subprocess.run(["explorer", "/select,", str(path.resolve())])
                else:
                    subprocess.run(["explorer", str(path.resolve())])
            elif platform.system() == "Darwin":
                if path.is_file():
                    subprocess.run(["open", "-R", str(path.resolve())])
                else:
                    subprocess.run(["open", str(path.resolve())])
            else:
                target = str(path.parent.resolve() if path.is_file() else path.resolve())
                subprocess.run(["xdg-open", target])
        except Exception as e:
            self.error("Open Location", f"Failed to open location:\n{e}")

    def create_new_folder(self):
        base = self.selected_path() or self.catalog_root
        if base.is_file(): base = base.parent
        name, ok = self.ask_text("New Folder", "Folder name:")
        if not ok or not name.strip(): return
        target = base / name.strip()
        try:
            target.mkdir(parents=True, exist_ok=False)
            self.read_folder_meta(target)
            self.proxy.refresh_desc(target)
        except FileExistsError:
            self.warn("Exists", "A file/folder with that name already exists.")
        except Exception as e:
            self.error("Error", f"Failed to create folder:\n{e}")

    def create_new_entry(self):
        base = self.selected_path() or self.catalog_root
        if base.is_file(): base = base.parent
        name, ok = self.ask_text("New Entry", "File name (without extension):")
        if not ok or not name.strip(): return
        safe = name.strip()
        if not safe.lower().endswith(".md"): safe += ".md"
        target = base / safe
        if target.exists():
            self.warn("Exists", "A file with that name already exists."); return
        try:
            target.write_text(NEW_ENTRY_TEMPLATE, encoding="utf-8")
        except Exception as e:
            self.error("Error", f"Failed to create file:\n{e}"); return
        self.load_file(target)
        sidx = self.fs_model.index(str(target))
        if sidx.isValid():
            pidx = self.proxy.mapFromSource(sidx)
            if pidx.isValid(): self.tree.setCurrentIndex(pidx)

    def rename_item(self):
        path = self.selected_path()
        if not path:
            self.info("Rename", "Select a file or folder to rename."); return

        new_name, ok = self.ask_text("Rename", "New name:", default=path.name)
        if not ok or not new_name.strip():
            return

        new_path = path.parent / new_name.strip()
        if new_path.exists():
            self.warn("Exists", "Target name already exists."); return

        try:
            if path.is_dir():
                old_folder_name = path.name
                path.rename(new_path)

                old_meta = new_path / f"{old_folder_name}.json"
                new_meta = new_path / f"{new_path.name}.json"

                if old_meta.exists():
                    if new_meta.exists():
                        try: old_meta.unlink()
                        except Exception: pass
                    else:
                        try: old_meta.rename(new_meta)
                        except Exception:
                            try:
                                data = old_meta.read_text(encoding="utf-8")
                                new_meta.write_text(data, encoding="utf-8")
                                old_meta.unlink()
                            except Exception:
                                pass

                self.proxy.refresh_desc(new_path)

                if self.current_folder and self.current_folder == path:
                    self.current_folder = new_path
                    self.path_label.setText(f"Folder: {new_path}")

            else:
                path.rename(new_path)
                self.proxy.refresh_desc(new_path)
                if self.current_path and self.current_path == path:
                    self.current_path = new_path
                    self.path_label.setText(f"File: {new_path}")

        except Exception as e:
            self.error("Error", f"Failed to rename:\n{e}")

    def delete_item(self):
        path = self.selected_path()
        if not path: return
        typ = "folder" if path.is_dir() else "file"
        if not self.ask_yes_no("Delete", f"Delete this {typ}?\n{path}"):
            return
        try:
            if path.is_dir(): shutil.rmtree(path)
            else: path.unlink()
        except Exception as e:
            self.error("Error", f"Failed to delete:\n{e}"); return
        if self.current_path and self.current_path == path:
            self.current_path = None
            self.path_label.setText("")
        if self.current_folder and self.current_folder == path:
            self.current_folder = None
            self.path_label.setText("")

# --- Boot ---------------------------------------------------------------------

def ensure_catalog_root(start_dir: Path | None = None) -> Path:
    root = DEFAULT_CATALOG_DIR if start_dir is None else start_dir
    if not root.exists():
        try:
            root.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
    if not root.exists() or not root.is_dir():
        dlg = QFileDialog()
        dlg.setFileMode(QFileDialog.Directory)
        dlg.setOption(QFileDialog.ShowDirsOnly, True)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setWindowTitle("Select Catalog Root Folder")
        apply_windows_dark_titlebar(dlg)
        if dlg.exec_():
            sel = dlg.selectedFiles()
            if sel:
                root = Path(sel[0])
                root.mkdir(parents=True, exist_ok=True)
    return root

def main():
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))

    icon = make_emoji_icon("🗂️", px=256)
    app.setWindowIcon(icon)

    root = ensure_catalog_root()
    win = CatalogWindow(root, icon)
    win.show()
    apply_windows_dark_titlebar(win)
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
