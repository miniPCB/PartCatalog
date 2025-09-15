#!/usr/bin/env python3
"""
Parts Catalog Tool — PyQt5 (Dark-only, Deeper Editor)
- One Save (toolbar + Ctrl+S) for both Files and Folders
- Editor tabs: Metadata, Used On, Netlist, Partlist, Pin Interface, Tests
- Folder metadata panel fills the right pane when a folder is selected
- Drag & Drop tree with MOVE support (files and folders), resizable splitter
- Tree shows Name + Description (for files: Title field from Markdown; for folders: TITLE from folder JSON)
- Folder metadata JSON (<Folder>/<Folder>.json) with fields: TITLE, DESCRIPTION (UPPERCASE), Summary, Owner, Tags, Created, Last Updated
- Archive: zips the folder containing this script; archive created in temp, then moved INSIDE that folder as YYYYMMDD_HHMMSS.zip
- App/window icon uses an emoji (🗂️)
"""

import sys
import shutil
import datetime
import json
import re
import tempfile
import subprocess
import platform
from pathlib import Path

from PyQt5.QtCore import Qt, QSortFilterProxyModel, QModelIndex
from PyQt5.QtGui import QKeySequence, QIcon, QPixmap, QPainter, QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFileSystemModel, QTreeView, QToolBar, QAction, QFileDialog,
    QInputDialog, QMessageBox, QLabel, QAbstractItemView, QFormLayout,
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QSpacerItem, QSizePolicy, QSplitter, QTabWidget, QTextEdit,
    QStyleFactory, QStyledItemDelegate
)

APP_TITLE = "Parts Catalog Tool"
DEFAULT_CATALOG_DIR = Path.cwd() / "catalog"

FIELD_ORDER = [
    ("Title", "Click or tap here to enter text."),
    ("Part Number", "PN-XXX"),
    ("Revision", "A1-00"),
    ("Interface", "I2C / SPI / UART / Analog / Digital"),
    ("Type of Design", "Specialized / Fundamental"),
    ("Level of Novelty", "Inventive / Common"),
    ("Performance Feedback", "Exists / Does Not Exist"),
    ("Verified Performance", "Yes / No"),
]

# --- Helpers / constants ------------------------------------------------------

def today_iso():
    return datetime.date.today().isoformat()

def now_stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def folder_meta_path(folder: Path) -> Path:
    return folder / f"{folder.name}.json"

MD_ROW_RE = re.compile(r'^\|\s*(?P<field>[^|]+?)\s*\|\s*(?P<value>[^|]*?)\s*\|$')

# === Updated headers per your spec ===========================================
PARTLIST_HEADERS = ["Ref Des", "Type", "Value/Description"]
PIN_HEADERS = ["Pin", "Name", "Description", "Note"]
TEST_HEADERS = ["Test No.", "Name", "Description", "Note"]

# --- New-entry template with updated tables -----------------------------------

def _divider_for(headers):
    # Divider row matching header text length for nice alignment
    return "| " + " | ".join("-" * len(h) for h in headers) + " |"

NEW_ENTRY_TEMPLATE = f"""# Circuit Metadata

**Last Updated:** {today_iso()}

| Field                  | Value                     |
| ---------------------- | ------------------------- |
| Title                  | Click or tap here to enter text. |
| Part Number            | PN-XXX                    |
| Revision               | A1-00                     |
| Interface              | I2C / SPI / UART / Analog / Digital |
| Type of Design         | Specialized / Fundamental |
| Level of Novelty       | Inventive / Common        |
| Performance Feedback   | Exists / Does Not Exist   |
| Verified Performance   | Yes / No                  |

## Used On

| PN         | Occurrences |
| ---------- | ----------- |
| (None)     | 0           |

## Netlist

(paste or type your netlist here)

## Partlist

| {" | ".join(PARTLIST_HEADERS)} |
{_divider_for(PARTLIST_HEADERS)}

## Pin Interface

| {" | ".join(PIN_HEADERS)} |
{_divider_for(PIN_HEADERS)}

## Tests

| {" | ".join(TEST_HEADERS)} |
{_divider_for(TEST_HEADERS)}
""".strip() + "\n"

# --- Render an emoji into a QIcon --------------------------------------------

def make_emoji_icon(emoji: str, px: int = 256) -> QIcon:
    """Render an emoji to a square pixmap and return as QIcon."""
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
        """Allow dragging everywhere; allow dropping only onto directories."""
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
        """Toggle Windows' immersive dark mode on a top-level window (Win10 1809+)."""
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
    """Ensures the inline editor in the tree has a thin border and proper padding."""
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setFrame(False)  # remove thick native frame
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
        self.resize(1280, 860)

        self.catalog_root = catalog_root
        self.current_path: Path | None = None  # file path when editing file
        self.current_folder: Path | None = None  # folder when editing folder meta

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

        # File system model / proxy
        self.fs_model = QFileSystemModel(self)
        self.fs_model.setReadOnly(False)
        self.fs_model.setRootPath(str(self.catalog_root))
        self.fs_model.setNameFilters(["*.md"])
        self.fs_model.setNameFilterDisables(False)

        # 🔧 handle in-place renames (double-click) so we can fix folder JSON
        self.fs_model.fileRenamed.connect(self.on_fs_file_renamed)

        self.proxy = DescProxyModel(self)
        self.proxy.setSourceModel(self.fs_model)

        # Left: Tree
        self.tree = QTreeView(self)
        self.tree.setModel(self.proxy)
        self.tree.setItemDelegate(SlimLineEditDelegate(self.tree))  # slim inline editor
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

        # Folder metadata panel (fills remaining space)
        self.folder_panel = QGroupBox("Folder Metadata", self)
        fp_layout = QFormLayout(self.folder_panel)
        # Fields for folder metadata
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

        # Metadata tab
        meta_tab = QWidget(self); meta_v = QVBoxLayout(meta_tab); meta_v.setContentsMargins(0,0,0,0); meta_v.setSpacing(8)
        self.fields_group = QGroupBox("Circuit Metadata", meta_tab)
        self.fields_form = QFormLayout(self.fields_group)
        self.field_widgets: dict[str, QLineEdit] = {}
        for label, placeholder in FIELD_ORDER:
            le = QLineEdit(self.fields_group); le.setPlaceholderText(placeholder)
            self.field_widgets[label] = le
            self.fields_form.addRow(label + ":", le)
        meta_v.addWidget(self.fields_group)
        self.tabs.addTab(meta_tab, "Metadata")

        # Used On tab
        used_tab = QWidget(self); used_v = QVBoxLayout(used_tab)
        self.used_table = QTableWidget(0, 2, used_tab)
        self.used_table.setHorizontalHeaderLabels(["PN", "Occurrences"])
        self.used_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.used_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.used_table.verticalHeader().setVisible(False)
        used_v.addWidget(self.used_table)
        used_btns = QHBoxLayout(); self.btn_add_row = QPushButton("Add Row"); self.btn_del_row = QPushButton("Remove Selected")
        self.btn_add_row.clicked.connect(self.add_used_row); self.btn_del_row.clicked.connect(self.remove_used_row)
        used_btns.addWidget(self.btn_add_row); used_btns.addWidget(self.btn_del_row); used_btns.addItem(QSpacerItem(40,20,QSizePolicy.Expanding,QSizePolicy.Minimum))
        used_v.addLayout(used_btns)
        self.tabs.addTab(used_tab, "Used On")

        # Netlist tab
        net_tab = QWidget(self); net_v = QVBoxLayout(net_tab)
        self.netlist_edit = QTextEdit(net_tab)
        self.netlist_edit.setPlaceholderText("(paste or type your netlist here)")
        net_v.addWidget(self.netlist_edit)
        self.tabs.addTab(net_tab, "Netlist")

        # Partlist tab (UPDATED HEADERS)
        pl_tab = QWidget(self); pl_v = QVBoxLayout(pl_tab)
        self.part_table = QTableWidget(0, len(PARTLIST_HEADERS), pl_tab)
        self.part_table.setHorizontalHeaderLabels(PARTLIST_HEADERS)
        # reasonable sizing
        self.part_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Ref Des
        self.part_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Type
        self.part_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)           # Value/Description
        self.part_table.verticalHeader().setVisible(False)
        pl_v.addWidget(self.part_table)
        pl_btns = QHBoxLayout()
        self.btn_pl_add = QPushButton("Add Row"); self.btn_pl_del = QPushButton("Remove Selected")
        self.btn_pl_add.clicked.connect(lambda: self.add_row(self.part_table, default=["", "", ""]))
        self.btn_pl_del.clicked.connect(lambda: self.remove_selected_row(self.part_table))
        pl_btns.addWidget(self.btn_pl_add); pl_btns.addWidget(self.btn_pl_del); pl_btns.addItem(QSpacerItem(40,20,QSizePolicy.Expanding,QSizePolicy.Minimum))
        pl_v.addLayout(pl_btns)
        self.tabs.addTab(pl_tab, "Partlist")

        # Pin Interface tab (UPDATED HEADERS)
        pin_tab = QWidget(self); pin_v = QVBoxLayout(pin_tab)
        self.pin_table = QTableWidget(0, len(PIN_HEADERS), pin_tab)
        self.pin_table.setHorizontalHeaderLabels(PIN_HEADERS)
        for c in range(len(PIN_HEADERS)):
            # Let Description stretch
            mode = QHeaderView.Stretch if PIN_HEADERS[c] == "Description" else QHeaderView.ResizeToContents
            self.pin_table.horizontalHeader().setSectionResizeMode(c, mode)
        self.pin_table.verticalHeader().setVisible(False)
        pin_v.addWidget(self.pin_table)
        pin_btns = QHBoxLayout()
        self.btn_pin_add = QPushButton("Add Row"); self.btn_pin_del = QPushButton("Remove Selected")
        self.btn_pin_add.clicked.connect(lambda: self.add_row(self.pin_table, default=["", "", "", ""]))
        self.btn_pin_del.clicked.connect(lambda: self.remove_selected_row(self.pin_table))
        pin_btns.addWidget(self.btn_pin_add); pin_btns.addWidget(self.btn_pin_del); pin_btns.addItem(QSpacerItem(40,20,QSizePolicy.Expanding,QSizePolicy.Minimum))
        pin_v.addLayout(pin_btns)
        self.tabs.addTab(pin_tab, "Pin Interface")

        # Tests tab (UPDATED HEADERS)
        test_tab = QWidget(self); test_v = QVBoxLayout(test_tab)
        self.test_table = QTableWidget(0, len(TEST_HEADERS), test_tab)
        self.test_table.setHorizontalHeaderLabels(TEST_HEADERS)
        for c in range(len(TEST_HEADERS)):
            mode = QHeaderView.Stretch if TEST_HEADERS[c] in ("Description", "Note") else QHeaderView.ResizeToContents
            self.test_table.horizontalHeader().setSectionResizeMode(c, mode)
        self.test_table.verticalHeader().setVisible(False)
        test_v.addWidget(self.test_table)
        test_btns = QHBoxLayout()
        self.btn_test_add = QPushButton("Add Row"); self.btn_test_del = QPushButton("Remove Selected")
        self.btn_test_add.clicked.connect(lambda: self.add_row(self.test_table, default=["", "", "", ""]))
        self.btn_test_del.clicked.connect(lambda: self.remove_selected_row(self.test_table))
        test_btns.addWidget(self.btn_test_add); test_btns.addWidget(self.btn_test_del); test_btns.addItem(QSpacerItem(40,20,QSizePolicy.Expanding,QSizePolicy.Minimum))
        test_v.addLayout(test_btns)
        self.tabs.addTab(test_tab, "Tests")

        right_layout.addWidget(self.tabs, 1)

        # Splitter
        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(self.tree)
        splitter.addWidget(right_container)
        splitter.setStretchFactor(0, 1); splitter.setStretchFactor(1, 2); splitter.setSizes([360, 920])

        central = QWidget(self); outer = QHBoxLayout(central); outer.setContentsMargins(8,8,8,8); outer.setSpacing(8)
        outer.addWidget(splitter); self.setCentralWidget(central)

        # Apply dark stylesheet
        self.apply_dark_styles()

        self.show_file_ui(False)

    def on_fs_file_renamed(self, dir_path_str: str, old_name: str, new_name: str):
        """When a file/folder is renamed inline in the tree, fix folder JSON and update selection labels."""
        try:
            dir_path = Path(dir_path_str)
            old_path = dir_path / old_name
            new_path = dir_path / new_name

            if new_path.is_dir():
                # Rename or remove the obsolete <old_name>.json inside the renamed folder
                old_meta = new_path / f"{old_name}.json"
                new_meta = new_path / f"{new_name}.json"

                if old_meta.exists():
                    if new_meta.exists():
                        # Correct file already present — old one is obsolete
                        try:
                            old_meta.unlink()
                        except Exception:
                            pass
                    else:
                        try:
                            old_meta.rename(new_meta)
                        except Exception:
                            # Fallback: copy then delete
                            try:
                                data = old_meta.read_text(encoding="utf-8")
                                new_meta.write_text(data, encoding="utf-8")
                                old_meta.unlink()
                            except Exception:
                                pass

                # Refresh description column for this folder
                self.proxy.refresh_desc(new_path)

                # If this was the folder we're editing, update references/UI label
                if self.current_folder and (self.current_folder == old_path or self.current_folder.name == old_name and self.current_folder.parent == dir_path):
                    self.current_folder = new_path
                    self.path_label.setText(f"Folder: {new_path}")

            else:
                # File rename: refresh description if it's a .md
                if new_path.suffix.lower() == ".md":
                    self.proxy.refresh_desc(new_path)
                if self.current_path and (self.current_path == old_path or (self.current_path.name == old_name and self.current_path.parent == dir_path)):
                    self.current_path = new_path
                    self.path_label.setText(f"File: {new_path}")

        except Exception:
            # Keep UI responsive even if something odd happens
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
            /* Slim inline editor inside the tree */
            QTreeView QLineEdit {
                background-color: #2A2D31; color: #E6E6E6; border: 1px solid #3A3F44; padding: 2px 4px;
            }
            QPushButton {
                background-color: #2F343A; color: #E6E6E6; border: 1px solid #444; border-radius: 6px; padding: 6px 12px;
            }
            QPushButton:hover { background-color: #3A4047; }
            QPushButton:pressed { background-color: #2A2F35; }
            QTreeView {
                background-color: #1E2124; alternate-background-color: #24272B; border: 1px solid #3A3F44;
            }
            QTreeView::item:selected { background: #3B4252; color: #E6E6E6; }
            QHeaderView::section {
                background-color: #2A2D31; color: #E6E6E6; border: 0px; padding: 6px; font-weight: 600;
            }
            QTableWidget {
                background-color: #1E2124; color: #E6E6E6; gridline-color: #3A3F44;
                alternate-background-color: #24272B; border: 1px solid #3A3F44; border-radius: 6px;
            }
            QTabBar::tab {
                background: #2A2D31; color: #E6E6E6; padding: 8px 12px; margin-right: 2px;
                border-top-left-radius: 6px; border-top-right-radius: 6px;
            }
            QTabBar::tab:selected { background: #3A3F44; }
            QTabBar::tab:hover { background: #34383D; }

            /* Popups (client area) */
            QMessageBox, QInputDialog, QFileDialog {
                background-color: #202225; color: #E6E6E6;
            }
            QMessageBox QPushButton, QInputDialog QPushButton, QFileDialog QPushButton {
                background-color: #2F343A; color: #E6E6E6; border: 1px solid #444;
                border-radius: 6px; padding: 6px 12px;
            }
            QMessageBox QPushButton:hover, QInputDialog QPushButton:hover, QFileDialog QPushButton:hover {
                background-color: #3A4047;
            }
        """)

    # --- Dialog helpers ---------------------------------------------------------
    def _apply_dark(self, dlg):
        try:
            apply_windows_dark_titlebar(dlg)
        except Exception:
            pass

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
        mb.setStandardButtons(QMessageBox.Ok)
        self._apply_dark(mb)
        mb.exec_()

    def error(self, title: str, text: str):
        mb = QMessageBox(self)
        mb.setWindowTitle(title)
        mb.setText(text)
        mb.setIcon(QMessageBox.Critical)
        mb.setStandardButtons(QMessageBox.Ok)
        self._apply_dark(mb)
        mb.exec_()

    # --- UI helpers -------------------------------------------------------------
    def show_file_ui(self, file_selected: bool):
        """Toggle between file editor (tabs) and folder metadata panel."""
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

    # --- Table helpers ----------------------------------------------------------
    def add_row(self, table: QTableWidget, default=None):
        r = table.rowCount(); table.insertRow(r)
        cols = table.columnCount()
        default = default or ["" for _ in range(cols)]
        for c in range(cols):
            table.setItem(r, c, QTableWidgetItem(default[c]))

    def remove_selected_row(self, table: QTableWidget):
        r = table.currentRow()
        if r >= 0: table.removeRow(r)

    # --- Selection / load -------------------------------------------------------
    def on_tree_selection(self, *_):
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
        meta = self.read_folder_meta(folder)
        # Fill fields
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

    def load_file(self, path: Path):
        try:
            text = path.read_text(encoding="utf-8")
        except Exception as e:
            self.error("Error", f"Failed to read file:\n{e}"); return

        self.current_path = path
        self.path_label.setText(f"File: {path}")

        fields, used_rows, netlist, part_rows, pin_rows, test_rows = self.parse_markdown(text)

        # Metadata fields
        for key, _ in FIELD_ORDER:
            self.field_widgets[key].setText(fields.get(key, ""))

        # Used On
        self.used_table.setRowCount(0)
        for pn, occ in used_rows:
            self.add_row(self.used_table, [pn, occ])

        # Netlist
        self.netlist_edit.setPlainText(netlist)

        # Partlist
        self.part_table.setRowCount(0)
        for row in part_rows:
            self.add_row(self.part_table, row + [""] * (len(PARTLIST_HEADERS) - len(row)))

        # Pin Interface
        self.pin_table.setRowCount(0)
        for row in pin_rows:
            self.add_row(self.pin_table, row + [""] * (len(PIN_HEADERS) - len(row)))

        # Tests
        self.test_table.setRowCount(0)
        for row in test_rows:
            self.add_row(self.test_table, row + [""] * (len(TEST_HEADERS) - len(row)))

        self.proxy.refresh_desc(path)

    # --- Parse / Build markdown -------------------------------------------------
    def parse_markdown(self, text: str):
        lines = [ln.rstrip("\n") for ln in text.splitlines()]
        fields = {k: "" for k, _ in FIELD_ORDER}
        used_rows, netlist = [], ""
        part_rows, pin_rows, test_rows = [], [], []

        def find_section(title: str):
            header = f"## {title}".lower()
            for i, ln in enumerate(lines):
                if ln.strip().lower() == header:
                    return i
            return None

        # Metadata table
        i = 0; n = len(lines)
        while i < n:
            if lines[i].strip().lower().startswith("| field") and "| value" in lines[i].lower():
                i += 2
                while i < n and lines[i].strip().startswith("|"):
                    m = MD_ROW_RE.match(lines[i].strip())
                    if m:
                        field = m.group("field").strip()
                        value = m.group("value").strip()
                        if field in fields: fields[field] = value
                    i += 1
                break
            i += 1

        # Used On (table PN/Occurrences)
        uix = find_section("Used On")
        if uix is not None:
            j = uix + 1
            while j < n and not lines[j].strip().startswith("|"): j += 1  # header
            j += 2  # skip header + divider
            while j < n and lines[j].strip().startswith("|"):
                m = MD_ROW_RE.match(lines[j].strip())
                if m:
                    used_rows.append([m.group("field").strip(), m.group("value").strip()])
                j += 1

        # Netlist (free text until next ## section)
        nix = find_section("Netlist")
        if nix is not None:
            j = nix + 1
            chunk = []
            while j < n and not lines[j].strip().startswith("## "):
                chunk.append(lines[j])
                j += 1
            while chunk and not chunk[0].strip(): chunk.pop(0)
            while chunk and not chunk[-1].strip(): chunk.pop()
            netlist = "\n".join(chunk)

        # Generic table parser for Partlist / Pin Interface / Tests
        def parse_table(at_index: int):
            rows = []
            if at_index is None: return rows
            j = at_index + 1
            while j < n and not lines[j].strip().startswith("|"): j += 1  # header row
            j += 1  # divider line
            while j < n and lines[j].strip().startswith("|"):
                raw = lines[j].strip().strip("|")
                cells = [c.strip() for c in raw.split("|")]
                rows.append(cells)
                j += 1
            return rows

        part_rows = parse_table(find_section("Partlist"))
        pin_rows  = parse_table(find_section("Pin Interface"))
        test_rows = parse_table(find_section("Tests"))

        return fields, used_rows, netlist, part_rows, pin_rows, test_rows

    def build_markdown(self, fields: dict, used_rows: list, netlist: str,
                       part_rows: list, pin_rows: list, test_rows: list) -> str:
        # Metadata
        meta_lines = [
            "| Field                  | Value                     |",
            "| ---------------------- | ------------------------- |",
        ]
        for key, _ in FIELD_ORDER:
            meta_lines.append(f"| {key:<22} | {fields.get(key,'').strip()} |")

        # Used On
        used = ["## Used On", "", "| PN         | Occurrences |", "| ---------- | ----------- |"]
        if used_rows:
            for pn, occ in used_rows:
                used.append(f"| {pn or '(None)'} | {occ or '0'} |")
        else:
            used.append("| (None)     | 0           |")

        # Netlist
        net = ["## Netlist", "", netlist.strip() if netlist.strip() else "(paste or type your netlist here)"]

        # Table builder
        def build_table(headers, rows):
            out = ["| " + " | ".join(headers) + " |",
                   "| " + " | ".join("-" * len(h) for h in headers) + " |"]
            for r in rows:
                rr = (r + [""] * len(headers))[:len(headers)]
                out.append("| " + " | ".join(rr) + " |")
            return out

        part = ["## Partlist", ""]
        part += build_table(PARTLIST_HEADERS, part_rows)

        pin = ["## Pin Interface", ""]
        pin += build_table(PIN_HEADERS, pin_rows)

        tests = ["## Tests", ""]
        tests += build_table(TEST_HEADERS, test_rows)

        out = [
            "# Circuit Metadata",
            "",
            f"**Last Updated:** {today_iso()}",
            "",
            "\n".join(meta_lines),
            "",
            "\n".join(used),
            "",
            "\n".join(net),
            "",
            "\n".join(part),
            "",
            "\n".join(pin),
            "",
            "\n".join(tests),
            ""
        ]
        return "\n".join(out)

    # --- Save -------------------------------------------------------------------
    def save_from_form(self):
        """
        If a file is selected -> save Markdown.
        If a folder is selected (current_path is None but current_folder set) -> save folder JSON.
        """
        if self.current_path and self.is_markdown(self.current_path):
            # Save file
            fields = {k: w.text().strip() for k, w in self.field_widgets.items()}

            used_rows = []
            for r in range(self.used_table.rowCount()):
                pn = self.used_table.item(r, 0).text().strip() if self.used_table.item(r,0) else ""
                oc = self.used_table.item(r, 1).text().strip() if self.used_table.item(r,1) else ""
                used_rows.append([pn, oc])

            netlist = self.netlist_edit.toPlainText()

            def harvest(table: QTableWidget, cols: int):
                rows = []
                for r in range(table.rowCount()):
                    row = []
                    for c in range(cols):
                        it = table.item(r, c)
                        row.append(it.text().strip() if it else "")
                    rows.append(row)
                return rows

            part_rows = harvest(self.part_table, len(PARTLIST_HEADERS))
            pin_rows  = harvest(self.pin_table,  len(PIN_HEADERS))
            test_rows = harvest(self.test_table, len(TEST_HEADERS))

            text = self.build_markdown(fields, used_rows, netlist, part_rows, pin_rows, test_rows)

            try:
                self.current_path.write_text(text, encoding="utf-8")
            except Exception as e:
                self.error("Error", f"Failed to save file:\n{e}"); return

            self.proxy.refresh_desc(self.current_path)
            self.info("Saved", "Catalog entry saved.")
            return

        # Otherwise save folder metadata if a folder is selected
        if self.current_folder and self.current_folder.exists() and self.current_folder.is_dir():
            meta_p = folder_meta_path(self.current_folder)
            # Preserve 'Created' if present
            created = today_iso()
            if meta_p.exists():
                try:
                    old = json.loads(meta_p.read_text(encoding="utf-8"))
                    created = old.get("Created", created)
                except Exception:
                    pass

            # Build tags list (store as list)
            raw_tags = (self.folder_tags.text() or "").strip()
            tags_list = [t.strip() for t in raw_tags.split(",") if t.strip()]

            meta = {
                "TITLE": (self.folder_title.text() or "").upper(),
                "DESCRIPTION": (self.folder_desc.text() or "").upper(),
                "Summary": self.folder_summary.toPlainText().strip(),
                "Owner": (self.folder_owner.text() or "").strip(),
                "Tags": tags_list,
                "Created": created,
                "Last Updated": today_iso(),
            }
            try:
                meta_p.write_text(json.dumps(meta, indent=2), encoding="utf-8")
            except Exception as e:
                self.error("Error", f"Failed to save folder metadata:\n{e}"); return

            # Refresh tree "Description" (TITLE)
            self.proxy.refresh_desc(self.current_folder)
            self.folder_created.setText(meta["Created"])
            self.folder_updated.setText(meta["Last Updated"])

            self.info("Saved", "Folder metadata saved.")
            return

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
            # Backfill defaults if older schema
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
        """Zip the folder containing the running script, then move the zip INSIDE that folder (avoid self-inclusion)."""
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
        """Open the selected folder (or the folder containing the selected file) in the OS file manager."""
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
                # Linux / others
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
            # Initialize meta
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
                # --- Folder rename flow ---
                old_folder_name = path.name
                # 1) Rename the folder
                path.rename(new_path)

                # 2) Fix the metadata json name inside the *renamed* folder
                old_meta = new_path / f"{old_folder_name}.json"
                new_meta = new_path / f"{new_path.name}.json"

                if old_meta.exists():
                    if new_meta.exists():
                        # If a correct json already exists, the old one is obsolete
                        try:
                            old_meta.unlink()
                        except Exception:
                            pass
                    else:
                        try:
                            old_meta.rename(new_meta)
                        except Exception:
                            # If rename fails for any reason, fall back to copy+delete
                            try:
                                data = old_meta.read_text(encoding="utf-8")
                                new_meta.write_text(data, encoding="utf-8")
                                old_meta.unlink()
                            except Exception:
                                pass

                # Refresh tree desc (shows folder TITLE)
                self.proxy.refresh_desc(new_path)

                # Update selection bookkeeping if needed
                if self.current_folder and self.current_folder == path:
                    self.current_folder = new_path
                    self.path_label.setText(f"Folder: {new_path}")

            else:
                # --- File rename flow ---
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

    # Used On helpers (add/remove a row)
    def add_used_row(self):
        r = self.used_table.rowCount()
        self.used_table.insertRow(r)
        self.used_table.setItem(r, 0, QTableWidgetItem(""))
        self.used_table.setItem(r, 1, QTableWidgetItem("1"))

    def remove_used_row(self):
        r = self.used_table.currentRow()
        if r >= 0:
            self.used_table.removeRow(r)

# --- Boot ---------------------------------------------------------------------

def ensure_catalog_root(start_dir: Path | None = None) -> Path:
    root = DEFAULT_CATALOG_DIR if start_dir is None else start_dir
    if not root.exists():
        try:
            root.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
    if not root.exists() or not root.is_dir():
        # Use a temporary dialog with dark titlebar + non-native Qt body
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

    # 🗂️ app icon
    icon = make_emoji_icon("🗂️", px=256)
    app.setWindowIcon(icon)

    root = ensure_catalog_root()
    win = CatalogWindow(root, icon)
    win.show()
    apply_windows_dark_titlebar(win)
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
