#!/usr/bin/env python3
"""
Parts Catalog Tool — PyQt5 (Dark-only, Focused Editor v15)

Revision History (Code)
-----------------------
- 2025-10-03 — Git pre-save pull, Git enable/disable, and debug console tweaks
  Author: Nolan Manteufel
  Description of Change:
    - Git safety before save:
      • Added _safe_pull_before_save(): fetches and pulls (rebase + autostash) before writing/committing.
      • On conflict, rebase aborts gracefully; file is still saved locally; commit/push are skipped with a clear status message.
    - Git scheduling unchanged (push happens ~60s after a successful commit), but now:
      • Respect “git_enabled” setting everywhere; when disabled, no pulls/commits/pushes are attempted.
      • “Sync push in: —” remains static while Git is disabled.
    - Settings & Tools:
      • Added “Enable Git (pull/commit/push)” checkbox (backed by settings["git_enabled"], default True).
      • No other UI layout changes.
    - Autosave/push loop:
      • _tick_autosave_countdown() updated to read git_enabled: when False, the push countdown is idle and the label stays “—”.
    - Save pipeline:
      • save_from_form() calls _safe_pull_before_save() before commit; skips commit/push on pull failure/conflict.
    - Debug console:
      • Removed the far-right toolbar buttons (Copy, Copy All, Save Log, Clear). 
      • Kept the right-click context menu actions exactly as they are.
    - Folder/File panel toggle:
      • Ensured the correct visibility logic via _toggle_panels() (folder panel only in folder mode; tabs only in file mode).
    - Repository stats (lines_total):
      • Guarded the line counter to avoid counting binary files and .git internals; heavy work is batched to avoid UI stalls.

- 2025-10-01 — Schematic PDFs, Viewer Fit, and PDF UX Hardening
  Author: Nolan Manteufel
  Description of Change:
    - Added a new “Schematic” main tab (immediately after “Metadata”).
      • Each Schematic subtab is created by scanning a configurable folder for PN_*.pdf.
      • Subtab label = the * portion of PN_*.pdf (e.g., PN_sch.pdf → “sch”).
      • Uses lazy-load: the PDF file is only rendered when its subtab is first selected.
    - New Settings field: “Schematic Folder”.
      • Accepts absolute or relative paths (relative paths resolve against the current file’s folder).
      • Includes a Browse… button; setting is persisted in parts_catalog_settings.json.
    - PDF viewer (PyMuPDF-based) improvements:
      • Wheel scroll (normal), Ctrl + Wheel zoom, left-drag to pan.
      • On load and window resize, auto “fit-to-width” (not cropped) until the user manually zooms.
      • Clear diagnostics: error/notice labels always show the full, resolved path attempted.
      • Hardened lazy-loader to recompute path if the stored property is missing.
    - Show PDFs in the file tree (read-only behavior):
      • PDFs appear in the left pane with Description “[PDF document]”.
      • PDFs cannot be renamed, deleted, dragged, or moved from within the app.
      • Selecting a PDF does not open any editors; right pane remains non-editable.
    - Safety: Disabled batch rename (multi-file) entirely to protect external references.
    - Cleanup: Removed the optional trigger that rebuilt schematic tabs when Revision History table cells changed.
      • Schematic tabs still rebuild on file load (and optionally on Part Number changes if enabled).

- 2025-09-17 — v13, Initial release
  Author: Nolan Manteufel
  Description of Change:
    Parts Catalog Tool — PyQt5 (Dark-only, Focused Editor v13)
    - Plain-text editors everywhere (no rich text / no pasted colors)
    - Text Zoom: Ctrl+= / Ctrl++ (Zoom In), Ctrl+- (Zoom Out), Ctrl+0 (Reset)
    - File counter in the nav pane footer: "Files in folder | Total"
    - Default revision author:
        • New Revision rows default the “By” column to:
            (Global Default Owner Name, if set) else (current folder Owner)
        • Configurable via Settings & Tools (toolbar)
    - Default new-file name (robust to deletions):
        • Proposes name = <prefix><number> (zero-padded)
          number is tracked per-prefix in settings (“last_file_counters”),
          then incremented until a free filename is found.
        • Examples: "",3 → "001"; "LM",3 → "LM001"
        • Configurable via Settings & Tools (prefix + width)
    - High-level Settings file (JSON) next to the script:
        • stores: owner_default_name, file_name_prefix, number_width, last_file_counters, stats
        • stats updated on saves (files_total, lines_total across entire repo, last_saved_path, last_updated)
    - Settings & Tools (toolbar):
        • Save Current (same as Ctrl+S)
        • Export Single File: concatenates all .md files in repository into one .md
          (always placed at the TOP LEVEL of the catalog root)
        • Archive (zip of the folder containing this script)
    - Single Save available via Tools dialog and Ctrl+S
    - Metadata tab has subtabs:
        • Introduction (Title, Part Number, Notes, etc.)
        • Revision History (global table)
        • Variant Details (list of items)
    - Editor tabs: Metadata, Netlist (plain), Partlist (plain),
                   Circuit Description (plain), Circuit Theory (plain),
                   Design Analysis (plain), Review (raw Markdown, plain)
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
    - Special behavior:
        • Any file whose name starts with "catlog_" OR "catalog_" will lock all tabs except the Review tab.

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
from typing import Optional
import shlex
import threading, queue
from pathlib import Path

from PyQt5.QtCore import Qt, QSortFilterProxyModel, QModelIndex, QTimer
from PyQt5.QtGui import QKeySequence, QIcon, QPixmap, QPainter, QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFileSystemModel, QTreeView, QToolBar, QFileDialog,
    QInputDialog, QMessageBox, QLabel, QAbstractItemView,
    QLineEdit, QHeaderView, QPushButton, QSpacerItem, QSizePolicy,
    QGroupBox, QSplitter, QTabWidget, QTableWidget, QTableWidgetItem,
    QStyleFactory, QStyledItemDelegate, QFormLayout, QPlainTextEdit,
    QDialog, QDialogButtonBox, QSpinBox, QCheckBox
)
from PyQt5.QtWidgets import QAction as QtAction


# ---------- Settings (high-level app config) -----------------------------------
def _script_dir() -> Path:
    try:
        if getattr(sys, 'frozen', False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()

APP_TITLE = "Parts Catalog Tool"
DEFAULT_CATALOG_DIR = _script_dir()
SETTINGS_PATH = _script_dir() / "parts_catalog_settings.json"

DEFAULT_SETTINGS = {
    "schematic_folder": "docs",    # Default subfolder (relative to the current file's folder) for schematic PDFs
    "owner_default_name": "",      # If non-empty, overrides folder Owner for default “By”
    "file_name_prefix": "",        # Prefix for new files (e.g., "LM")
    "number_width": 3,             # Zero-pad width (e.g., 3 → 001)
    "last_file_counters": {},      # per-prefix counters, e.g., {"": 7, "LM": 12}
    "stats": {
        "files_total": 0,          # *.md files across repository
        "lines_total": 0,          # lines across text files in the repository (.md, .json, .py, etc.)
        "last_saved_path": "",
        "last_updated": ""
    },
    "git_remote_url": "",              # e.g., "git@github.com:org/repo.git" (SSH recommended)
    "git_branch": "main",              # default branch to pull/push
    "auto_push_delay_seconds": 60,     # countdown refreshed after each save
    "show_debug_console": True,
    "git_enabled": True         
}

def load_settings() -> dict:
    s = DEFAULT_SETTINGS.copy()
    s["stats"] = DEFAULT_SETTINGS["stats"].copy()
    s["last_file_counters"] = DEFAULT_SETTINGS["last_file_counters"].copy()
    try:
        if SETTINGS_PATH.exists():
            disk = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
            for k, v in disk.items():
                if k == "stats" and isinstance(v, dict):
                    s["stats"].update(v)
                elif k == "last_file_counters" and isinstance(v, dict):
                    s["last_file_counters"].update(v)
                else:
                    s[k] = v
    except Exception:
        pass
    return s

def save_settings(s: dict):
    try:
        out = DEFAULT_SETTINGS.copy()
        out["stats"] = DEFAULT_SETTINGS["stats"].copy()
        out["last_file_counters"] = DEFAULT_SETTINGS["last_file_counters"].copy()
        for k, v in s.items():
            if k == "stats" and isinstance(v, dict):
                out["stats"].update(v)
            elif k == "last_file_counters" and isinstance(v, dict):
                out["last_file_counters"].update(v)
            else:
                out[k] = v
        SETTINGS_PATH.write_text(json.dumps(out, indent=2), encoding="utf-8")
    except Exception:
        pass

# ---------- Git helpers (no UI) -----------------------------------------------
def _run(cmd: list[str], cwd: Path) -> tuple[int, str, str]:
    _dbg(f"GIT RUN: {' '.join(shlex.quote(x) for x in cmd)}  (cwd={cwd})")
    try:
        env = os.environ.copy()
        env.setdefault("GIT_TERMINAL_PROMPT", "0")
        env.setdefault(
            "GIT_SSH_COMMAND",
            "ssh -o BatchMode=yes -o ConnectTimeout=5 -o StrictHostKeyChecking=accept-new"
        )
        p = subprocess.Popen(
            cmd, cwd=str(cwd),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, env=env
        )
        out, err = p.communicate()
        rc = p.returncode
        _dbg(f"GIT EXIT {rc} | out: {_clip(out)} | err: {_clip(err)}")
        return rc, (out or "").strip(), (err or "").strip()
    except Exception as e:
        _dbg(f"GIT EXC: {e}")
        return 1, "", str(e)

def _git(cwd: Path, *args: str) -> tuple[int, str, str]:
    _dbg(f"GIT: git {' '.join(args)}")
    return _run(["git", *args], cwd)

def ensure_git_repo(repo_root: Path, settings: dict) -> None:
    _dbg(f"Ensure repo at: {repo_root}")
    if not (repo_root / ".git").exists():
        _dbg("Repo not initialized → git init")
        _git(repo_root, "init")
    rc, out, _ = _git(repo_root, "symbolic-ref", "--short", "HEAD")
    if rc != 0:
        branch = (settings.get("git_branch") or "main").strip() or "main"
        _git(repo_root, "checkout", "-b", branch)
        _dbg(f"Ensure current branch = {(settings.get('git_branch') or 'main')}")
    remote = os.environ.get("PARTS_CATALOG_GIT_REMOTE") or (settings.get("git_remote_url") or "").strip()
    if remote:
        rc, remotes, _ = _git(repo_root, "remote")
        rems = remotes.splitlines() if rc == 0 else []
        if "origin" not in rems:
            _git(repo_root, "remote", "add", "origin", remote)
        else:
            _git(repo_root, "remote", "set-url", "origin", remote)

def git_pull(repo_root: Path, branch: str) -> None:
    rc, remotes, _ = _git(repo_root, "remote")
    if rc == 0 and "origin" in remotes.splitlines():
        _git(repo_root, "fetch", "origin", branch)
        _git(repo_root, "merge", f"origin/{branch}")

def _highest_rev_from_table(rev_table) -> str:
    try:
        rows = rev_table.rowCount()
        for r in range(rows - 1, -1, -1):
            it = rev_table.item(r, 0)
            val = (it.text().strip() if it else "")
            if val:
                return val
    except Exception:
        pass
    return "-"

def git_commit_all(repo_root: Path, message: str) -> bool:
    _git(repo_root, "add", "-A")
    rc, _, _ = _git(repo_root, "diff", "--cached", "--quiet")
    if rc == 0:
        return False
    rc, _, _ = _git(repo_root, "commit", "-m", message)
    return rc == 0

def git_push(repo_root: Path, branch: str) -> None:
    rc, remotes, _ = _git(repo_root, "remote")
    if rc == 0 and "origin" in remotes.splitlines():
        _git(repo_root, "push", "origin", f"HEAD:{branch}")

# ---------- App data model constants ------------------------------------------
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
SECTION_TITLES = { "cd": "Circuit Description", "ct": "Circuit Theory", "da": "Design Analysis" }

def today_iso():
    return datetime.date.today().isoformat()

def now_stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def folder_meta_path(folder: Path) -> Path:
    return folder / f"{folder.name}.json"

MD_ROW_RE = re.compile(r'^\|\s*(?P<field>[^|]+?)\s*\|\s*(?P<value>[^|]*?)\s*\|$')
DIVIDER_CELL_RE = re.compile(r'^:?-{2,}:?$')

def _divider_for(headers):
    return "| " + " | ".join("-" * len(h) for h in headers) + " |"

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

# ---------- NEW ENTRY TEMPLATE -------------------------------------------------
def _new_entry_template(default_owner: str) -> str:
    return (f"""# Circuit Metadata

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
| - | {today_iso()} | Initial release | {default_owner} |

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
""").strip() + "\n"

# ---------- PDF Viewer (lazy, multi-page) -------------------------------------
try:
    import fitz  # PyMuPDF
    _HAS_FITZ = True
except Exception:
    _HAS_FITZ = False

from PyQt5.QtWidgets import QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QLabel, QScrollBar
from PyQt5.QtGui import QImage, QPixmap, QPainter
from PyQt5.QtCore import Qt, QPoint

class PDFGraphicsView(QGraphicsView):
    """
    Minimal multi-page PDF renderer using PyMuPDF.
    - Renders all pages vertically in a single QGraphicsScene.
    - Normal wheel scroll; Ctrl + wheel for zoom; left-drag to pan.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self.setScene(self._scene)
        self._pages: list[QGraphicsPixmapItem] = []
        self._panning = False
        self._last_pos = QPoint()
        self._scale = 1.0
        self._user_zoomed = False  # if True, stop auto-fit-to-width on resize

        self.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.setDragMode(QGraphicsView.NoDrag)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.setMouseTracking(True)

    def fit_to_width(self):
        """Scale so the widest page fits the viewport width."""
        if not self._pages:
            return
        # Width of the widest rendered page (scene units = pixels at current transform=identity)
        page_width = max(p.pixmap().width() for p in self._pages if p.pixmap() and not p.pixmap().isNull())
        if page_width <= 0:
            return

        view_w = max(1, self.viewport().width())
        # Reset transform and compute uniform scale
        self.resetTransform()
        factor = view_w / page_width
        self._scale = factor
        self.scale(factor, factor)

        # Scroll to top-left so page isn't “cropped” from the left
        self.horizontalScrollBar().setValue(self.horizontalScrollBar().minimum())
        self.verticalScrollBar().setValue(self.verticalScrollBar().minimum())

    def clear(self):
        self._scene.clear()
        self._pages.clear()
        self.resetTransform()
        self._scale = 1.0

    def load_pdf(self, path: str):
        self.clear()
        if not _HAS_FITZ:
            self._scene.addText("PyMuPDF (fitz) not installed.\n\npip install pymupdf")
            return
        try:
            doc = fitz.open(path)
        except Exception as e:
            self._scene.addText(f"Failed to open PDF:\n{path}\n\n{e}")
            return

        y = 0
        for i in range(len(doc)):
            page = doc.load_page(i)
            # Render at 144 dpi-ish (2x), adjust if you want sharper
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pm = QPixmap.fromImage(img)
            item = QGraphicsPixmapItem(pm)
            item.setPos(0, y)
            self._scene.addItem(item)
            self._pages.append(item)
            y += pm.height() + 12  # vertical gap

        # Fit to page width by default
        self._user_zoomed = False
        self.fit_to_width()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not self._user_zoomed:
            self.fit_to_width()

    # Zoom only when Ctrl is pressed
    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            angle = event.angleDelta().y()
            factor = 1.15 if angle > 0 else 1/1.15
            self._scale = max(0.1, min(5.0, self._scale * factor))
            self.scale(factor, factor)
            self._user_zoomed = True
        else:
            super().wheelEvent(event)

    # Left-drag to pan
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._panning = True
            self._last_pos = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._panning:
            delta = event.pos() - self._last_pos
            self._last_pos = event.pos()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - delta.y())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._panning:
            self._panning = False
            self.setCursor(Qt.ArrowCursor)
        super().mouseReleaseEvent(event)

# ---------- Icon helper -------------------------------------------------------
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

# ---------- Proxy model for tree ---------------------------------------------
class DescProxyModel(QSortFilterProxyModel):
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
        name = sm.fileName(idx).lower()
        return name.endswith(".md") or name.endswith(".pdf")

    def flags(self, index):
        base = super().flags(index)
        if not index.isValid():
            return Qt.ItemIsDropEnabled
        sidx = self.mapToSource(index)
        sm = self.sourceModel()
        if sm.isDir(sidx):
            return base | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled

        # Files:
        name = sm.fileName(sidx).lower()
        if name.endswith(".pdf"):
            # Read-only: no drag, no drop
            return base & ~(Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
        # .md files can be dragged (to move) but not dropped-onto
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

# ---------- Windows dark titlebar (no-op on other OS) -------------------------
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
            res = dwm.DwmSetWindowAttribute(wintypes.HWND(hwnd), attribute, ctypes.byref(pv), ctypes.sizeof(pv))
            if res != 0:
                attribute = wintypes.DWORD(DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1)
                dwm.DwmSetWindowAttribute(wintypes.HWND(hwnd), attribute, ctypes.byref(pv), ctypes.sizeof(pv))
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

# ---------- Delegate -----------------------------------------------------------
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

# ---------- Settings & Tools Dialog -------------------------------------------
class SettingsToolsDialog(QDialog):
    def __init__(self, parent, settings: dict, do_save_cb, do_archive_cb, do_export_cb):
        super().__init__(parent)
        try:
            apply_windows_dark_titlebar(self)
        except Exception:
            pass

        self.setWindowTitle("Settings & Tools")
        self.settings = settings
        self.do_save_cb = do_save_cb
        self.do_archive_cb = do_archive_cb
        self.do_export_cb = do_export_cb

        root = QVBoxLayout(self)

        # --- SETTINGS
        settings_box = QGroupBox("Settings", self)
        sf = QFormLayout(settings_box)

        self.owner_default = QLineEdit(self)
        self.owner_default.setText(settings.get("owner_default_name",""))

        self.prefix = QLineEdit(self)
        self.prefix.setText(settings.get("file_name_prefix",""))

        self.width = QSpinBox(self); self.width.setRange(1, 8)
        self.width.setValue(int(settings.get("number_width",3)))

        # Schematic folder (string path; relative or absolute)
        self.schematic_folder_edit = QLineEdit(self)
        self.schematic_folder_edit.setPlaceholderText("e.g., docs  (relative to current file’s folder)")
        self.schematic_folder_edit.setText(settings.get("schematic_folder", "docs"))

        # Add a browse button next to it
        schem_row = QHBoxLayout()
        schem_row.addWidget(self.schematic_folder_edit)
        btn_browse_schem = QPushButton("Browse…", self)
        def _pick_schematic_dir():
            # Let user choose a folder; store as absolute. (User may still type a relative path manually.)
            dlg = QFileDialog(self, "Pick schematic folder")
            dlg.setFileMode(QFileDialog.Directory)
            dlg.setOption(QFileDialog.ShowDirsOnly, True)
            dlg.setOption(QFileDialog.DontUseNativeDialog, True)
            # Dark theming if available
            try:
                apply_windows_dark_titlebar(dlg)
            except Exception:
                pass
            if dlg.exec_() == QFileDialog.Accepted:
                path = dlg.selectedFiles()[0]
                if path:
                    self.schematic_folder_edit.setText(path)
        btn_browse_schem.clicked.connect(_pick_schematic_dir)
        schem_row.addWidget(btn_browse_schem)

        sf.addRow("Documents Folder:", schem_row)
        self.git_remote = QLineEdit(self); self.git_remote.setPlaceholderText("git@github.com:org/repo.git (SSH recommended)")
        self.git_remote.setText(settings.get("git_remote_url",""))

        self.git_branch = QLineEdit(self); self.git_branch.setText(settings.get("git_branch","main"))

        self.push_delay = QSpinBox(self); self.push_delay.setRange(5, 900); self.push_delay.setValue(int(settings.get("auto_push_delay_seconds",60)))

        sf.addRow("Git Remote (origin):", self.git_remote)
        sf.addRow("Git Branch:", self.git_branch)
        sf.addRow("Push Delay (sec):", self.push_delay)
        # Enable/disable Git entirely
        self.git_enabled_cb = QCheckBox("Enable Git (pull / commit / push)", self)
        self.git_enabled_cb.setChecked(bool(self.settings.get("git_enabled", True)))
        sf.addRow(self.git_enabled_cb)

        sf.addRow("Default Owner Name:", self.owner_default)
        sf.addRow("New File Prefix:", self.prefix)
        sf.addRow("Number Width:", self.width)

        stats = settings.get("stats", {})
        self.stats_label = QLabel(
            f"Files Total (*.md): {stats.get('files_total',0)} | "
            f"Lines in Repository (text files): {stats.get('lines_total',0)}\n"
            f"Last Saved: {stats.get('last_saved_path','')} @ {stats.get('last_updated','')}",
            self
        )
        self.stats_label.setStyleSheet("color:#A0A6AD;")
        sf.addRow(self.stats_label)
        root.addWidget(settings_box)

        # --- TOOLS
        tools_box = QGroupBox("Tools", self)
        tl = QVBoxLayout(tools_box)

        # Save Current
        btn_save = QPushButton("💾 Save Current (Ctrl+S)", self)
        btn_save.clicked.connect(lambda: self.do_save_cb())
        tl.addWidget(btn_save)

        # Export Single File (always to TOP LEVEL of catalog root)
        tl.addSpacing(6)
        tl.addWidget(QLabel("Export Single File — Concatenate all .md files (saved at catalog root)", self))
        fname_layout = QHBoxLayout()
        fname_layout.addWidget(QLabel("Output filename:", self))
        self.export_name = QLineEdit(self)
        self.export_name.setText(f"catalog_all_{now_stamp()}.md")
        fname_layout.addWidget(self.export_name)
        tl.addLayout(fname_layout)
        btn_export = QPushButton("⬇️ Export Single File", self)
        btn_export.clicked.connect(self._on_export_clicked)
        tl.addWidget(btn_export)

        tl.addSpacing(12)
        # Archive moved here
        btn_archive = QPushButton("📦 Archive (zip script folder)", self)
        btn_archive.clicked.connect(self.do_archive_cb)
        tl.addWidget(btn_archive)

        root.addWidget(tools_box)

        # OK / Cancel
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

        self.setMinimumWidth(560)

    def _on_export_clicked(self):
        name = (self.export_name.text().strip() or f"catalog_all_{now_stamp()}.md")
        self.do_export_cb(name)

    def result_settings(self) -> dict:
        out = self.settings.copy()
        out["git_remote_url"] = self.git_remote.text().strip()
        out["git_branch"] = self.git_branch.text().strip() or "main"
        out["auto_push_delay_seconds"] = int(self.push_delay.value())
        out["owner_default_name"] = self.owner_default.text().strip()
        out["file_name_prefix"] = self.prefix.text().strip()
        out["number_width"] = int(self.width.value())
        out["schematic_folder"] = self.schematic_folder_edit.text().strip()
        out["git_enabled"] = bool(self.git_enabled_cb.isChecked())
        return out

# ---------- Background Git queue (non-blocking) -------------------------------
class _GitBg:
    """Serializes Git tasks off the UI thread."""
    def __init__(self, repo_root: Path, ui_post: callable):
        self.repo_root = repo_root
        self._q: "queue.Queue[callable]" = queue.Queue()
        self._ui_post = ui_post
        t = threading.Thread(target=self._loop, daemon=True)
        t.start()

    def _loop(self):
        while True:
            fn = self._q.get()
            try:
                fn()
            except Exception:
                pass
            finally:
                self._q.task_done()

    def submit(self, fn: callable):
        self._q.put(fn)

    # helpers to marshal UI updates safely
    def ui(self, fn: callable):
        self._ui_post(fn)

# ---------- Debug logging (thread-safe) ---------------------------------------
def _ts():
    import datetime as _dt
    return _dt.datetime.now().strftime("%H:%M:%S")

class _UILogger:
    """Append debug lines into a QPlainTextEdit from any thread."""
    def __init__(self, post_to_ui_callable, append_fn_callable):
        self._post = post_to_ui_callable
        self._append = append_fn_callable

    def log(self, msg: str):
        line = f"[{_ts()}] {msg}"
        try:
            # Marshal to UI thread
            self._post(lambda: self._append(line))
        except Exception:
            pass

# Global debug hook used by git helpers (safe no-op if unset)
_DBG_HOOK = None

def set_global_debug_logger(fn):
    """Register a callable(str) used by helpers for debug output."""
    global _DBG_HOOK
    _DBG_HOOK = fn

def _dbg(msg: str):
    try:
        if _DBG_HOOK:
            _DBG_HOOK(msg)
    except Exception:
        pass

def _clip(s: str, limit: int = 600) -> str:
    s = (s or "")
    return (s if len(s) <= limit else s[:limit] + " …(trunc)")

# ---------- Main Window --------------------------------------------------------
class CatalogWindow(QMainWindow):
    def __init__(self, catalog_root: Path, app_icon: QIcon):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(app_icon)
        self.resize(1280, 900)

        self.settings = load_settings()

        self.catalog_root = catalog_root
        self.current_path: Path | None = None
        self.current_folder: Path | None = None

        # Dirty / autosave
        self.dirty = False
        # Review tab dirty flag
        self.review_dirty = False
        self.suppress_dirty = False

        self.autosave_interval_ms = 30000  # 30s
        self.autosave_interval_s = self.autosave_interval_ms // 1000
        self.autosave_remaining_s = self.autosave_interval_s

        self.base_tab_titles = []

        # Text zoom state
        self._base_font_pt = 12
        self._zoom_steps = 0  # relative to base

        # Toolbar
        tb = QToolBar("Main", self)
        tb.setMovable(False)
        self.addToolBar(tb)

        act_new_entry = QtAction("🧩 New Entry", self);  act_new_entry.triggered.connect(self.create_new_entry);  tb.addAction(act_new_entry)
        act_new_folder = QtAction("🗂️ New Folder", self); act_new_folder.triggered.connect(self.create_new_folder); tb.addAction(act_new_folder)
        act_rename     = QtAction("✏️ Rename", self);     act_rename.triggered.connect(self.rename_item);           tb.addAction(act_rename)
        act_delete     = QtAction("🗑️ Delete", self);     act_delete.triggered.connect(self.delete_item);           tb.addAction(act_delete)

        tb.addSeparator()

        act_open_loc = QtAction("📂 Open Location", self)
        act_open_loc.setToolTip("Open the selected folder (or the folder containing the selected file) in your file manager")
        act_open_loc.triggered.connect(self.open_file_location)
        tb.addAction(act_open_loc)

        # Save and Archive live in Settings & Tools; keep Ctrl+S here:
        self.act_save = QtAction("Save (Ctrl+S)", self)
        self.act_save.setShortcut(QKeySequence.Save)
        self.act_save.triggered.connect(self.save_from_form)
        self.addAction(self.act_save)

        # Settings & Tools
        tb.addSeparator()
        self.act_settings_tools = QtAction("⚙️ Settings & Tools", self)
        self.act_settings_tools.triggered.connect(self.open_settings_tools)
        tb.addAction(self.act_settings_tools)

        # Zoom actions
        tb.addSeparator()
        self.act_zoom_in = QtAction("Zoom In", self)
        self.act_zoom_out = QtAction("Zoom Out", self)
        self.act_zoom_reset = QtAction("Reset Zoom", self)
        self.act_zoom_in.setShortcuts([QKeySequence("Ctrl++"), QKeySequence("Ctrl+=")])
        self.act_zoom_out.setShortcut(QKeySequence("Ctrl+-"))
        self.act_zoom_reset.setShortcut(QKeySequence("Ctrl+0"))
        self.act_zoom_in.triggered.connect(lambda: self.adjust_zoom(+1))
        self.act_zoom_out.triggered.connect(lambda: self.adjust_zoom(-1))
        self.act_zoom_reset.triggered.connect(self.reset_zoom)
        tb.addAction(self.act_zoom_in); tb.addAction(self.act_zoom_out); tb.addAction(self.act_zoom_reset)

        # Right-aligned autosave countdown label
        spacer = QWidget(self); spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred); tb.addWidget(spacer)
        self.autosave_label = QLabel("Autosave in: —", self)
        self.autosave_label.setStyleSheet("color:#A0A6AD; padding: 0 6px;")
        tb.addWidget(self.autosave_label)

        # Post-to-UI helper (uses Qt event loop)
        self._ui_post = lambda fn: QTimer.singleShot(0, fn)
        self._gitbg = _GitBg(self.catalog_root, self._ui_post)

        # Push countdown state + label
        self.git_push_delay_s = int(self.settings.get("auto_push_delay_seconds", 60))
        self.git_push_remaining_s: Optional[int] = None
        self._push_inflight = False

        self.push_label = QLabel("Sync push in: —", self)
        self.push_label.setStyleSheet("color:#A0A6AD; padding: 0 6px;")
        tb.addWidget(self.push_label)

        # File system model / proxy
        self.fs_model = QFileSystemModel(self)
        self.fs_model.setReadOnly(False)
        self.fs_model.setRootPath(str(self.catalog_root))
        self.fs_model.setNameFilters(["*.md", "*.pdf"])
        self.fs_model.setNameFilterDisables(False)
        self.fs_model.fileRenamed.connect(self.on_fs_file_renamed)
        try:
            self.fs_model.directoryLoaded.connect(lambda _: self.update_file_counter())
        except Exception:
            pass

        self.proxy = DescProxyModel(self)
        self.proxy.setSourceModel(self.fs_model)

        # ===== LEFT PANE: tree + debug console + footer =====
        left_container = QWidget(self)
        left_v = QVBoxLayout(left_container)
        left_v.setContentsMargins(0, 0, 0, 0)
        left_v.setSpacing(4)

        # --- Tree ---
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

        # DnD (move)
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDropIndicatorShown(True)
        self.tree.setDefaultDropAction(Qt.MoveAction)
        self.tree.setDragDropMode(QAbstractItemView.DragDrop)
        self.tree.setDragDropOverwriteMode(False)

        # Multi-select (batch rename is disabled but selection is useful)
        self.tree.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.tree.selectionModel().selectionChanged.connect(self.on_tree_selection)

        # Add the tree first
        left_v.addWidget(self.tree, 1)

        # --- Debug console (under tree) ---
        # (QPlainTextEdit is already imported at top; no extra import needed here.)
        self.debug_console = QPlainTextEdit(self)
        self.debug_console.setReadOnly(True)
        # allow selecting + copying while still read-only
        self.debug_console.setTextInteractionFlags(
            Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard
        )
        self.debug_console.setMaximumBlockCount(2000)  # keep memory bounded
        self.debug_console.setPlaceholderText("Debug console… (Git + app messages)")
        self.debug_console.setStyleSheet(
            "QPlainTextEdit { background:#15171A; color:#B8C1CC; border:1px solid #3A3F44; }"
        )
        self.debug_console.setVisible(bool(self.settings.get("show_debug_console", True)))
        self.debug_console.setFixedHeight(140)

        def _append_debug_line(s: str):
            self.debug_console.appendPlainText(s)

        # Post-to-UI helper for logger (uses singleShot)
        self._ui_post_for_log = lambda fn: QTimer.singleShot(0, fn)
        self._logger = _UILogger(self._ui_post_for_log, _append_debug_line)
        set_global_debug_logger(self.debug)

        # Add console under the tree
        left_v.addWidget(self.debug_console, 0)
        
        # --- Debug console actions (copy/copy-all/save/clear)
        from PyQt5.QtWidgets import QAction

        def _dbg_copy_all():
            self.debug_console.selectAll()
            self.debug_console.copy()
            cur = self.debug_console.textCursor()
            cur.movePosition(cur.End)
            self.debug_console.setTextCursor(cur)

        def _dbg_save():
            path, _ = QFileDialog.getSaveFileName(
                self, "Save Debug Log", "debug_log.txt",
                "Text Files (*.txt);;All Files (*)"
            )
            if path:
                try:
                    Path(path).write_text(self.debug_console.toPlainText(), encoding="utf-8")
                    self.statusBar().showMessage(f"Saved log → {path}", 3000)
                except Exception as e:
                    self.error("Save Log", f"Failed to save log:\n{e}")

        # Context menu
        self.debug_console.setContextMenuPolicy(Qt.CustomContextMenu)
        def _dbg_menu(pos):
            menu = self.debug_console.createStandardContextMenu()
            menu.addSeparator()
            act_copy_all = QtAction("Copy All", menu); act_copy_all.triggered.connect(_dbg_copy_all); menu.addAction(act_copy_all)
            act_save = QtAction("Save Log…", menu); act_save.triggered.connect(_dbg_save); menu.addAction(act_save)
            act_clear = QtAction("Clear Log", menu); act_clear.triggered.connect(self.debug_console.clear); menu.addAction(act_clear)
            menu.exec_(self.debug_console.mapToGlobal(pos))
        self.debug_console.customContextMenuRequested.connect(_dbg_menu)

        # # Optional toolbar buttons
        # act_dbg_copy = QtAction("Copy", self); act_dbg_copy.setShortcut(QKeySequence.Copy); act_dbg_copy.triggered.connect(self.debug_console.copy)
        # act_dbg_copy_all = QtAction("Copy All", self); act_dbg_copy_all.setShortcut(QKeySequence("Ctrl+Shift+C")); act_dbg_copy_all.triggered.connect(_dbg_copy_all)
        # act_dbg_save = QtAction("Save Log…", self); act_dbg_save.setShortcut(QKeySequence("Ctrl+Shift+S")); act_dbg_save.triggered.connect(_dbg_save)
        # act_dbg_clear = QtAction("Clear Log", self); act_dbg_clear.setShortcut(QKeySequence("Ctrl+Shift+L")); act_dbg_clear.triggered.connect(self.debug_console.clear)
        # tb.addSeparator(); tb.addAction(act_dbg_copy); tb.addAction(act_dbg_copy_all); tb.addAction(act_dbg_save); tb.addAction(act_dbg_clear)

        # --- Footer (file counters) ---
        footer = QWidget(self)
        footer_h = QHBoxLayout(footer)
        footer_h.setContentsMargins(8, 4, 8, 4)
        footer_h.setSpacing(6)
        self.counter_label = QLabel("Files in folder: 0 | Total: 0", self)
        self.counter_label.setStyleSheet("color:#CFCFCF;")
        footer_h.addWidget(self.counter_label)
        footer_h.addStretch(1)
        left_v.addWidget(footer, 0)
        # ===== end LEFT PANE =====

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
        self.folder_summary = QPlainTextEdit(self.folder_panel); self.folder_summary.setPlaceholderText("Summary / notes for this folder…")
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

        # Editor tabs — all PLAIN text editors
        self.tabs = QTabWidget(self)

        # Metadata / subtabs
        meta_tab = QWidget(self); meta_v = QVBoxLayout(meta_tab); meta_v.setContentsMargins(0,0,0,0); meta_v.setSpacing(8)
        self.meta_inner = QTabWidget(meta_tab)

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
        rev_v.addWidget(self.rev_table); rev_v.addLayout(rev_buttons)
        self.meta_inner.addTab(rev_tab, "Revision History")

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
        var_v.addWidget(self.variant_table); var_v.addLayout(var_buttons)
        self.meta_inner.addTab(var_tab, "Variant Details")

        meta_v.addWidget(self.meta_inner)
        self.tabs.addTab(meta_tab, "Metadata")
        # --- Schematic tab (lazy PDF subtabs) ---
        self.schematic_tab = QWidget(self)
        schem_v = QVBoxLayout(self.schematic_tab); schem_v.setContentsMargins(0,0,0,0)
        self.schematic_tabs = QTabWidget(self.schematic_tab)
        schem_v.addWidget(self.schematic_tabs)
        self.tabs.addTab(self.schematic_tab, "Documents")
        # Connect once for lazy loads
        self.schematic_tabs.currentChanged.connect(self._on_schematic_tab_changed)


        # Plain text tabs
        net_tab = QWidget(self); net_v = QVBoxLayout(net_tab)
        self.netlist_edit = QPlainTextEdit(net_tab); self.netlist_edit.setPlaceholderText("(paste or type your netlist here)")
        net_v.addWidget(self.netlist_edit); self.tabs.addTab(net_tab, "Netlist")

        pl_tab = QWidget(self); pl_v = QVBoxLayout(pl_tab)
        self.partlist_edit = QPlainTextEdit(pl_tab)
        self.partlist_edit.setPlaceholderText(
            "(paste or type your partlist table here)\n"
            f"| {' | '.join(PARTLIST_HEADERS)} |\n{_divider_for(PARTLIST_HEADERS)}"
        )
        pl_v.addWidget(self.partlist_edit); self.tabs.addTab(pl_tab, "Partlist")

        self.cd_edit = QPlainTextEdit(self); self.cd_edit.setPlaceholderText("(Click or type to add content…)")
        cd_tab = QWidget(self); cd_v = QVBoxLayout(cd_tab); cd_v.addWidget(self.cd_edit); self.tabs.addTab(cd_tab, SECTION_TITLES["cd"])

        self.ct_edit = QPlainTextEdit(self); self.ct_edit.setPlaceholderText("(Click or type to add content…)")
        ct_tab = QWidget(self); ct_v = QVBoxLayout(ct_tab); ct_v.addWidget(self.ct_edit); self.tabs.addTab(ct_tab, SECTION_TITLES["ct"])

        self.da_edit = QPlainTextEdit(self); self.da_edit.setPlaceholderText("(Click or type to add content…)")
        da_tab = QWidget(self); da_v = QVBoxLayout(da_tab); da_v.addWidget(self.da_edit); self.tabs.addTab(da_tab, SECTION_TITLES["da"])

        review_tab = QWidget(self); review_v = QVBoxLayout(review_tab)
        self.review_edit = QPlainTextEdit(review_tab)
        self.review_edit.setPlaceholderText("Raw Markdown view. Edits here will be saved verbatim.")
        self.review_edit.textChanged.connect(self._on_review_changed)
        review_v.addWidget(self.review_edit); self.tabs.addTab(review_tab, "Review")

        right_layout.addWidget(self.tabs, 1)

        # After building tabs: set base titles, wire dirty + countdown autosave
        self.base_tab_titles = [self.tabs.tabText(i) for i in range(self.tabs.count())]
        self.tabs.currentChanged.connect(self.update_dirty_indicator)

        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self._tick_autosave_countdown)
        self.countdown_timer.start(1000)

        self._wire_dirty_signals()
        self.apply_monospace_font()

        splitter = QSplitter(Qt.Horizontal, self)
        splitter.addWidget(left_container); splitter.addWidget(right_container)
        splitter.setStretchFactor(0, 1); splitter.setStretchFactor(1, 2); splitter.setSizes([360, 980])

        central = QWidget(self); outer = QHBoxLayout(central); outer.setContentsMargins(8,8,8,8); outer.setSpacing(8)
        outer.addWidget(splitter); self.setCentralWidget(central)

        self.apply_dark_styles()
        self.show_file_ui(False)
        self.update_file_counter()

        # --- Git: init quickly, then pull in background ---
        if self.settings.get("git_enabled", True):
            try:
                ensure_git_repo(self.catalog_root, self.settings)
            except Exception:
                pass

            def _bg_pull():
                try:
                    br = (self.settings.get("git_branch") or "main")
                    self.debug(f"Pull: start ({br}) @ {datetime.datetime.now().isoformat(timespec='seconds')}")
                    git_pull(self.catalog_root, br)
                    self.debug("Pull: OK")
                    self._gitbg.ui(lambda: self.push_label.setText("Sync push in: —"))
                except Exception as ex:
                    self.debug(f"Pull: FAIL → {ex}")
                    self._gitbg.ui(lambda: None)

            self._gitbg.submit(_bg_pull)
        else:
            self.debug("Git: disabled")
            self.push_label.setText("Sync push in: —")
            self.git_push_remaining_s = None
        
    # ---------- Zoom helpers ----------------------------------------------------
    def _mono_font(self, pt: int) -> QFont:
        f = QFont()
        f.setStyleHint(QFont.Monospace)
        f.setFamily("Consolas" if platform.system() == "Windows" else "Monospace")
        f.setFixedPitch(True); f.setPointSize(pt)
        return f

    def _all_plain_editors(self) -> list:
        return [self.folder_summary, self.netlist_edit, self.partlist_edit, self.cd_edit, self.ct_edit, self.da_edit, self.review_edit]

    def apply_monospace_font(self):
        pt = max(6, self._base_font_pt + self._zoom_steps)
        f = self._mono_font(pt)
        for ed in self._all_plain_editors():
            ed.setFont(f)

    def adjust_zoom(self, delta_steps: int):
        self._zoom_steps = max(-6, min(18, self._zoom_steps + delta_steps))
        self.apply_monospace_font()

    def reset_zoom(self):
        self._zoom_steps = 0
        self.apply_monospace_font()

    # ---------- Settings / Tools / stats ---------------------------------------
    def open_settings_tools(self):
        dlg = SettingsToolsDialog(
            self,
            self.settings,
            do_save_cb=lambda: self.save_from_form(silent=False),
            do_archive_cb=self.archive_script_folder,
            do_export_cb=self.export_single_file_dialog_cb  # now only filename
        )
        if dlg.exec_() == QDialog.Accepted:
            self.settings = dlg.result_settings()
            save_settings(self.settings)
            if self.settings.get("git_enabled", True):
                try:
                    ensure_git_repo(self.catalog_root, self.settings)
                    def _bg_pull():
                        try:
                            git_pull(self.catalog_root, (self.settings.get("git_branch") or "main"))
                            self._gitbg.ui(lambda: self.statusBar().showMessage("Pulled latest from origin", 2500))
                        except Exception:
                            self._gitbg.ui(lambda: self.statusBar().showMessage("Pull skipped (no origin/creds).", 4000))
                    self._gitbg.submit(_bg_pull)
                except Exception:
                    pass
            else:
                # Turn off any pending push countdown and keep label static
                self.git_push_remaining_s = None
                self.push_label.setText("Sync push in: —")

            # If a file is open, rebuild schematic subtabs using the new folder
            if self.current_path and self.current_path.is_file():
                self._rebuild_schematic_tabs_from_scan()


    # ---------- Schematic (scan folder for PN_*.pdf) ---------------------------
    def _resolve_schematic_root(self, base_dir: Path) -> Path:
        """
        Resolve the schematic folder based on settings:
        - If settings["schematic_folder"] is absolute, use it.
        - Else resolve it relative to the current file's folder.
        - Default to 'docs' under the file's folder.
        """
        cfg = (self.settings.get("schematic_folder") or "").strip()
        if cfg:
            p = Path(cfg)
            if p.is_absolute():
                return p
            else:
                return (base_dir / p).resolve()
        # default
        return (base_dir / "docs").resolve()

    def _scan_schematic_pdfs(self, pn: str, base_dir: Path) -> list[tuple[str, Path]]:
        """
        Return [(label, pdf_path), ...] matching PN_*.pdf in the schematic folder.
        label = the '*' part (filename minus 'PN_' and '.pdf').
        """
        out: list[tuple[str, Path]] = []
        if not pn:
            return out
        root = self._resolve_schematic_root(base_dir)
        try:
            if not root.exists() or not root.is_dir():
                return out
            # Strict match on PN_*.pdf
            for p in sorted(root.glob(f"{pn}_*.pdf")):
                name = p.name
                # Expect PN_<label>.pdf
                label = name[len(pn) + 1 : -4] if name.lower().endswith(".pdf") else name
                label = label.strip() or "—"
                out.append((label, p.resolve()))
        except Exception:
            pass
        return out

    def _build_schematic_tabs_from_items(self, items: list[tuple[str, Path]]):
        """Rebuild schematic subtabs from a list of (label, path). Uses lazy-load."""
        # Clear old
        while self.schematic_tabs.count():
            self.schematic_tabs.removeTab(0)

        if not items:
            msg = QLabel(
                "No matching PDFs found.\n"
                "Make sure 'Part Number' is set and the schematic folder contains files named PN_*.pdf.",
                self.schematic_tab
            )
            msg.setAlignment(Qt.AlignCenter)
            holder = QWidget(self.schematic_tab); lay = QVBoxLayout(holder); lay.addWidget(msg)
            self.schematic_tabs.addTab(holder, "—")
            return

        for label, pdf_path in items:
            holder = QWidget(self.schematic_tab)
            lay = QVBoxLayout(holder); lay.setContentsMargins(0,0,0,0)
            info = QLabel(f"Lazy load: {pdf_path}\n\n(Select this tab to load the PDF.)", holder)
            info.setAlignment(Qt.AlignCenter)
            lay.addWidget(info)

            holder.setProperty("pdf_path", str(pdf_path))
            holder.setProperty("loaded", False)
            self.schematic_tabs.addTab(holder, label)

    def _rebuild_schematic_tabs_from_scan(self):
        """Convenience: rebuild schematic tabs for the currently loaded file."""
        if not self.current_path or not self.current_path.is_file():
            return
        pn = (self.field_widgets.get("Part Number").text() or "").strip()
        base_dir = self.current_path.parent
        items = self._scan_schematic_pdfs(pn, base_dir)
        self._build_schematic_tabs_from_items(items)

    def export_single_file_dialog_cb(self, out_name: str):
        # Destination is always TOP LEVEL of catalog_root
        out_path = self.catalog_root / (out_name if out_name.lower().endswith(".md") else (out_name + ".md"))
        try:
            count, skipped, total_bytes = self._export_concat_markdown(out_path)
        except Exception as e:
            self.error("Export Single File", f"Failed to export:\n{e}")
            return
        self.info(
            "Export Single File",
            f"Created:\n{out_path}\n\nConcatenated {count} file(s) (skipped {skipped}).\nSize: {total_bytes} bytes."
        )

    def _iter_repo_files(self, root: Path):
        try:
            for r, _, files in os.walk(root):
                base = Path(r)
                for f in files:
                    yield base / f
        except Exception:
            return

    def _count_repo_lines(self) -> int:
        """Count lines only in text files, skipping VCS/venv/build dirs and huge files."""
        text_exts = {
            ".md", ".txt", ".json", ".py", ".yml", ".yaml", ".ini", ".cfg", ".toml",
            ".csv", ".bat", ".ps1"
        }
        skip_dirs = {
            ".git", ".hg", ".svn", "__pycache__", ".mypy_cache", ".pytest_cache",
            ".ruff_cache", "node_modules", "build", "dist", ".venv", "venv"
        }
        max_file_bytes = 2_000_000  # skip unusually large files

        total = 0
        for root, dirnames, files in os.walk(self.catalog_root):
            # Prune directories in-place to avoid descending into them
            dirnames[:] = [
                d for d in dirnames
                if d not in skip_dirs and not d.startswith(".")  # skip hidden dirs too
            ]
            for name in files:
                p = Path(root) / name
                if p.suffix.lower() not in text_exts:
                    continue
                try:
                    st = p.stat()
                    if st.st_size > max_file_bytes:
                        continue
                    with open(p, "r", encoding="utf-8", errors="ignore") as fh:
                        for _ in fh:
                            total += 1
                except Exception:
                    # ignore unreadable files
                    continue
        return total

    def _update_stats_on_save(self, path: Path, text: str):
        files_total = self._count_md_recursive(self.catalog_root)
        lines_total_repo = self._count_repo_lines()

        stats = self.settings.get("stats", {})
        stats["files_total"] = files_total
        stats["lines_total"] = lines_total_repo
        stats["last_saved_path"] = str(path)
        stats["last_updated"] = datetime.datetime.now().isoformat(timespec="seconds")
        self.settings["stats"] = stats
        save_settings(self.settings)

    def _default_owner_for_context(self) -> str:
        s_owner = (self.settings.get("owner_default_name") or "").strip()
        if s_owner:
            return s_owner
        if self.current_folder:
            try:
                meta = self.read_folder_meta(self.current_folder)
                own = (meta.get("Owner","") or "").strip()
                return own
            except Exception:
                pass
        return ""

    # ----- Robust "next default" filename (per-prefix counter + collision check)
    def _ensure_counter_for_prefix(self, prefix: str) -> int:
        c = self.settings.get("last_file_counters", {})
        if prefix not in c:
            c[prefix] = 0
            self.settings["last_file_counters"] = c
            save_settings(self.settings)
        return c[prefix]

    def _set_counter_for_prefix(self, prefix: str, value: int):
        c = self.settings.get("last_file_counters", {})
        c[prefix] = max(0, int(value))
        self.settings["last_file_counters"] = c
        save_settings(self.settings)

    def _extract_number_from_name(self, prefix: str, name_no_ext: str):
        if not name_no_ext.startswith(prefix):
            return None
        tail = name_no_ext[len(prefix):]
        return int(tail) if tail.isdigit() else None

    def _next_default_filename(self) -> str:
        prefix = (self.settings.get("file_name_prefix") or "").strip()
        width = int(self.settings.get("number_width", 3))
        seq = self._ensure_counter_for_prefix(prefix) + 1

        base = self.selected_path() or self.catalog_root
        if base.is_file():
            base = base.parent

        while True:
            candidate = f"{prefix}{str(seq).zfill(width)}"
            target = base / f"{candidate}.md"
            if not target.exists():
                break
            seq += 1

        return f"{prefix}{str(seq).zfill(width)}"

    def _commit_counter_after_creation(self, created_path: Path):
        name_no_ext = created_path.stem
        prefix = (self.settings.get("file_name_prefix") or "").strip()
        possible_prefixes = {prefix}
        if name_no_ext and name_no_ext[0].isdigit():
            possible_prefixes.add("")
        best_prefix = None
        best_num = None
        for px in sorted(possible_prefixes, key=lambda p: len(p), reverse=True):
            n = self._extract_number_from_name(px, name_no_ext)
            if n is not None:
                best_prefix = px
                best_num = n
                break
        if best_prefix is not None and best_num is not None:
            current = self._ensure_counter_for_prefix(best_prefix)
            if best_num > current:
                self._set_counter_for_prefix(best_prefix, best_num)

    # ---------- Import / Export (catalog-wide) ---------------------------------
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
                i += 1; continue

            if current_rel:
                if ln == "```json meta":
                    block = []
                    i += 1
                    while i < n and lines[i].strip() != "```":
                        block.append(lines[i]); i += 1
                    if i < n and lines[i].strip() == "```": i += 1
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
                        block.append(lines[i]); i += 1
                    if i < n and lines[i].strip() == "```": i += 1
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
        self.update_file_counter()
        if errors:
            self.warn("Import Catalog", "Import completed with some errors:\n" + "\n".join(errors[:20]) + ("\n..." if len(errors) > 20 else ""))
        else:
            self.info("Import Catalog", "Import completed successfully.")

    def _export_concat_markdown(self, out_path: Path):
        root = self.catalog_root.resolve()
        out_path = out_path.resolve()

        md_files = []
        for folder, _, files in os.walk(root):
            folder_p = Path(folder)
            for name in files:
                if name.lower().endswith(".md"):
                    p = (folder_p / name).resolve()
                    md_files.append(p)

        md_files = sorted(md_files, key=lambda p: str(p.relative_to(root)).lower())

        included = []
        skipped = 0
        for p in md_files:
            if p == out_path:
                skipped += 1
                continue
            included.append(p)

        lines = []
        header = [
            "# Catalog — Concatenated Export",
            f"Generated: {datetime.datetime.now().isoformat(timespec='seconds')}",
            f"Root: {root}",
            "",
            "---",
            ""
        ]
        lines.extend(header)
        for p in included:
            rel = str(p.relative_to(root))
            lines.append(f"# File: {rel}")
            lines.append("")
            try:
                content = p.read_text(encoding="utf-8")
            except Exception as e:
                content = f"<!-- ERROR READING {rel}: {e} -->"
            lines.append(content.rstrip("\n"))
            lines.append("")
            lines.append("---")
            lines.append("")

        text = "\n".join(lines).rstrip() + "\n"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(text, encoding="utf-8")
        return (len(included), skipped, len(text.encode("utf-8")))

    # ---------- Tree rename handling / styling / dialogs ------------------------
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
        self.update_file_counter()

    def apply_dark_styles(self):
        self.setStyleSheet("""
            QWidget { background-color: #202225; color: #E6E6E6; }
            QToolBar { background: #1B1E20; spacing: 6px; border: 0px; }
            QToolButton { color: #E6E6E6; }
            QLabel { color: #E6E6E6; }
            QGroupBox { border: 1px solid #3A3F44; border-radius: 6px; margin-top: 12px; padding-top: 8px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #CFCFCF; }
            QLineEdit, QPlainTextEdit { background-color: #2A2D31; color: #E6E6E6; border: 1px solid #3A3F44; border-radius: 6px; padding: 6px; }
            QTreeView QLineEdit { background-color: #2A2D31; color: #E6E6E6; border: 1px solid #3A3F44; padding: 2px 4px; }
            QTreeView { background-color: #1E2124; alternate-background-color: #24272B; border: 1px solid #3A3F44; }
            QTreeView::item:selected { background: #3B4252; color: #E6E6E6; }
            QHeaderView::section { background-color: #2A2D31; color: #E6E6E6; border: 0px; padding: 6px; font-weight: 600; }
            QTabBar::tab { background: #2A2D31; color: #E6E6E6; padding: 8px 12px; margin-right: 2px; border-top-left-radius: 6px; border-top-right-radius: 6px; }
            QTabBar::tab:selected { background: #3A3F44; }
            QTabBar::tab:hover { background: #34383D; }
            QMessageBox, QInputDialog, QFileDialog { background-color: #202225; color: #E6E6E6; }
            QMessageBox QPushButton, QInputDialog QPushButton, QFileDialog QPushButton {
                background-color: #2F343A; color: #E6E6E6; border: 1px solid #444; border-radius: 6px; padding: 6px 12px;
            }
            QMessageBox QPushButton:hover, QInputDialog QPushButton:hover, QFileDialog QPushButton:hover { background-color: #3A4047; }
        """)

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
        dlg = QInputDialog(self); dlg.setWindowTitle(title); dlg.setLabelText(label); dlg.setTextValue(default)
        self._apply_dark(dlg)
        ok = dlg.exec_() == dlg.Accepted
        return (dlg.textValue(), ok)

    def ask_yes_no(self, title: str, text: str) -> bool:
        mb = QMessageBox(self); mb.setWindowTitle(title); mb.setText(text)
        mb.setIcon(QMessageBox.Question); mb.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        self._apply_dark(mb)
        return mb.exec_() == QMessageBox.Yes

    def info(self, title: str, text: str):
        mb = QMessageBox(self); mb.setWindowTitle(title); mb.setText(text)
        mb.setIcon(QMessageBox.Information); mb.setStandardButtons(QMessageBox.Ok)
        self._apply_dark(mb); mb.exec_()

    def warn(self, title: str, text: str):
        mb = QMessageBox(self); mb.setWindowTitle(title); mb.setText(text)
        mb.setIcon(QMessageBox.Warning); self._apply_dark(mb)
        mb.setStandardButtons(QMessageBox.Ok); mb.exec_()

    def error(self, title: str, text: str):
        mb = QMessageBox(self); mb.setWindowTitle(title); mb.setText(text)
        mb.setIcon(QMessageBox.Critical); mb.setStandardButtons(QMessageBox.Ok)
        self._apply_dark(mb); mb.exec_()

    # ---------- Dirty / autosave -----------------------------------------------
    def _wire_dirty_signals(self):
        for ed in [self.netlist_edit, self.partlist_edit, self.cd_edit, self.ct_edit, self.da_edit]:
            ed.textChanged.connect(self._mark_dirty)
        self.review_edit.textChanged.connect(self._mark_dirty)
        for le in self.field_widgets.values():
            le.textChanged.connect(self._mark_dirty)
        self.rev_table.itemChanged.connect(self._mark_dirty)
        self.variant_table.itemChanged.connect(self._mark_dirty)
        self.folder_title.textChanged.connect(self._mark_dirty)
        self.folder_desc.textChanged.connect(self._mark_dirty)
        self.folder_summary.textChanged.connect(self._mark_dirty)
        self.folder_owner.textChanged.connect(self._mark_dirty)
        self.folder_tags.textChanged.connect(self._mark_dirty)
        # Rebuild schematic subtabs when Part Number changes (optional)
        if "Part Number" in self.field_widgets:
            self.field_widgets["Part Number"].textChanged.connect(lambda *_: self._rebuild_schematic_tabs_from_scan())


    def _rebuild_schematic_tabs_if_visible(self):
        # Only rebuild if a file is loaded
        if not self.current_path:
            return
        current_pn = (self.field_widgets.get("Part Number").text() or "").strip()
        base_dir = self.current_path.parent
        revs = self._collect_revs_from_table()
        self._build_schematic_tabs(current_pn, revs, base_dir)

    def _on_review_changed(self):
        self.review_dirty = True
        self._mark_dirty()

    def _strip_dot(self, s: str) -> str:
        return s.lstrip("● ").strip()

    def _mark_dirty(self, *args):
        if self.suppress_dirty:
            return
        if not self.dirty:
            self.dirty = True
            self.autosave_remaining_s = self.autosave_interval_s
            self.autosave_label.setText(f"Autosave in: {self.autosave_remaining_s}s")
            self.update_dirty_indicator()

    def update_dirty_indicator(self):
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
        """Drive autosave and (when enabled) non-blocking git push countdown once per second."""

        # ---------- AUTOSAVE ----------
        if not self.current_path and not self.current_folder:
            self.autosave_label.setText("Autosave in: —")
            self.autosave_remaining_s = self.autosave_interval_s
        else:
            if not self.dirty:
                self.autosave_label.setText("Autosave in: —")
                self.autosave_remaining_s = self.autosave_interval_s
            else:
                if self.autosave_remaining_s <= 0:
                    self.save_from_form(silent=True)
                    self._reset_autosave_countdown()
                else:
                    self.autosave_label.setText(f"Autosave in: {self.autosave_remaining_s}s")
                    self.autosave_remaining_s -= 1

        # ---------- GIT PUSH COUNTDOWN ----------
        # If Git is disabled, keep the label static and skip all countdown/push logic.
        if not self.settings.get("git_enabled", True):
            self.push_label.setText("Sync push in: —")
            self.git_push_remaining_s = None
            return

        # If no countdown is active, keep the label static.
        if self.git_push_remaining_s is None:
            self.push_label.setText("Sync push in: —")
            return

        # Countdown reached zero → attempt push (unless one is already running).
        if self.git_push_remaining_s <= 0:
            if getattr(self, "_push_inflight", False):
                # A push is already running; just clear the countdown.
                self.git_push_remaining_s = None
                self.push_label.setText("Sync push in: —")
                return

            def _bg_push():
                try:
                    self._push_inflight = True
                    self._gitbg.ui(lambda: self.push_label.setText("Pushing…"))
                    branch = (self.settings.get("git_branch") or "main")
                    self.debug(f"Push: start → branch={branch}")
                    rc, out, err = _git(self.catalog_root, "push", "origin", f"HEAD:{branch}")
                    if rc == 0:
                        rc2, sha, _ = _git(self.catalog_root, "rev-parse", "--short", "HEAD")
                        sha_msg = f" ({sha})" if (rc2 == 0 and sha) else ""
                        self.debug(f"Push: ok{sha_msg}")
                        self._gitbg.ui(lambda: self.statusBar().showMessage(f"Push OK → origin{sha_msg}", 3000))
                    else:
                        msg = (err or out or "unknown error")
                        self.debug(f"Push: fail → {msg}")
                        self._gitbg.ui(lambda: self.statusBar().showMessage(f"Push failed: {msg}", 6000))
                except Exception as ex:
                    self.debug(f"Push: exception → {ex!r}")
                    self._gitbg.ui(lambda: self.statusBar().showMessage("Push failed (exception).", 5000))
                finally:
                    self._push_inflight = False
                    self._gitbg.ui(lambda: self.push_label.setText("Sync push in: —"))

            # Schedule background push and clear countdown.
            self.debug("Push: countdown reached 0 → dispatching")
            self._gitbg.submit(_bg_push)
            self.git_push_remaining_s = None
            return

        # Countdown still running
        self.push_label.setText(f"Sync push in: {self.git_push_remaining_s}s")
        self.git_push_remaining_s -= 1

    # ---------- File counter helpers -------------------------------------------
    def _count_md_recursive(self, folder: Path) -> int:
        cnt = 0
        try:
            for root, _, files in os.walk(folder):
                cnt += sum(1 for f in files if f.lower().endswith(".md"))
        except Exception:
            pass
        return cnt

    def _count_md_shallow(self, folder: Path) -> int:
        try:
            return sum(1 for f in folder.iterdir() if f.is_file() and f.suffix.lower() == ".md")
        except Exception:
            return 0

    def update_file_counter(self):
        root_total = self._count_md_recursive(self.catalog_root)
        sp = self.selected_path()
        folder = sp if (sp and sp.is_dir()) else (sp.parent if (sp and sp.is_file()) else self.catalog_root)
        in_folder = self._count_md_shallow(folder)
        self.counter_label.setText(f"Files in folder: {in_folder} | Total: {root_total}")

    # ---------- UI / selection helpers -----------------------------------------
    def show_file_ui(self, file_selected: bool):
        self.tabs.setVisible(file_selected)
        self.folder_panel.setVisible(not file_selected)  # intentional mistake for testing
        # NOTE: fixed below in load/show calls

    def _lock_non_review_tabs_if_export(self, current_file: Path | None):
        """Disable all tabs except Review if filename starts with 'catlog_' or 'catalog_'."""
        is_export = False
        if current_file and current_file.is_file():
            stem = current_file.stem.lower()
            is_export = stem.startswith("catlog_") or stem.startswith("catalog_")
        # Find Review tab index
        review_idx = None
        for i in range(self.tabs.count()):
            if self._strip_dot(self.tabs.tabText(i)).lower() == "review":
                review_idx = i
                break
        if review_idx is None:
            return
        for i in range(self.tabs.count()):
            self.tabs.setTabEnabled(i, (i == review_idx) if is_export else True)
        if is_export:
            self.tabs.setCurrentIndex(review_idx)

    def selected_source_index(self) -> QModelIndex | None:
        sel = self.tree.selectionModel().selectedIndexes()
        if not sel:
            return None
        idx = sel[0]
        if idx.column() != 0:
            idx = self.proxy.index(idx.row(), 0, idx.parent())
        return self.proxy.mapToSource(idx)

    def selected_path(self) -> Path | None:
        sidx = self.selected_source_index()
        if not sidx or not sidx.isValid():
            return None
        return Path(self.fs_model.filePath(sidx))

    def selected_paths(self) -> list:
        sel = self.tree.selectionModel().selectedIndexes()
        if not sel:
            return []
        col0 = []
        for idx in sel:
            if idx.column() != 0:
                continue
            sidx = self.proxy.mapToSource(idx)
            if sidx.isValid():
                col0.append(Path(self.fs_model.filePath(sidx)))
        return sorted(set(col0), key=lambda p: str(p).lower())

    # ---------- Metadata helpers -----------------------------------------------
    def add_rev_row(self):
        r = self.rev_table.rowCount()
        self.rev_table.insertRow(r)
        defaults = ["", today_iso(), "", self._default_owner_for_context()]
        if r == 0:
            defaults = ["-", today_iso(), "initial release", self._default_owner_for_context()]
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

    # ---------- Selection / load -----------------------------------------------
    def on_tree_selection(self, *_):
        if self.dirty:
            self.save_from_form(silent=True)
            self._reset_autosave_countdown()

        path = self.selected_path()
        if not path:
            self.update_file_counter()
            return
        if path.is_dir():
            self.current_path = None
            self.current_folder = path
            self.path_label.setText(f"Folder: {path}")
            self._toggle_panels(folder_mode=True)
            self.load_folder_meta(path)
            self.update_file_counter()
            return
        if path.is_file() and path.suffix.lower() == ".md":
            self.current_folder = None
            self._toggle_panels(folder_mode=False)
            self.load_file(path)
            self.update_file_counter()

    def _toggle_panels(self, folder_mode: bool):
        # Correct version of show_file_ui (without the intentional mistake above)
        self.tabs.setVisible(not folder_mode)
        self.folder_panel.setVisible(folder_mode)

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

        (fields, rev_rows, variant_items, netlist, partlist_text, cd, ct, da) = self.parse_markdown(text)

        self.suppress_dirty = True
        for key, _ in FIELD_ORDER:
            self.field_widgets[key].setText(fields.get(key, ""))

        self.rev_table.setRowCount(0)
        for row in rev_rows:
            rr = (row + [""] * len(REV_HEADERS))[:len(REV_HEADERS)]
            r = self.rev_table.rowCount()
            self.rev_table.insertRow(r)
            for c, val in enumerate(rr):
                self.rev_table.setItem(r, c, QTableWidgetItem(val))

        self.variant_table.setRowCount(0)
        for item in variant_items:
            r = self.variant_table.rowCount()
            self.variant_table.insertRow(r)
            self.variant_table.setItem(r, 0, QTableWidgetItem(item))

        self.netlist_edit.setPlainText(netlist)
        self.partlist_edit.setPlainText(partlist_text)
        self.cd_edit.setPlainText(cd)
        self.ct_edit.setPlainText(ct)
        self.da_edit.setPlainText(da)

        self.review_edit.blockSignals(True)
        self.review_edit.setPlainText(text)
        self.review_edit.blockSignals(False)
        self.review_dirty = False

        self.suppress_dirty = False
        self._clear_dirty()
        self._reset_autosave_countdown()
        self.proxy.refresh_desc(path)
        
        # Build schematic subtabs by scanning the configured folder for PN_*.pdf
        self._rebuild_schematic_tabs_from_scan()

        # Lock tabs if this is an export/aggregate file
        self._lock_non_review_tabs_if_export(self.current_path)

    # ---------- Parse / Build markdown -----------------------------------------
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
            out.append(lines[j]); j += 1
        while out and not out[0].strip(): out.pop(0)
        while out and not out[-1].strip(): out.pop()
        return "\n".join(out)

    def _parse_table_at(self, lines, start_idx: int):
        rows = []
        if start_idx is None:
            return rows
        j = start_idx + 1
        n = len(lines)
        # find header
        while j < n and not lines[j].strip().startswith("|"):
            if lines[j].strip().startswith("## "):
                return rows
            j += 1
        if j >= n:
            return rows
        # skip divider lines
        j += 1
        while j < n and lines[j].strip().startswith("|") and _is_md_divider_line(lines[j]):
            j += 1
        # consume body
        while j < n and lines[j].strip().startswith("|"):
            if _is_md_divider_line(lines[j]):
                j += 1; continue
            raw = lines[j].strip().strip("|")
            cells = [c.strip() for c in raw.split("|")]
            rows.append(cells); j += 1
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

        rhx = self._find_section(lines, "Revision History")
        if rhx is not None:
            j = rhx + 1
            while j < n and not lines[j].strip().startswith("|"):
                if lines[j].strip().startswith("## "):
                    j = None; break
                j += 1
            if j is not None:
                j += 1
                while j < n and lines[j].strip().startswith("|") and _is_md_divider_line(lines[j]):
                    j += 1
                while j < n and lines[j].strip().startswith("|"):
                    if _is_md_divider_line(lines[j]):
                        j += 1; continue
                    raw = lines[j].strip().strip("|")
                    cells = [c.strip() for c in raw.split("|")]
                    rev_rows.append(cells); j += 1

        vdx = self._find_section(lines, "Variant Details")
        if vdx is not None:
            variant_items = self._parse_bulleted_list(lines, vdx)

        nix = self._find_section(lines, "Netlist")
        if nix is not None:
            netlist = self._read_section_text(lines, nix)

        pix = self._find_section(lines, "Partlist")
        if pix is not None:
            tmp = self._read_section_text(lines, pix)
            partlist_text = tmp if tmp.strip() else f"| {' | '.join(PARTLIST_HEADERS)} |\n{_divider_for(PARTLIST_HEADERS)}"

        cdx = self._find_section(lines, SECTION_TITLES["cd"])
        if cdx is not None: cd = self._read_section_text(lines, cdx)
        ctx = self._find_section(lines, SECTION_TITLES["ct"])
        if ctx is not None: ct = self._read_section_text(lines, ctx)
        dax = self._find_section(lines, SECTION_TITLES["da"])
        if dax is not None: da = self._read_section_text(lines, dax)

        return fields, rev_rows, variant_items, netlist, partlist_text, cd, ct, da

    def build_markdown(self, fields: dict, rev_rows: list, variant_items: list, netlist: str, partlist_text: str, cd: str, ct: str, da: str) -> str:
        intro_lines = [
            "| Field                  | Value                     |",
            "| ---------------------- | ------------------------- |",
        ]
        for key, _ in FIELD_ORDER:
            intro_lines.append(f"| {key:<22} | {fields.get(key,'').strip()} |")

        def build_table(headers, rows):
            out = ["| " + " | ".join(headers) + " |", _divider_for(headers)]
            for r in rows:
                if _is_divider_cells([x.strip() for x in r]): continue
                rr = (r + [""] * len(headers))[:len(headers)]
                out.append("| " + " | ".join(rr) + " |")
            return "\n".join(out)

        def build_variant_list(items):
            items = [it.strip() for it in items if it.strip()]
            if not items: return "- (none)"
            return "\n".join(f"- {it}" for it in items)

        ptxt = (partlist_text or "").strip()
        if not ptxt:
            ptxt = f"| {' | '.join(PARTLIST_HEADERS)} |\n{_divider_for(PARTLIST_HEADERS)}"

        def block(h, body):
            body = (body or "").strip()
            return [f"## {h}", "", (body if body else "(Click or type to add content…)")]

        out = [
            "# Circuit Metadata", "",
            f"**Last Updated:** {today_iso()}", "",
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
    # ---------- Schematic (PDF) helpers ---------------------------------------
    def _collect_revs_from_table(self) -> list[str]:
        """Return ordered, de-duplicated list of Rev column values (non-empty)."""
        seen = set()
        revs = []
        for r in range(self.rev_table.rowCount()):
            it = self.rev_table.item(r, 0)
            rev = (it.text().strip() if it else "")
            if rev and rev not in seen:
                seen.add(rev)
                revs.append(rev)
        return revs

    def _build_schematic_tabs(self, pn: str, revs: list[str], base_dir: Path):
        """Rebuild the schematic sub-tabs based on PN and revision list."""
        # Clear old tabs
        while self.schematic_tabs.count():
            self.schematic_tabs.removeTab(0)

        if not pn:
            msg = QLabel("Part Number is empty.\nPopulate Introduction → Part Number.", self.schematic_tab)
            msg.setAlignment(Qt.AlignCenter)
            holder = QWidget(self.schematic_tab); lay = QVBoxLayout(holder); lay.addWidget(msg)
            self.schematic_tabs.addTab(holder, "—")
            return

        if not revs:
            msg = QLabel("No revisions found.\nAdd rows in Metadata → Revision History.", self.schematic_tab)
            msg.setAlignment(Qt.AlignCenter)
            holder = QWidget(self.schematic_tab); lay = QVBoxLayout(holder); lay.addWidget(msg)
            self.schematic_tabs.addTab(holder, "—")
            return

        docs_dir = base_dir / "docs"
        for rev in revs:
            pdf_name = f"{pn}_{rev}.pdf"
            pdf_path = (docs_dir / pdf_name).resolve()

            # Lazy placeholder
            holder = QWidget(self.schematic_tab)
            lay = QVBoxLayout(holder); lay.setContentsMargins(0,0,0,0)
            info = QLabel(f"Lazy load: {pdf_path}\n\n(Select this tab to load the PDF.)", holder)
            info.setAlignment(Qt.AlignCenter)
            lay.addWidget(info)

            holder.setProperty("pdf_path", str(pdf_path))
            holder.setProperty("loaded", False)
            self.schematic_tabs.addTab(holder, rev)

    def _on_schematic_tab_changed(self, idx: int):
        """Lazy-load the PDF viewer when a subtab is first selected."""
        if idx < 0:
            return

        w = self.schematic_tabs.widget(idx)
        if w is None:
            return

        if w.property("loaded"):
            return

        # 1) Try to get the path from the tab's stored property
        pdf_path = w.property("pdf_path")
        pdf_path = str(pdf_path) if pdf_path else ""

        # 2) If missing/empty, RECOMPUTE from Settings + PN + tab label
        if not pdf_path:
            try:
                label = self.schematic_tabs.tabText(idx).strip()
                pn = (self.field_widgets.get("Part Number").text() or "").strip()
                base_dir = self.current_path.parent if self.current_path else self.catalog_root
                root = self._resolve_schematic_root(base_dir)
                candidate = root / f"{pn}_{label}.pdf"
                pdf_path = str(candidate.resolve())
            except Exception:
                pdf_path = ""

        # Clear placeholder widgets
        layout = w.layout()
        for i in reversed(range(layout.count())):
            item = layout.takeAt(i)
            if item and item.widget():
                item.widget().deleteLater()

        # Always show the computed path (diagnostic)
        if not pdf_path:
            msg = QLabel(
                "PDF not found (empty path).\n\n"
                "Tip: Check Introduction → Part Number and Settings → Schematic Folder.",
                w
            )
            msg.setAlignment(Qt.AlignCenter)
            layout.addWidget(msg)
            w.setProperty("loaded", True)
            return

        # Load or show error messages with FULL path
        if not _HAS_FITZ:
            msg = QLabel(f"PyMuPDF (fitz) not installed.\n\npip install pymupdf\n\nPath:\n{pdf_path}", w)
            msg.setAlignment(Qt.AlignCenter)
            layout.addWidget(msg)
            w.setProperty("loaded", True)
            return

        if not os.path.exists(pdf_path):
            msg = QLabel(f"PDF not found:\n{pdf_path}", w)
            msg.setAlignment(Qt.AlignCenter)
            layout.addWidget(msg)
            w.setProperty("loaded", True)
            return

        # Success → render
        view = PDFGraphicsView(w)
        layout.addWidget(view)
        view.load_pdf(pdf_path)
        w.setProperty("loaded", True)

    # ---------- Save ------------------------------------------------------------
    def save_from_form(self, silent: bool = False):
        """Save current file or folder metadata, commit the change, and schedule a background push."""
        self.debug(f"Save: start → file={self.current_path} folder={self.current_folder}")

        # Decide if we should try to commit/push after saving (default yes).
        # We’ll flip this to False if the pre-save pull indicates conflicts/failure.
        allow_commit_push = True

        # Only attempt a pre-save pull when we’re saving something real (file or folder)
        something_selected = (self.current_path and self.current_path.is_file()) or (self.current_folder and self.current_folder.is_dir())
        if something_selected:
            ok_to_commit = self._safe_pull_before_save()
            if not ok_to_commit:
                allow_commit_push = False

        # ---- Save Markdown file ----
        if self.current_path and self.current_path.is_file() and self.current_path.suffix.lower() == ".md":
            # If the Review tab was edited or Review is active, save raw text verbatim
            if self.review_dirty or self._strip_dot(self.tabs.tabText(self.tabs.currentIndex())) == "Review":
                raw = self.review_edit.toPlainText()
                try:
                    self.current_path.write_text(raw, encoding="utf-8")
                except Exception as e:
                    self.error("Error", f"Failed to save file:\n{e}")
                    return

                if not silent:
                    self.load_file(self.current_path)

                self.review_dirty = False
                self._clear_dirty()
                self._reset_autosave_countdown()
                self._update_stats_on_save(self.current_path, raw)

                # Commit and schedule push (if allowed)
                if allow_commit_push and bool(self.settings.get("git_enabled", True)):
                    # Commit and schedule push (this file only)
                    try:
                        highest = _highest_rev_from_table(self.rev_table) if hasattr(self, "rev_table") else "-"
                        msg = self._commit_msg_for_file(self.current_path, highest)
                        # Stage and commit ONLY the current file
                        _git(self.catalog_root, "add", str(self.current_path))
                        rc_chk, _, _ = _git(self.catalog_root, "diff", "--cached", "--quiet", "--", str(self.current_path))
                        if rc_chk == 0:
                            # No staged change for this path → nothing to commit
                            self.statusBar().showMessage("No changes to commit.", 2000)
                            self.debug("Commit: no changes (file-specific)")
                        else:
                            self.debug(f"Commit (file): {self.current_path.name} → {msg}")
                            rc_c, _, err_c = _git(self.catalog_root, "commit", "-m", msg, "--", str(self.current_path))
                            if rc_c == 0:
                                rc_sha, sha, _ = _git(self.catalog_root, "rev-parse", "--short", "HEAD")
                                if rc_sha == 0 and sha:
                                    self.statusBar().showMessage(f"Committed {sha}: {msg}", 3000)
                                self._schedule_git_push()
                            else:
                                self.statusBar().showMessage(f"Commit failed: {err_c or 'unknown error'}", 5000)
                                self.debug(f"Commit: fail → {err_c or 'unknown error'}")
                    except Exception as ex:
                        self.debug(f"Commit: exception → {ex!r}")
                else:
                    # Either git disabled or pull failed → skip commit/push
                    if bool(self.settings.get("git_enabled", True)) and not allow_commit_push:
                        self.debug("Commit: skipped due to pre-save pull failure/conflict")

                return

            # Structured save (form → markdown)
            fields = {k: w.text().strip() for k, w in self.field_widgets.items()}

            rev_rows = []
            for r in range(self.rev_table.rowCount()):
                row = []
                for c in range(len(REV_HEADERS)):
                    it = self.rev_table.item(r, c)
                    row.append(it.text().strip() if it else "")
                # Backfill empty “By” with default owner if blank
                if len(row) >= 4 and not row[3].strip():
                    row[3] = self._default_owner_for_context()
                rev_rows.append(row)

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

            text = self.build_markdown(fields, rev_rows, variant_items, netlist, partlist_text, cd, ct, da)

            try:
                self.current_path.write_text(text, encoding="utf-8")
            except Exception as e:
                self.error("Error", f"Failed to save file:\n{e}")
                return

            self.review_edit.blockSignals(True)
            self.review_edit.setPlainText(text)
            self.review_edit.blockSignals(False)
            self.review_dirty = False
            self._clear_dirty()
            self._reset_autosave_countdown()
            self.proxy.refresh_desc(self.current_path)
            self._update_stats_on_save(self.current_path, text)

            # Commit and schedule push (if allowed)
            if allow_commit_push and bool(self.settings.get("git_enabled", True)):
                try:
                    highest = _highest_rev_from_table(self.rev_table) if hasattr(self, "rev_table") else "-"
                    msg = f"REV {highest}"
                    self.debug(f"Commit: {msg}")
                    if git_commit_all(self.catalog_root, msg):
                        rc_sha, sha, _ = _git(self.catalog_root, "rev-parse", "--short", "HEAD")
                        if rc_sha == 0 and sha:
                            self.statusBar().showMessage(f"Committed {sha}: {msg}", 3000)
                        self._schedule_git_push()
                    else:
                        self.statusBar().showMessage("No changes to commit.", 2000)
                except Exception:
                    pass
            else:
                if bool(self.settings.get("git_enabled", True)) and not allow_commit_push:
                    self.debug("Commit: skipped due to pre-save pull failure/conflict")

            return

        # ---- Save folder metadata ----
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
                self.error("Error", f"Failed to save folder metadata:\n{e}")
                return
            # Commit and schedule push (folder metadata JSON only)
            try:
                meta_p = folder_meta_path(self.current_folder)
                msg = f"Folder '{self.current_folder.name}' metadata updated"

                _git(self.catalog_root, "add", str(meta_p))
                rc_chk, _, _ = _git(self.catalog_root, "diff", "--cached", "--quiet", "--", str(meta_p))
                if rc_chk == 0:
                    self.statusBar().showMessage("No changes to commit.", 2000)
                    self.debug("Commit: no changes (folder meta)")
                else:
                    self.debug(f"Commit (folder meta): {meta_p.name} → {msg}")
                    rc_c, _, err_c = _git(self.catalog_root, "commit", "-m", msg, "--", str(meta_p))
                    if rc_c == 0:
                        rc_sha, sha, _ = _git(self.catalog_root, "rev-parse", "--short", "HEAD")
                        if rc_sha == 0 and sha:
                            self.statusBar().showMessage(f"Committed {sha}: {msg}", 3000)
                        self._schedule_git_push()
                    else:
                        self.statusBar().showMessage(f"Commit failed: {err_c or 'unknown error'}", 5000)
                        self.debug(f"Commit: fail → {err_c or 'unknown error'}")
            except Exception as ex:
                self.debug(f"Commit: exception → {ex!r}")
            else:
                if bool(self.settings.get("git_enabled", True)) and not allow_commit_push:
                    self.debug("Commit: skipped due to pre-save pull failure/conflict")

            self.proxy.refresh_desc(self.current_folder)
            self.folder_created.setText(meta["Created"])
            self.folder_updated.setText(meta["Last Updated"])
            self._clear_dirty()
            self._reset_autosave_countdown()
            if not silent:
                self.info("Saved", "Folder metadata saved.")
            return

        # ---- Nothing selected to save ----
        if not silent:
            self.info("Save", "Select a folder or a Markdown file to save.")

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
        self.update_file_counter()

    def create_new_entry(self):
        base = self.selected_path() or self.catalog_root
        if base.is_file(): base = base.parent

        suggested = self._next_default_filename()
        name, ok = self.ask_text("New Entry", "File name (without extension):", default=suggested)
        if not ok or not name.strip(): return
        safe = name.strip()
        if not safe.lower().endswith(".md"): safe += ".md"
        target = base / safe
        if target.exists():
            self.warn("Exists", "A file with that name already exists."); return

        try:
            default_owner = self._default_owner_for_context()
            target.write_text(_new_entry_template(default_owner), encoding="utf-8")
        except Exception as e:
            self.error("Error", f"Failed to create file:\n{e}"); return

        # Commit counter based on actual created filename
        self._commit_counter_after_creation(target)

        self.load_file(target)
        sidx = self.fs_model.index(str(target))
        if sidx.isValid():
            pidx = self.proxy.mapFromSource(sidx)
            if pidx.isValid(): self.tree.setCurrentIndex(pidx)
        self.update_file_counter()

    def rename_item(self):
        paths = self.selected_paths()
        if not paths:
            self.info("Rename", "Select one or more files/folders to rename.")
            return

        # Batch case
        if len(paths) >= 2:
            self.warn("Rename Disabled",
                      "Batch rename has been disabled to prevent breaking external references.")
            return
            # files = [p for p in paths if p.is_file()]
            # if not files:
            #     self.warn("Rename", "Batch rename only supports files. Select files and try again.")
            #     return

            # base_default = (self.settings.get("file_name_prefix") or "").strip()
            # base, ok = self.ask_text("Batch Rename", "Base prefix for new names (e.g., LM):", default=base_default)
            # if not ok:
            #     return
            # base = (base or "").strip()

            # start_text, ok = self.ask_text("Batch Rename", "Starting number:", default="1")
            # if not ok:
            #     return
            # try:
            #     start = max(1, int(start_text.strip()))
            # except Exception:
            #     self.warn("Rename", "Starting number must be an integer.")
            #     return

            # width_text, ok = self.ask_text("Batch Rename", "Number width (zero-padded):", default=str(int(self.settings.get("number_width", 3))))
            # if not ok:
            #     return
            # try:
            #     width = max(1, int(width_text.strip()))
            # except Exception:
            #     self.warn("Rename", "Width must be an integer.")
            #     return

            # files = sorted(files, key=lambda p: (str(p.parent).lower(), p.name.lower()))
            # i = start
            # errors = []
            # renamed = 0

            # for p in files:
            #     parent = p.parent
            #     ext = p.suffix  # keep original extension
            #     new_name = f"{base}{str(i).zfill(width)}{ext}"
            #     target = parent / new_name
            #     if target.exists():
            #         errors.append(f"Exists: {target}")
            #         i += 1
            #         continue
            #     try:
            #         p.rename(target)
            #         self.proxy.refresh_desc(target)
            #         renamed += 1
            #         i += 1
            #     except Exception as e:
            #         errors.append(f"{p.name}: {e}")

            # self.update_file_counter()
            # if errors:
            #     self.warn("Batch Rename", f"Renamed {renamed} item(s).\nSome items failed:\n" + "\n".join(errors[:20]) + ("\n..." if len(errors) > 20 else ""))
            # else:
            #     self.info("Batch Rename", f"Renamed {renamed} item(s).")
            # return

        # Single-item rename
        path = paths[0]
        new_name, ok = self.ask_text("Rename", "New name:", default=path.name)
        if not ok or not new_name.strip():
            return

        new_path = path.parent / new_name.strip()
        if new_path.exists():
            self.warn("Exists", "Target name already exists.")
            return

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

        self.update_file_counter()

    def debug(self, msg: str):
        try:
            self._logger.log(msg)
        except Exception:
            pass

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
            self.current_path = None; self.path_label.setText("")
        if self.current_folder and self.current_folder == path:
            self.current_folder = None; self.path_label.setText("")
        self.update_file_counter()

    def _schedule_git_push(self):
        if not self.settings.get("git_enabled", True):
            self.git_push_remaining_s = None
            self.push_label.setText("Sync push in: —")
            self.debug("Push: skipped (git disabled)")
            return

        try:
            delay = int(self.settings.get("auto_push_delay_seconds", 60))
        except Exception:
            delay = 60
        delay = max(1, delay)

        self.git_push_delay_s = delay
        self.git_push_remaining_s = delay
        self.push_label.setText(f"Sync push in: {self.git_push_remaining_s}s")
        self.debug(f"Push: scheduled in {self.git_push_remaining_s}s")

    def _note(self, msg: str, ms: int = 4000):
        try:
            self.statusBar().showMessage(msg, ms)
        except Exception:
            pass
    
    def _safe_pull_before_save(self) -> bool:
        """
        Try to pull latest changes before we write/commit.
        Returns True if it's safe to proceed with commit/push, False if we should skip commit.
        Never raises; logs to the debug console and status bar instead.
        """
        # Respect the setting (default True if not present)
        if not bool(self.settings.get("git_enabled", True)):
            self.debug("Pull: skipped (git disabled)")
            return True

        # Require an 'origin' remote; if not present, just proceed locally
        rc, remotes, _ = _git(self.catalog_root, "remote")
        if rc != 0 or "origin" not in (remotes or "").splitlines():
            self.debug("Pull: skipped (no origin)")
            return True

        branch = (self.settings.get("git_branch") or "main").strip() or "main"

        try:
            # Fetch first for clarity
            self.debug(f"Pull: pre-save fetch origin/{branch}")
            _git(self.catalog_root, "fetch", "origin", branch)

            # Pull with rebase; autostash helps when there are local changes
            self.debug(f"Pull: pre-save pull --rebase --autostash from origin/{branch}")
            rc, out, err = _git(self.catalog_root, "pull", "--rebase", "--autostash", "origin", branch)
            if rc == 0:
                self.debug("Pull: pre-save OK")
                return True

            # If rebase conflicts, try to abort cleanly and skip commit/push
            self.debug(f"Pull: pre-save failed; attempting rebase --abort → {err or out}")
            _git(self.catalog_root, "rebase", "--abort")
            self._gitbg.ui(lambda: self.statusBar().showMessage(
                "Remote changes conflict with your local edits. Saved locally; commit/push skipped. Resolve and save again.",
                6000
            ))
            return False

        except Exception as ex:
            self.debug(f"Pull: pre-save exception → {ex!r}")
            # Conservative: allow save, but skip commit/push
            self._gitbg.ui(lambda: self.statusBar().showMessage(
                "Pre-save pull failed. Saved locally; commit/push skipped.", 6000
            ))
            return False

    def archive_script_folder(self):
        """Zip the folder containing this script and drop the archive inside it."""
        try:
            script_dir = _script_dir()
        except Exception:
            self.error("Archive", "Could not determine script directory.")
            return

        ts = now_stamp()
        temp_base = Path(tempfile.gettempdir()) / ts

        try:
            # Create zip of the script folder (parent=root_dir, base_dir=folder name)
            shutil.make_archive(str(temp_base), 'zip',
                                root_dir=str(script_dir.parent),
                                base_dir=script_dir.name)
        except Exception as e:
            self.error("Archive", f"Failed to create archive:\n{e}")
            return

        temp_zip = Path(str(temp_base) + ".zip")
        if not temp_zip.exists():
            self.error("Archive", "Archive creation failed (file missing).")
            return

        dest_zip = script_dir / f"{ts}.zip"
        try:
            if dest_zip.exists():
                # Avoid overwrite if same-second archive already exists
                dest_zip = script_dir / f"{ts}_1.zip"
            shutil.move(str(temp_zip), str(dest_zip))
        except Exception as e:
            self.error("Archive", f"Failed to move archive into folder:\n{e}")
            return

        self.info("Archive", f"Created: {dest_zip}")

    def _commit_msg_for_file(self, file_path: Path, highest_rev: str | None) -> str:
        """
        Build a per-file commit message like 'PN-123, REV A'.
        Fallback subject:
        - '.md' and '.pdf' files → 'Document'
        - everything else → 'System File'
        """
        # Try the PN from the Introduction form
        pn = ""
        try:
            pn = (self.field_widgets.get("Part Number").text() or "").strip()
        except Exception:
            pn = ""

        if not pn:
            ext = (file_path.suffix or "").lower()
            pn = "Document" if ext in (".md", ".pdf") else "System File"

        rev = (highest_rev or "").strip() if highest_rev is not None else ""
        if rev and rev != "-":
            return f"{pn}, REV {rev}"
        return pn

# ---------- Boot ---------------------------------------------------------------
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
    # Ensure settings file exists early
    if not SETTINGS_PATH.exists():
        save_settings(load_settings())

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
