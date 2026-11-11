#!/usr/bin/env python3
"""
VDS Legacy Excel Converter (PySide6/Qt build)
==============================================
Same tool as the Tkinter build - pick a folder, convert every legacy
.xls/.xlsm/.xlt/.xlsb/.xlc inside it to modern .xlsx via Excel COM.
Only the UI toolkit differs; see VDS_Xls_To_Xlsx.py's module docstring
for the full feature description.
"""

# --- 3SVerse runtime tamper guard (t197): frozen EXE main debugger = silent exit ---
import runtime_guard as _rguard3sv
_rguard3sv.enforce()

import sys


def _pkg_version(dist):
    try:
        import importlib.metadata as _md
        return _md.version(dist)
    except Exception:
        return None


def _ensure(pip_name, imp_name):
    if _pkg_version(pip_name) is not None:
        return
    try:
        import importlib.util as _iu
        if _iu.find_spec(imp_name.split(".")[0]) is None:
            raise ImportError(imp_name)
    except ImportError:
        print(f"[WARN] {pip_name} is missing; install it before using this feature.")


for _p, _i in [("pywin32", "win32com")]:
    _ensure(_p, _i)

import os, time, threading, queue, traceback
from datetime import datetime

from theme_manager_qt import ThemeManagerQt

# Excel file-format codes
XLSX = 51        # xlOpenXMLWorkbook
LEGACY_EXTS = (".xls", ".xlsm", ".xlt", ".xlsb", ".xlc")   # convert these -> .xlsx
RESTART_EVERY = 60   # restart Excel periodically to avoid memory bloat on big batches

_CANCEL = threading.Event()


def _find_files(folder, recurse):
    out = []
    if recurse:
        for root, _, files in os.walk(folder):
            for f in files:
                out.append(os.path.join(root, f))
    else:
        out = [os.path.join(folder, f) for f in os.listdir(folder)]
    res = []
    for p in out:
        base = os.path.basename(p)
        if base.startswith("~$"):
            continue  # Excel lock files
        if os.path.splitext(base)[1].lower() in LEGACY_EXTS:
            res.append(p)
    return sorted(res)


class Converter:
    def __init__(self, log):
        self.log = log
        self.xl = None
        self._opened = 0

    def _start_excel(self):
        import win32com.client  # LAZY: heavy COM import kept off the startup path
        self.xl = win32com.client.DispatchEx("Excel.Application")
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        try:
            self.xl.AutomationSecurity = 3   # block macros from prompting
        except Exception:
            pass
        try:
            self.xl.AskToUpdateLinks = False
        except Exception:
            pass

    def _stop_excel(self):
        if self.xl is not None:
            try:
                self.xl.Quit()
            except Exception:
                pass
        self.xl = None

    def _recycle_if_needed(self):
        self._opened += 1
        if self._opened % RESTART_EVERY == 0:
            self._stop_excel()
            time.sleep(1)
            self._start_excel()

    def convert_one(self, path, overwrite, delete_original):
        out = os.path.splitext(path)[0] + ".xlsx"
        if os.path.exists(out) and not overwrite:
            return "skip"
        wb = None
        try:
            try:
                wb = self.xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=True)
            except Exception:
                wb = self.xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0,
                                             ReadOnly=True, CorruptLoad=1)
            wb.SaveAs(os.path.abspath(out), FileFormat=XLSX)
            wb.Close(False)
            wb = None
            self._recycle_if_needed()
            if delete_original:
                try:
                    os.remove(path)
                except Exception as e:
                    self.log(f"      (kept original — delete failed: {e})", "warning")
            return "ok"
        except Exception as e:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            try:
                self._stop_excel()
                self._start_excel()
            except Exception:
                pass
            return f"error: {e}"

    def run(self, files, overwrite, delete_original):
        self._start_excel()
        ok = skip = err = 0
        try:
            for i, p in enumerate(files, 1):
                if _CANCEL.is_set():
                    self.log("  Cancelled by user.", "warning")
                    break
                name = os.path.relpath(p, os.path.commonpath(files)) if len(files) > 1 else os.path.basename(p)
                r = self.convert_one(p, overwrite, delete_original)
                if r == "ok":
                    ok += 1
                    self.log(f"  [{i}/{len(files)}] OK {name}")
                elif r == "skip":
                    skip += 1
                    self.log(f"  [{i}/{len(files)}] SKIP {name} (xlsx exists)")
                else:
                    err += 1
                    self.log(f"  [{i}/{len(files)}] FAIL {name} — {r}", "error")
        finally:
            self._stop_excel()
        return ok, skip, err


# ==========================================================
# QT UI LAYER
# ==========================================================
from PySide6.QtCore import Qt, QObject, Signal, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit,
    QPushButton, QCheckBox, QProgressBar, QPlainTextEdit, QFileDialog, QMessageBox
)

from header_manager_qt import FixedHeaderManagerQt

EMBEDDED_ICON_B64 = (
    open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_icon_b64.txt"), "r").read().strip()
    if not getattr(sys, "frozen", False)
    else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_icon_b64.txt"), "r").read().strip()
)
ICON_ICO_NAME = "VDS_icon.ico"
LOGO_PNG_NAME = "VDS_Logo.png"
APP_TITLE = "VDS Desk Solutions - Legacy Excel Converter"

LOG_COLORS = {
    "info": "#90CDF4",
    "success": "#68D391",
    "error": "#FC8181",
    "warning": "#F6E05E",
    "": "#a8d8ff",
}


def _extract_embedded_icon(b64, filename):
    try:
        if not b64:
            return None
        import base64 as _b64, tempfile
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None


def _set_window_icon_qt(window):
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        ico_path = os.path.join(meipass, ICON_ICO_NAME)
        if os.path.exists(ico_path):
            window.setWindowIcon(QIcon(ico_path))
            return
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, ICON_ICO_NAME)
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    try:
        ico_path = _extract_embedded_icon(EMBEDDED_ICON_B64, "VDS_app_icon.ico")
        if ico_path:
            window.setWindowIcon(QIcon(ico_path))
    except Exception:
        pass


class LogBridge(QObject):
    """Thread-safe log delivery (converter worker thread -> GUI thread)."""
    message = Signal(str, str)   # text, tag
    progress = Signal(int, int)  # value, maximum
    finished = Signal()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        _set_window_icon_qt(self)
        self._busy = False

        self.theme_manager = ThemeManagerQt(default="dark", app_name="VDS_Xls_To_Xlsx")

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header_mgr = FixedHeaderManagerQt(outer, title=APP_TITLE)
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_PNG_NAME)
        if os.path.exists(logo_path):
            self.header_mgr.set_logo(logo_path=logo_path, text="VDS")
        self.header_mgr.add_theme_toggle(self.theme_manager, callback=self._apply_theme)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 18, 24, 18)
        outer.addWidget(body, 1)

        folder_row = QHBoxLayout()
        self.folder_edit = QLineEdit()
        folder_row.addWidget(self.folder_edit, 1)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self._browse)
        folder_row.addWidget(browse_btn)
        body_layout.addLayout(folder_row)

        opt_row = QHBoxLayout()
        self.recurse_cb = QCheckBox("Include subfolders")
        self.recurse_cb.setChecked(True)
        self.overwrite_cb = QCheckBox("Overwrite existing .xlsx")
        self.overwrite_cb.setChecked(True)
        self.delete_cb = QCheckBox("Delete original after converting")
        self.delete_cb.setChecked(True)
        for cb in (self.recurse_cb, self.overwrite_cb, self.delete_cb):
            opt_row.addWidget(cb)
        opt_row.addStretch(1)
        body_layout.addLayout(opt_row)

        act_row = QHBoxLayout()
        self.run_btn = QPushButton("Convert files")
        self.run_btn.clicked.connect(self._start)
        act_row.addWidget(self.run_btn)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(lambda: _CANCEL.set())
        self.cancel_btn.setEnabled(False)
        act_row.addWidget(self.cancel_btn)
        self.progress = QProgressBar()
        act_row.addWidget(self.progress, 1)
        body_layout.addLayout(act_row)

        body_layout.addWidget(self._label("Activity Log", bold=True))
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background: #10182e; color: #a8d8ff; font-family: Consolas;")
        body_layout.addWidget(self.log_box, 1)

        self.header_mgr.add_copyright(self.theme_manager, outer)

        self._bridge = LogBridge()
        self._bridge.message.connect(self._append_log)
        self._bridge.progress.connect(self._set_progress)
        self._bridge.finished.connect(self._on_done)

        self._apply_theme()
        self.resize(1100, 720)
        QTimer.singleShot(0, self.showMaximized)

    @staticmethod
    def _label(text, bold=False):
        from PySide6.QtWidgets import QLabel
        lbl = QLabel(text)
        if bold:
            lbl.setStyleSheet("font-weight: bold;")
        return lbl

    def _apply_theme(self):
        self.theme_manager.apply_theme_to_app(QApplication.instance())
        self.header_mgr.update_for_theme()

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "Select folder with legacy Excel files")
        if d:
            self.folder_edit.setText(d)

    def _log(self, msg, tag=""):
        self._bridge.message.emit(msg, tag)

    def _append_log(self, msg, tag):
        color = LOG_COLORS.get(tag, LOG_COLORS[""])
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.appendHtml(f'<span style="color:{color}">[{ts}]  {msg}</span>')

    def _set_progress(self, value, maximum):
        self.progress.setMaximum(maximum)
        self.progress.setValue(value)

    def _start(self):
        if self._busy:
            return
        folder = self.folder_edit.text().strip()
        if not folder or not os.path.isdir(folder):
            QMessageBox.critical(self, "No folder", "Please browse to a valid folder first.")
            return
        _CANCEL.clear()
        self._busy = True
        self.cancel_btn.setEnabled(True)
        threading.Thread(target=self._run, args=(folder,), daemon=True).start()

    def _run(self, folder):
        log = self._log
        log("=" * 54)
        log(f"Converter started {datetime.now():%Y-%m-%d %H:%M:%S}")
        log(f"Folder: {folder}")
        log("=" * 54)
        try:
            files = _find_files(folder, self.recurse_cb.isChecked())
            log(f"  Legacy Excel files found: {len(files)}", "info")
            if not files:
                log("  Nothing to convert.", "warning")
                self._bridge.finished.emit()
                return
            self._bridge.progress.emit(0, len(files))
            conv = Converter(log)
            t0 = time.time()
            ok = skip = err = 0
            conv._start_excel()
            try:
                base = os.path.commonpath(files) if len(files) > 1 else folder
                for i, p in enumerate(files, 1):
                    if _CANCEL.is_set():
                        log("  Cancelled.", "warning")
                        break
                    name = os.path.relpath(p, base)
                    r = conv.convert_one(p, self.overwrite_cb.isChecked(), self.delete_cb.isChecked())
                    if r == "ok":
                        ok += 1
                        log(f"  [{i}/{len(files)}] OK {name}")
                    elif r == "skip":
                        skip += 1
                        log(f"  [{i}/{len(files)}] SKIP {name} (xlsx exists)")
                    else:
                        err += 1
                        log(f"  [{i}/{len(files)}] FAIL {name} — {r}", "error")
                    self._bridge.progress.emit(i, len(files))
            finally:
                conv._stop_excel()
            log("\n" + "=" * 54)
            log(f"  Done in {time.time() - t0:,.0f}s — converted {ok}, skipped {skip}, failed {err}.",
                "success" if err == 0 else "warning")
            log("=" * 54)
        except Exception as e:
            log(f"[FATAL] {e}", "error")
            log(traceback.format_exc(), "error")
        self._bridge.finished.emit()

    def _on_done(self):
        self._busy = False
        self.cancel_btn.setEnabled(False)


# ==========================================================
# ENTRY POINT
# ==========================================================
def _enable_dpi_awareness():
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _close_pyi_splash():
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("VDSDesk.LegacyExcelConverter")
    except Exception:
        pass
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    _close_pyi_splash()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
