from __future__ import annotations

import faulthandler
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from app.ui_qt.app import QtSpedApp


def main() -> None:
    crash_log = Path(__file__).resolve().parent / "qt_crash.log"
    try:
        crash_file = crash_log.open("a", encoding="utf-8")
        faulthandler.enable(crash_file)
    except Exception:
        crash_file = None
    app = QApplication(sys.argv)
    app.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    window = QtSpedApp()
    window.showMaximized()
    try:
        sys.exit(app.exec())
    finally:
        if crash_file is not None:
            crash_file.close()


if __name__ == "__main__":
    main()
