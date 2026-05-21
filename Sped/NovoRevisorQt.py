from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from app.ui_qt.app import QtSpedApp


def main() -> None:
    app = QApplication(sys.argv)
    window = QtSpedApp()
    window.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
