from __future__ import annotations

import ctypes
from tkinter import Tk

from app.ui.app import SpedApp


def _enable_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main() -> None:
    _enable_dpi_awareness()
    root = Tk()
    SpedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()