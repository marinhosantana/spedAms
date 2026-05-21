from __future__ import annotations

from tkinter import Tk

from app.ui_next.app import NextSpedApp


def main() -> None:
    root = Tk()
    NextSpedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
