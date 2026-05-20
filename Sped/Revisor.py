from __future__ import annotations

from tkinter import Tk

from app.ui.app import SpedApp


def main() -> None:
    root = Tk()
    SpedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()