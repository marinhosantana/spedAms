from __future__ import annotations

from tkinter import StringVar, Tk, Toplevel, ttk


class ProgressDialogHandle:
    def __init__(
        self,
        root: Tk,
        dialog: Toplevel,
        message_var: StringVar,
        percent_var: StringVar,
        progress_bar: ttk.Progressbar,
    ) -> None:
        self.root = root
        self.dialog = dialog
        self.message_var = message_var
        self.percent_var = percent_var
        self.progress_bar = progress_bar
        self._closed = False

    def update(self, current: int, total: int, message: str) -> None:
        total_safe = max(total, 1)
        percentage = int((max(0, min(current, total_safe)) * 100) / total_safe)
        self.message_var.set(message)
        self.percent_var.set(f"{percentage}%")
        self.progress_bar["value"] = percentage
        self.dialog.update_idletasks()
        self.root.update_idletasks()

    def show_now(self) -> None:
        self.dialog.deiconify()
        self.dialog.lift()
        try:
            self.dialog.focus_force()
        except Exception:
            pass
        self.dialog.update_idletasks()
        self.dialog.update()

    def reset(self, message: str = "") -> None:
        if message:
            self.message_var.set(message)
        self.percent_var.set("0%")
        self.progress_bar["value"] = 0
        self.dialog.update_idletasks()
        self.root.update_idletasks()

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.dialog.grab_release()
        except Exception:
            pass
        try:
            if self.dialog.winfo_exists():
                self.dialog.destroy()
        except Exception:
            pass
