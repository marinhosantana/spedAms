from __future__ import annotations

from tkinter import Tk, ttk


COLORS = {
    "app": "#eef3f7",
    "sidebar": "#17212b",
    "sidebar_active": "#25384a",
    "sidebar_text": "#e9f0f7",
    "panel": "#ffffff",
    "border": "#cfd9e3",
    "text": "#1e2a35",
    "muted": "#607282",
    "accent": "#2f80ed",
    "accent_hover": "#256fce",
    "success": "#1f8f5f",
}


def configure_theme(root: Tk) -> None:
    style = ttk.Style()
    style.theme_use("clam")
    root.configure(bg=COLORS["app"])

    style.configure("Next.TFrame", background=COLORS["app"])
    style.configure("Next.Panel.TFrame", background=COLORS["panel"], borderwidth=1, relief="solid")
    style.configure("Next.Sidebar.TFrame", background=COLORS["sidebar"])
    style.configure("Next.TLabel", background=COLORS["app"], foreground=COLORS["text"], font=("Segoe UI", 10))
    style.configure("Next.Panel.TLabel", background=COLORS["panel"], foreground=COLORS["text"], font=("Segoe UI", 10))
    style.configure("Next.Title.TLabel", background=COLORS["app"], foreground=COLORS["text"], font=("Segoe UI", 16, "bold"))
    style.configure("Next.Sidebar.TLabel", background=COLORS["sidebar"], foreground=COLORS["sidebar_text"], font=("Segoe UI", 14, "bold"))
    style.configure("Next.Muted.TLabel", background=COLORS["app"], foreground=COLORS["muted"], font=("Segoe UI", 9))
    style.configure("Next.PanelMuted.TLabel", background=COLORS["panel"], foreground=COLORS["muted"], font=("Segoe UI", 9))
    style.configure("Next.TButton", padding=(12, 7), font=("Segoe UI", 9))
    style.configure("Next.Primary.TButton", background=COLORS["accent"], foreground="#ffffff", bordercolor=COLORS["accent"], padding=(12, 7))
    style.map("Next.Primary.TButton", background=[("active", COLORS["accent_hover"])], foreground=[("active", "#ffffff")])
    style.configure("Next.Sidebar.TButton", background=COLORS["sidebar"], foreground=COLORS["sidebar_text"], borderwidth=0, anchor="w", padding=(14, 10))
    style.map("Next.Sidebar.TButton", background=[("active", COLORS["sidebar_active"])], foreground=[("active", COLORS["sidebar_text"])])
    style.configure("Treeview", rowheight=25, font=("Segoe UI", 9), background="#ffffff", fieldbackground="#ffffff", foreground=COLORS["text"])
    style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#dde7f0", foreground=COLORS["text"])
