"""Farmstead Evaluation Generator v33 runner.

Visual/UI refresh only.

This version loads the working v32 program and changes only interface styling:
- softer modern color palette
- cleaner buttons and tabs
- improved spacing/typography
- more polished cards and text boxes
- no narrative, document-generation, proofreading, or data logic changes

Run with Run_Farmstead_Eval_v33.bat.
"""

import os
import runpy
import tkinter as tk
from tkinter import ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_V32 = os.path.join(BASE_DIR, "main_v32.py")

base = runpy.run_path(BASE_V32, run_name="farmstead_generator_v32")
App = base["App"]

# Modern visual palette. These are intentionally conservative so the app still
# looks professional for CNMP/EC evaluation work.
MODERN_BG = "#f3f7f5"
MODERN_CARD = "#ffffff"
MODERN_CARD_ALT = "#f8faf9"
MODERN_BLUE = "#184e77"
MODERN_GREEN = "#2d6a4f"
MODERN_ACCENT = "#0f766e"
MODERN_ACCENT_DARK = "#0b5f59"
MODERN_TEXT = "#1f2933"
MODERN_MUTED = "#64748b"
MODERN_BORDER = "#d9e2dd"
MODERN_HEADER = "#ffffff"
MODERN_TABLE_BLUE = "#dbeafe"

# Update the globals used by inherited UI methods. This keeps the existing UI
# structure and behavior intact while refreshing the look.
for method_name in ["create_ui", "configure_styles", "add_text", "load_logo"]:
    method = getattr(App, method_name, None)
    if method and hasattr(method, "__globals__"):
        g = method.__globals__
        g["APP_BG"] = MODERN_BG
        g["APP_CARD"] = MODERN_CARD
        g["APP_BLUE"] = MODERN_BLUE
        g["APP_GREEN"] = MODERN_GREEN
        g["APP_ACCENT"] = MODERN_ACCENT
        g["APP_TEXT"] = MODERN_TEXT
        g["APP_MUTED"] = MODERN_MUTED

_ORIGINAL_CONFIGURE_STYLES = App.configure_styles
_ORIGINAL_ADD_TEXT = App.add_text
_ORIGINAL_ADD_ENTRY = App.add_entry


def modern_configure_styles(self):
    """Visual-only style refresh."""
    style = ttk.Style(self)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    # Base frames and labels.
    style.configure("TFrame", background=MODERN_BG)
    style.configure("Header.TFrame", background=MODERN_HEADER)
    style.configure("Toolbar.TFrame", background=MODERN_HEADER)
    style.configure("TLabel", background=MODERN_BG, foreground=MODERN_TEXT, font=("Segoe UI", 10))
    style.configure("HeaderTitle.TLabel", background=MODERN_HEADER, foreground=MODERN_BLUE, font=("Segoe UI Semibold", 24))
    style.configure("HeaderSub.TLabel", background=MODERN_HEADER, foreground=MODERN_MUTED, font=("Segoe UI", 10))
    style.configure("SectionTitle.TLabel", background=MODERN_BG, foreground=MODERN_GREEN, font=("Segoe UI Semibold", 18))
    style.configure("Help.TLabel", background=MODERN_BG, foreground=MODERN_MUTED, font=("Segoe UI", 9))

    # Cards / grouped areas.
    style.configure(
        "TLabelframe",
        background=MODERN_BG,
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "TLabelframe.Label",
        background=MODERN_BG,
        foreground=MODERN_BLUE,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "Card.TLabelframe",
        background=MODERN_CARD,
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=MODERN_CARD,
        foreground=MODERN_BLUE,
        font=("Segoe UI Semibold", 10),
    )

    # Notebook tabs.
    style.configure("TNotebook", background=MODERN_BG, borderwidth=0, tabmargins=(6, 5, 6, 0))
    style.configure(
        "TNotebook.Tab",
        padding=(17, 9),
        font=("Segoe UI Semibold", 9),
        background="#e8f0ec",
        foreground=MODERN_MUTED,
        borderwidth=0,
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", MODERN_CARD), ("active", "#eef6f2")],
        foreground=[("selected", MODERN_ACCENT), ("active", MODERN_BLUE)],
    )

    # Buttons.
    style.configure(
        "TButton",
        font=("Segoe UI", 9),
        padding=(10, 7),
        background="#edf2f7",
        foreground=MODERN_TEXT,
        borderwidth=1,
        relief="flat",
    )
    style.map(
        "TButton",
        background=[("active", "#e2e8f0"), ("pressed", "#cbd5e1")],
        foreground=[("active", MODERN_TEXT)],
    )
    style.configure(
        "Accent.TButton",
        font=("Segoe UI Semibold", 10),
        padding=(12, 8),
        foreground="#ffffff",
        background=MODERN_ACCENT,
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Accent.TButton",
        background=[("active", MODERN_ACCENT_DARK), ("pressed", "#084c47")],
        foreground=[("active", "#ffffff")],
    )

    # Inputs.
    style.configure(
        "TEntry",
        fieldbackground="#ffffff",
        foreground=MODERN_TEXT,
        borderwidth=1,
        relief="solid",
        padding=(5, 4),
    )
    style.configure(
        "TCombobox",
        fieldbackground="#ffffff",
        foreground=MODERN_TEXT,
        borderwidth=1,
        relief="solid",
        padding=(5, 4),
    )
    style.configure("Vertical.TScrollbar", background="#dbe5df", troughcolor=MODERN_BG, borderwidth=0)

    self.configure(background=MODERN_BG)


def modern_add_text(self, parent, label, height=5, default=""):
    """Visual-only replacement for text box creation."""
    ttk.Label(parent, text=label, font=("Segoe UI Semibold", 10), foreground=MODERN_BLUE).pack(anchor="w", pady=(10, 3))
    text = tk.Text(
        parent,
        height=height,
        wrap="word",
        undo=True,
        font=("Segoe UI", 10),
        relief="solid",
        borderwidth=1,
        padx=10,
        pady=8,
        background="#ffffff",
        foreground=MODERN_TEXT,
        insertbackground=MODERN_ACCENT,
        highlightthickness=1,
        highlightbackground=MODERN_BORDER,
        highlightcolor=MODERN_ACCENT,
    )
    text.pack(fill="x", expand=False, padx=(0, 8))
    text.insert("1.0", default)
    return text


def modern_add_entry(self, parent, label, default="", help_text=""):
    """Visual-only replacement for entry rows with cleaner spacing."""
    frame = ttk.Frame(parent)
    frame.pack(fill="x", pady=5, padx=(0, 8))
    row = ttk.Frame(frame)
    row.pack(fill="x")
    ttk.Label(row, text=label, width=34, foreground=MODERN_TEXT).pack(side="left")
    entry = ttk.Entry(row)
    entry.pack(side="left", fill="x", expand=True, ipady=2)
    entry.insert(0, default)
    if help_text:
        ttk.Label(
            frame,
            text="What to enter: " + help_text,
            style="Help.TLabel",
            wraplength=1000,
            justify="left",
        ).pack(anchor="w", padx=(240, 0), pady=(2, 3))
    return entry


App.configure_styles = modern_configure_styles
App.add_text = modern_add_text
App.add_entry = modern_add_entry


if __name__ == "__main__":
    App().mainloop()
