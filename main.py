"""
app.py — UI for the Stock Transfer Calculator.
All business logic lives in logic.py — edit that file to change calculations.

Run:  python app.py
Deps: pip install customtkinter openpyxl pandas
"""

import customtkinter as ctk
import threading
import os
import sys
from tkinter import filedialog, messagebox
from edit_outputfile import preview_output

# ── Import your logic ──────────────────────────────────────────────────────
from logic import run_transfer

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Palette ────────────────────────────────────────────────────────────────
C = {
    "bg":       "#0b0e1a",
    "surface":  "#12162a",
    "card":     "#181d35",
    "card2":    "#1e2440",
    "accent":   "#3d7fff",
    "accent_h": "#2563eb",
    "green":    "#10b981",
    "amber":    "#f59e0b",
    "red":      "#ef4444",
    "text":     "#e2e8f0",
    "muted":    "#64748b",
    "border":   "#252d4a",
    "border2":  "#2e3a5e",
    "white":    "#ffffff",
}


FONT_MONO = "Courier New"
FONT_UI = "Segoe UI" if sys.platform == "win32" else "Helvetica Neue"


# ──────────────────────────────────────────────────────────────────────────
# Collapsible Section
# ──────────────────────────────────────────────────────────────────────────
class Section(ctk.CTkFrame):
    def __init__(self, master, title: str, icon: str = "●", **kw):
        super().__init__(master, fg_color=C["card"], corner_radius=14,
                         border_color=C["border"], border_width=1, **kw)
        self._expanded = True

        self.header = ctk.CTkFrame(
            self, fg_color="transparent", cursor="hand2")
        self.header.pack(fill="x")
        self.header.bind("<Button-1>", self._toggle)

        left = ctk.CTkFrame(self.header, fg_color="transparent")
        left.pack(side="left", padx=18, pady=14)
        left.bind("<Button-1>", self._toggle)

        self.icon_lbl = ctk.CTkLabel(left, text=icon, font=(FONT_MONO, 15),
                                     text_color=C["accent"], width=26)
        self.icon_lbl.pack(side="left", padx=(0, 8))
        self.icon_lbl.bind("<Button-1>", self._toggle)

        self.title_lbl = ctk.CTkLabel(left, text=title,
                                      font=(FONT_UI, 13, "bold"),
                                      text_color=C["text"])
        self.title_lbl.pack(side="left")
        self.title_lbl.bind("<Button-1>", self._toggle)

        self.chevron = ctk.CTkLabel(self.header, text="▾",
                                    font=(FONT_MONO, 13), text_color=C["muted"])
        self.chevron.pack(side="right", padx=18)
        self.chevron.bind("<Button-1>", self._toggle)

        self.divider = ctk.CTkFrame(self, height=1, fg_color=C["border"])
        self.divider.pack(fill="x")

        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=18, pady=16)

    def _toggle(self, _=None):
        self._expanded = not self._expanded
        if self._expanded:
            self.divider.pack(fill="x")
            self.body.pack(fill="both", expand=True, padx=18, pady=16)
            self.chevron.configure(text="▾")
        else:
            self.divider.pack_forget()
            self.body.pack_forget()
            self.chevron.configure(text="▸")


# ──────────────────────────────────────────────────────────────────────────
# File Picker Row
# ──────────────────────────────────────────────────────────────────────────
class FileRow(ctk.CTkFrame):
    def __init__(self, master, label: str, color: str = None, **kw):
        super().__init__(master, fg_color=C["card2"], corner_radius=10,
                         border_color=C["border2"], border_width=1, **kw)
        self._path = ""
        color = color or C["accent"]

        ctk.CTkLabel(self, text=label, font=(FONT_MONO, 10, "bold"),
                     fg_color=color, corner_radius=6,
                     width=72, height=26, text_color=C["white"]).pack(
            side="left", padx=(12, 10), pady=12)

        self.dot = ctk.CTkLabel(self, text="○", font=(FONT_MONO, 12),
                                text_color=C["muted"])
        self.dot.pack(side="left", padx=(0, 6))

        self.path_lbl = ctk.CTkLabel(self, text="No file selected",
                                     text_color=C["muted"],
                                     font=(FONT_MONO, 11), anchor="w")
        self.path_lbl.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(self, text="Browse ↗", width=90, height=30,
                      fg_color="transparent", hover_color=C["border2"],
                      border_color=C["border2"], border_width=1,
                      font=(FONT_MONO, 10), text_color=C["text"],
                      command=self._browse).pack(side="right", padx=12, pady=12)

    def _browse(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if fp:
            self._path = fp
            self.path_lbl.configure(
                text=os.path.basename(fp), text_color=C["text"])
            self.dot.configure(text="●", text_color=C["green"])

    def get(self) -> str:
        return self._path

    def is_set(self) -> bool:
        return bool(self._path)


# ──────────────────────────────────────────────────────────────────────────
# Stat Card
# ──────────────────────────────────────────────────────────────────────────
class StatCard(ctk.CTkFrame):
    def __init__(self, master, label: str, **kw):
        super().__init__(master, fg_color=C["card2"], corner_radius=10,
                         border_color=C["border2"], border_width=1, **kw)
        self.val_lbl = ctk.CTkLabel(self, text="—",
                                    font=(FONT_MONO, 22, "bold"),
                                    text_color=C["accent"])
        self.val_lbl.pack(pady=(14, 2))
        ctk.CTkLabel(self, text=label, font=(FONT_UI, 10),
                     text_color=C["muted"]).pack(pady=(0, 14))

    def set(self, val, color=None):
        self.val_lbl.configure(text=str(val),
                               text_color=color or C["accent"])


# ──────────────────────────────────────────────────────────────────────────
# Main Application
# ──────────────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("StockShift — Transfer Calculator")
        self.geometry("900x820")
        self.minsize(720, 600)
        self.configure(fg_color=C["bg"])
        self._build()

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._sidebar()
        self._main()

    # ── Sidebar ───────────────────────────────────────────────────────────

    def _sidebar(self):
        sb = ctk.CTkFrame(
            self, fg_color=C["surface"], corner_radius=0, width=230)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_rowconfigure(9, weight=1)
        sb.grid_columnconfigure(0, weight=1)

        logo_frame = ctk.CTkFrame(sb, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(28, 8), sticky="w")
        ctk.CTkLabel(logo_frame, text="⇄", font=(FONT_MONO, 30, "bold"),
                     text_color=C["accent"]).pack(side="left", padx=(0, 12))
        txt = ctk.CTkFrame(logo_frame, fg_color="transparent")
        txt.pack(side="left")
        ctk.CTkLabel(txt, text="StockShift", font=(FONT_UI, 16, "bold"),
                     text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(txt, text="Transfer Calculator", font=(FONT_UI, 9),
                     text_color=C["muted"]).pack(anchor="w")

        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).grid(
            row=1, column=0, padx=16, pady=12, sticky="ew")

        nav = [
            ("📂  Input Files",   True),
            ("⚙   Configuration", False),
            ("▶   Run",           False),
            ("📊  Results",       False),
            ("🗒   Activity Log",  False),
        ]
        for i, (label, active) in enumerate(nav):
            ctk.CTkButton(
                sb, text=label, anchor="w", height=40,
                fg_color=C["accent"] if active else "transparent",
                hover_color=C["border"], corner_radius=8,
                font=(FONT_UI, 12),
                text_color=C["white"] if active else C["muted"]
            ).grid(row=i + 2, column=0, padx=12, pady=2, sticky="ew")

        ctk.CTkFrame(sb, height=1, fg_color=C["border"]).grid(
            row=9, column=0, padx=16, pady=8, sticky="ew")

        self.status_lbl = ctk.CTkLabel(sb, text="● Idle",
                                       font=(FONT_MONO, 10),
                                       text_color=C["muted"])
        self.status_lbl.grid(row=10, column=0, padx=20,
                             pady=(0, 24), sticky="w")

    # ── Main scroll area ──────────────────────────────────────────────────

    def _main(self):
        scroll = ctk.CTkScrollableFrame(
            self, fg_color=C["bg"],
            scrollbar_button_color=C["border"],
            scrollbar_button_hover_color=C["border2"])
        scroll.grid(row=0, column=1, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        P = {"padx": 24, "pady": 8, "sticky": "ew"}

        # Page header
        hdr = ctk.CTkFrame(scroll, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=24, pady=(24, 6), sticky="ew")
        ctk.CTkLabel(hdr, text="Transfer Calculator",
                     font=(FONT_UI, 22, "bold"),
                     text_color=C["text"]).pack(side="left")
        ctk.CTkLabel(hdr, text=" v2.0 ", font=(FONT_MONO, 9),
                     text_color=C["muted"], fg_color=C["border"],
                     corner_radius=4).pack(side="left", padx=10, pady=6)

        # ── 1. Input Files ──
        s1 = Section(scroll, "Input Files", icon="📂")
        s1.grid(row=1, column=0, **P)

        # Add note about automatic cleaning
        note_frame = ctk.CTkFrame(s1.body, fg_color=C["bg"], corner_radius=8)
        note_frame.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(note_frame, text="ℹ️  Sales file will be automatically cleaned (headers fixed)",
                     font=(FONT_UI, 11), text_color=C["muted"],
                     anchor="w").pack(padx=12, pady=8)

        self.row_source = FileRow(s1.body, "Sales (Source)", color=C["accent"])
        self.row_source.pack(fill="x", pady=(0, 8))
        self.row_d1 = FileRow(
            s1.body, "From Stock (Samakhosi)", color="#7c3aed")
        self.row_d1.pack(fill="x", pady=(0, 8))
        self.row_d2 = FileRow(
            s1.body, "To Stock (Destination)", color="#0891b2")
        self.row_d2.pack(fill="x")

        # ── 2. Configuration ──
        s2 = Section(scroll, "Configuration", icon="⚙")
        s2.grid(row=2, column=0, **P)

        cfg = ctk.CTkFrame(s2.body, fg_color="transparent")
        cfg.pack(fill="x")
        cfg.grid_columnconfigure((0, 1), weight=1)

        # From Location Name (Display only - for information)
        from_frame = ctk.CTkFrame(cfg, fg_color=C["card2"], corner_radius=10,
                                  border_color=C["border2"], border_width=1)
        from_frame.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        ctk.CTkLabel(from_frame, text="From Location (Fixed)",
                     font=(FONT_UI, 10, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=14, pady=(12, 4))
        ctk.CTkLabel(from_frame, text="Samakhosi (default)",
                     font=(FONT_MONO, 13), text_color=C["text"],
                     anchor="w").pack(fill="x", padx=14, pady=(0, 14))

        # To Location Name
        to_frame = ctk.CTkFrame(cfg, fg_color=C["card2"], corner_radius=10,
                                border_color=C["border2"], border_width=1)
        to_frame.grid(row=0, column=1, padx=(8, 0), sticky="ew")
        ctk.CTkLabel(to_frame, text="To Location Name",
                     font=(FONT_UI, 10, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=14, pady=(12, 4))
        self.to_name_var = ctk.StringVar()
        ctk.CTkEntry(to_frame, textvariable=self.to_name_var,
                     placeholder_text="e.g. Bhaktapur",
                     font=(FONT_MONO, 13), fg_color=C["bg"],
                     border_color=C["border2"], text_color=C["text"],
                     height=36).pack(fill="x", padx=14, pady=(0, 14))

        # Output File
        out_frame = ctk.CTkFrame(cfg, fg_color=C["card2"], corner_radius=10,
                                 border_color=C["border2"], border_width=1)
        out_frame.grid(row=1, column=0, columnspan=2,
                       pady=(12, 0), sticky="ew")
        ctk.CTkLabel(out_frame, text="Output File",
                     font=(FONT_UI, 10, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=14, pady=(12, 4))
        out_row = ctk.CTkFrame(out_frame, fg_color="transparent")
        out_row.pack(fill="x", padx=14, pady=(0, 14))
        self.out_var = ctk.StringVar(value="transfer_output.xlsx")
        ctk.CTkEntry(out_row, textvariable=self.out_var,
                     font=(FONT_MONO, 11), fg_color=C["bg"],
                     border_color=C["border2"], text_color=C["text"],
                     height=36).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(out_row, text="…", width=36, height=36,
                      fg_color=C["border"], hover_color=C["border2"],
                      font=(FONT_MONO, 14),
                      command=self._choose_out).pack(side="right")

        # ── 3. Run ──
        s3 = Section(scroll, "Run", icon="▶")
        s3.grid(row=3, column=0, **P)

        self.run_btn = ctk.CTkButton(
            s3.body, text="  Run Transfer Calculation  →",
            height=52, font=(FONT_UI, 14, "bold"),
            fg_color=C["accent"], hover_color=C["accent_h"],
            corner_radius=10, command=self._run_threaded)
        self.run_btn.pack(fill="x")

        self.progress = ctk.CTkProgressBar(
            s3.body, height=4,
            fg_color=C["border"], progress_color=C["accent"])
        self.progress.pack(fill="x", pady=(10, 0))
        self.progress.set(0)

        # ── 4. Results ──
        s4 = Section(scroll, "Results", icon="📊")
        s4.grid(row=4, column=0, **P)

        stat_row = ctk.CTkFrame(s4.body, fg_color="transparent")
        stat_row.pack(fill="x")
        stat_row.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.stat_total = StatCard(stat_row, "Total Products")
        self.stat_xfer = StatCard(stat_row, "To Transfer")
        self.stat_skip = StatCard(stat_row, "Skipped")
        self.stat_total.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        self.stat_xfer.grid(row=0, column=2, padx=2,      sticky="ew")
        self.stat_skip.grid(row=0, column=3, padx=(5, 0), sticky="ew")

        # ── 5. Log ──
        s5 = Section(scroll, "Activity Log", icon="🗒")
        s5.grid(row=5, column=0, padx=24, pady=(8, 28), sticky="ew")

        toolbar = ctk.CTkFrame(s5.body, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(toolbar, text="Console Output",
                     font=(FONT_UI, 10), text_color=C["muted"]).pack(side="left")
        ctk.CTkButton(toolbar, text="Clear", width=54, height=24,
                      fg_color="transparent", hover_color=C["border"],
                      font=(FONT_MONO, 9), text_color=C["muted"],
                      border_color=C["border"], border_width=1,
                      command=self._clear_log).pack(side="right")

        self.log = ctk.CTkTextbox(
            s5.body, height=190,
            fg_color=C["bg"], text_color=C["text"],
            font=(FONT_MONO, 11),
            border_color=C["border"], border_width=1,
            corner_radius=8, wrap="word")
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")

    # ── Helpers ───────────────────────────────────────────────────────────

    def _choose_out(self):
        folder = filedialog.askdirectory()
        if folder:
            name = os.path.basename(
                self.out_var.get()) or "transfer_output.xlsx"
            self.out_var.set(os.path.join(folder, name))

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _set_status(self, text: str, color: str = C["muted"]):
        self.status_lbl.configure(text=f"● {text}", text_color=color)

    # ── Run ───────────────────────────────────────────────────────────────

    def _run_threaded(self):
        self.run_btn.configure(state="disabled", text="  Running…")
        self.progress.set(0)
        self._set_status("Running…", C["amber"])
        threading.Thread(target=self._process, daemon=True).start()

    def _finish(self, ok: bool):
        self.run_btn.configure(state="normal",
                               text="  Run Transfer Calculation  →")
        self.progress.set(1 if ok else 0)
        self._set_status("Done ✓" if ok else "Error ✗",
                         C["green"] if ok else C["red"])

    def _process(self):
        """Validate inputs, call logic.run_transfer(), update UI with results."""
        try:
            # ── Validate ────────────────────────────────────────────────
            to_name = self.to_name_var.get().strip()
            if not to_name:
                messagebox.showerror(
                    "Missing", "Enter a TO location name.")
                self.after(0, self._finish, False)
                return

            for label, row in [("Sales (Source)", self.row_source),
                               ("From Stock (Samakhosi)", self.row_d1),
                               ("To Stock (Destination)", self.row_d2)]:
                if not row.is_set():
                    messagebox.showerror(
                        "Missing", f"{label} file not selected.")
                    self.after(0, self._finish, False)
                    return

            # ── Log inputs ──────────────────────────────────────────────
            self.after(0, self._log, "━" * 46)
            self.after(0, self._log, f"  From Location : Samakhosi (fixed)")
            self.after(0, self._log, f"  To Location   : {to_name}")
            self.after(
                0, self._log, f"  Source File   : {os.path.basename(self.row_source.get())}")
            self.after(0, self._log, "  Cleaning sales file automatically...")
            self.after(
                0, self._log, f"  From Stock    : {os.path.basename(self.row_d1.get())}")
            self.after(
                0, self._log, f"  To Stock      : {os.path.basename(self.row_d2.get())}")
            self.after(0, self._log, "  Running calculation…")
            self.after(0, self.progress.set, 0.2)

            # ── Call logic with original parameters ────────────────────────
            stats = run_transfer(
                source_path=self.row_source.get(),
                d1_path=self.row_d1.get(),
                d2_path=self.row_d2.get(),
                d2_name=to_name,  # This maps to the d2_name parameter in your logic
                out_path=self.out_var.get().strip() or "transfer_output.xlsx",
            )

            preview_output(stats['out_path'])

            # ── Update UI with returned stats ───────────────────────────
            self.after(0, self.progress.set, 0.9)
            self.after(0, self.stat_total.set,
                       stats['total'],        C["accent"])
            self.after(0, self.stat_xfer.set,
                       stats['to_transfer'],  C["green"])
            self.after(0, self.stat_skip.set,
                       stats['skipped'],      C["muted"])

            self.after(0, self._log, f"  Total products   : {stats['total']}")
            self.after(
                0, self._log, f"  Items to transfer: {stats['to_transfer']}")
            self.after(
                0, self._log, f"  Skipped          : {stats['skipped']}")
            self.after(0, self._log, f"✓  Saved → {stats['out_path']}")
            self.after(0, self._log, "━" * 46)

            self.after(0, self.progress.set, 1.0)
            self.after(0, self._finish, True)
            self.after(0, messagebox.showinfo, "Done ✓",
                       f"Transfer sheet saved!\n\n{stats['out_path']}")

        except Exception as exc:
            self.after(0, self._log, f"✗  ERROR: {exc}")
            self.after(0, self._finish, False)
            self.after(0, messagebox.showerror, "Error", str(exc))


# ── Entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        import customtkinter  # noqa: F401
    except ImportError:
        import subprocess
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install",
             "customtkinter", "openpyxl", "pandas"])

    App().mainloop()
