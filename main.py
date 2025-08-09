# AddtoPDF 
# Copyright (c) 2025 Ravinder Singh
# Licensed under the AGPL-3.0-only (see LICENSE) or a commercial license from the author.

import os
import threading
import queue
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Callable, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from PyPDF2 import PdfMerger
except Exception as e:
    raise SystemExit("PyPDF2 is required. Install with: pip install PyPDF2 or use install requirements") from e

COM_AVAILABLE = False
if os.name == "nt":
    try:
        import comtypes.client  
        COM_AVAILABLE = True
    except Exception:
        COM_AVAILABLE = False

APP_NAME = "HardPdfMerger"
APP_AUTHOR = "Ravinder Singh"
APP_TAGLINE = "Convert & merge Office files into a single PDF — fully offline"
SUPPORTED_EXTS = [".doc", ".docx", ".ppt", ".pptx", ".pdf"]


# =========================
# Conversion helpers
# =========================
def has_soffice() -> bool:
    return shutil.which("soffice") is not None

def convert_via_soffice(src: str, out_pdf: str):
    tmpdir = tempfile.mkdtemp(prefix="hardpdf_")
    try:
        cmd = ["soffice", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmpdir, src]
        result = subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        gen = Path(tmpdir) / (Path(src).stem + ".pdf")
        if not gen.exists():
            raise RuntimeError("LibreOffice did not produce a PDF.")
        Path(out_pdf).write_bytes(gen.read_bytes())
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

def convert_word_to_pdf_com(word_path: str, output_path: str):
    word = comtypes.client.CreateObject('Word.Application')  # type: ignore
    word.Visible = False
    try:
        doc = word.Documents.Open(word_path)
        doc.SaveAs(output_path, FileFormat=17)  # wdFormatPDF
        doc.Close()
    finally:
        word.Quit()

def convert_ppt_to_pdf_com(ppt_path: str, output_path: str):
    powerpoint = comtypes.client.CreateObject('Powerpoint.Application')  # type: ignore
    powerpoint.Visible = False
    try:
        ppt = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        ppt.SaveAs(output_path, FileFormat=32)  # ppSaveAsPDF
        ppt.Close()
    finally:
        powerpoint.Quit()

def convert_office_to_pdf(src: Path, out_pdf: Path):
    suffix = src.suffix.lower()
    if suffix in [".doc", ".docx", ".ppt", ".pptx"]:
        # Prefer MS Office (if available), then fallback to LibreOffice
        if COM_AVAILABLE and os.name == "nt":
            try:
                if suffix in [".doc", ".docx"]:
                    convert_word_to_pdf_com(str(src), str(out_pdf))
                    return
                else:
                    convert_ppt_to_pdf_com(str(src), str(out_pdf))
                    return
            except Exception:
                pass  # fallback
        if has_soffice():
            convert_via_soffice(str(src), str(out_pdf))
            return
        raise RuntimeError("No converter available. Install Microsoft Office (Windows) or LibreOffice.")
    raise ValueError("Unsupported extension for conversion.")

def list_supported_files(folder: str) -> List[Path]:
    p = Path(folder)
    return [f for f in p.iterdir() if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS]

def convert_all_to_pdfs(
    files: List[Path],
    progress_callback: Callable[[int, int, str], None],
    status_callback: Callable[[str, str], None],
) -> List[str]:
    total = len(files)
    done = 0
    out_list: List[str] = []

    for f in files:
        done += 1
        ext = f.suffix.lower()
        try:
            if ext == ".pdf":
                status_callback(f.name, "Queued")
                out_list.append(str(f))
                progress_callback(done, total, f"Queued PDF: {f.name}")
            else:
                status_callback(f.name, "Converting")
                pdf_out = str(f.with_suffix(".pdf"))
                progress_callback(done, total, f"Converting {f.name} → {Path(pdf_out).name}")
                convert_office_to_pdf(f, Path(pdf_out))
                status_callback(f.name, "Converted")
                out_list.append(pdf_out)
        except Exception as e:
            status_callback(f.name, "Failed")
            progress_callback(done, total, f"Failed: {f.name} ({e})")

    return out_list

def merge_pdfs(pdf_paths: List[str], output_path: str):
    merger = PdfMerger()
    try:
        for pdf in pdf_paths:
            merger.append(pdf)
        merger.write(output_path)
    finally:
        merger.close()


# =========================
# UI
# =========================
class CleanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.minsize(840, 560)

        # state
        self._processing = False
        self._worker: Optional[threading.Thread] = None
        self._q: "queue.Queue[Tuple]" = queue.Queue()
        self._files: List[Path] = []

        self._init_style()
        self._build_layout()
        self._set_status("Ready.")

    # ---- Style / Theme ----
    def _init_style(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        # Palette
        self["bg"] = "#F5F6FA"
        self.style.configure("App.TFrame", background="#F5F6FA")
        self.style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        self.style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), background="#F5F6FA")
        self.style.configure("Sub.TLabel", font=("Segoe UI", 10), foreground="#6B7280", background="#F5F6FA")
        self.style.configure("CardTitle.TLabel", font=("Segoe UI", 12, "bold"), background="#FFFFFF")
        self.style.configure("Body.TLabel", font=("Segoe UI", 10), background="#FFFFFF")
        self.style.configure("Accent.TButton", font=("Segoe UI Semibold", 11), padding=10)
        self.style.configure("Ghost.TButton", font=("Segoe UI", 10), padding=8)
        self.style.configure("TEntry", padding=6)
        self.style.configure("TProgressbar", thickness=10)

    # ---- Layout ----
    def _build_layout(self):
        root = ttk.Frame(self, style="App.TFrame")
        root.pack(fill="both", expand=True)

        # Header
        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill="x", padx=20, pady=(16, 8))
        ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text=APP_TAGLINE, style="Sub.TLabel").pack(anchor="w")

        # Card
        card = ttk.Frame(root, style="Card.TFrame")
        card.pack(fill="both", expand=True, padx=20, pady=8)
        card.grid_columnconfigure(0, weight=1)

        # Folder picker
        pick_row = ttk.Frame(card, style="Card.TFrame")
        pick_row.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        ttk.Label(pick_row, text="Choose a folder:", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 4))

        entry_row = ttk.Frame(pick_row, style="Card.TFrame")
        entry_row.pack(fill="x")
        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(entry_row, textvariable=self.folder_var)
        self.folder_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(entry_row, text="Browse…", style="Accent.TButton", command=self._pick_folder).pack(side="left", padx=(8, 0))

        ttk.Label(card, text=f"Supported: {', '.join(SUPPORTED_EXTS)}", style="Body.TLabel").grid(row=1, column=0, sticky="w", padx=18)

        # File table
        table_wrap = ttk.Frame(card, style="Card.TFrame")
        table_wrap.grid(row=2, column=0, sticky="nsew", padx=18, pady=(12, 8))
        card.grid_rowconfigure(2, weight=1)

        columns = ("name", "type", "status")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=10)
        self.tree.heading("name", text="File")
        self.tree.heading("type", text="Type")
        self.tree.heading("status", text="Status")
        self.tree.column("name", width=520, anchor="w")
        self.tree.column("type", width=80, anchor="center")
        self.tree.column("status", width=120, anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")

        vsb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        # Action bar
        act = ttk.Frame(card, style="Card.TFrame")
        act.grid(row=3, column=0, sticky="ew", padx=18, pady=(8, 18))
        act.grid_columnconfigure(0, weight=1)

        self.run_btn = ttk.Button(act, text="Convert & Merge", style="Accent.TButton", command=self._start)
        self.run_btn.grid(row=0, column=0, sticky="w")

        self.clear_btn = ttk.Button(act, text="Clear", style="Ghost.TButton", command=self._clear)
        self.clear_btn.grid(row=0, column=1, padx=(8, 0), sticky="w")

        # Progress + status
        prog = ttk.Frame(root, style="App.TFrame")
        prog.pack(fill="x", padx=20, pady=(0, 12))
        self.pbar = ttk.Progressbar(prog, orient="horizontal", mode="determinate", maximum=100)
        self.pbar.pack(fill="x")
        self.status_lbl = ttk.Label(prog, text="", style="Sub.TLabel")
        self.status_lbl.pack(anchor="w", pady=(6, 0))

        # Footer
        footer = ttk.Frame(root, style="App.TFrame")
        footer.pack(fill="x", padx=20, pady=(0, 16))
        ttk.Label(footer, text=f"© 2025 {APP_AUTHOR} • AGPL-3.0-only", style="Sub.TLabel").pack(side="right")

    # ---- Helpers ----
    def _set_status(self, text: str):
        self.status_lbl.config(text=text)

    def _pick_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        self.folder_var.set(folder)
        self._files = list_supported_files(folder)
        self._refresh_table()

        count = len(self._files)
        self._set_status(f"Found {count} supported file(s).")

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for f in self._files:
            self.tree.insert("", "end", iid=f.name, values=(f.name, f.suffix.lower(), "Ready"))

    def _update_row_status(self, filename: str, status: str):
        if self.tree.exists(filename):
            vals = list(self.tree.item(filename, "values"))
            vals[2] = status
            self.tree.item(filename, values=vals)

    def _clear(self):
        self.folder_var.set("")
        self._files = []
        self.tree.delete(*self.tree.get_children())
        self.pbar["value"] = 0
        self._set_status("Ready.")

    # ---- Processing ----
    def _start(self):
        if self._processing:
            return
        folder = self.folder_var.get().strip()
        if not folder or not Path(folder).exists():
            messagebox.showerror("Invalid folder", "Please select a valid folder.")
            return
        self._files = list_supported_files(folder)
        if not self._files:
            messagebox.showwarning("No files", "No supported files found in the selected folder.")
            return

        out = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save merged PDF as"
        )
        if not out:
            return

        self._processing = True
        self.run_btn.state(["disabled"])
        self.clear_btn.state(["disabled"])
        self._set_status("Processing… Converting files, then merging to a single PDF.")

        # kick worker
        self._worker = threading.Thread(target=self._worker_main, args=(self._files, out), daemon=True)
        self._worker.start()
        self.after(120, self._poll_worker)

    def _progress_cb(self, cur: int, total: int, msg: str):
        self._q.put(("progress", cur, total, msg))

    def _status_cb(self, filename: str, status: str):
        self._q.put(("row", filename, status))

    def _worker_main(self, files: List[Path], out_pdf: str):
        try:
            pdfs = convert_all_to_pdfs(files, self._progress_cb, self._status_cb)
            if not pdfs:
                raise RuntimeError("No files to merge.")
            self._q.put(("stage", f"Merging {len(pdfs)} PDF(s)…"))
            merge_pdfs(pdfs, out_pdf)
            self._q.put(("done", out_pdf))
        except Exception as e:
            self._q.put(("error", str(e)))

    def _poll_worker(self):
        try:
            while True:
                item = self._q.get_nowait()
                kind = item[0]
                if kind == "progress":
                    _, cur, total, msg = item
                    pct = int((cur / max(total, 1)) * 100)
                    self.pbar["value"] = pct
                    self._set_status(msg)
                elif kind == "row":
                    _, filename, status = item
                    self._update_row_status(filename, status)
                elif kind == "stage":
                    _, msg = item
                    self._set_status(msg)
                elif kind == "done":
                    _, path = item
                    self._finish(True, path=path)
                elif kind == "error":
                    _, err = item
                    self._finish(False, error=err)
        except queue.Empty:
            pass

        if self._worker and self._worker.is_alive():
            self.after(120, self._poll_worker)

    def _finish(self, ok: bool, path: Optional[str] = None, error: Optional[str] = None):
        self._processing = False
        self.run_btn.state(["!disabled"])
        self.clear_btn.state(["!disabled"])

        if ok:
            self.pbar["value"] = 100
            self._set_status("Completed.")
            messagebox.showinfo("Success", f"PDF saved to:\n{path}")
            # mark all rows as "Merged"
            for f in self._files:
                self._update_row_status(f.name, "Merged")
        else:
            self._set_status("Failed.")
            messagebox.showerror("Error", error or "Unknown error")


def main():
    app = CleanApp()
    app.mainloop()

if __name__ == "__main__":
    main()
