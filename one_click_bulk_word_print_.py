"""
One-Click Word Bulk Printer (Windows)
Reads External Id + ASN quantity from an Excel workbook, sums duplicate IDs across PO sections,
checks for matching Word files, shows a preview, then asks before printing to the default printer.

Requirements:
  - Windows
  - Microsoft Word desktop installed
  - Python 3.10+
  - pip install openpyxl pywin32

Run:
  python bulk_word_print.py
"""

import os
import re
import csv
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
from pathlib import Path
from datetime import datetime

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

WORD_EXTENSIONS = (".docx", ".doc", ".docm")


def normalize_header(value):
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_id(value):
    if value is None:
        return ""
    # Excel may store IDs as strings; preserve the exact ID text otherwise.
    return str(value).strip()


def find_header_row(ws):
    """Find the row containing External Id and ASN quantity, even if PO sections repeat."""
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        normalized = [normalize_header(v) for v in row]
        has_external = any(v in ("externalid", "externalidentifier") for v in normalized)
        has_qty = any(v in ("asnquantity", "quantity", "asnqty") for v in normalized)
        if has_external and has_qty:
            return row_idx, normalized
    return None, None


def read_quantities(excel_path, sheet_name=None):
    if load_workbook is None:
        raise RuntimeError("openpyxl is not installed. Run: py -m pip install openpyxl")
    wb = load_workbook(excel_path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    totals = defaultdict(int)
    source_rows = defaultdict(list)
    missing_qty_rows = []
    header_row, headers = find_header_row(ws)
    if header_row is None:
        wb.close()
        raise ValueError(
            "Could not find a header row containing 'External Id' and 'ASN quantity'. "
            "Check that the workbook has these column headings."
        )

    external_col = next(i for i, h in enumerate(headers) if h in ("externalid", "externalidentifier"))
    qty_col = next(i for i, h in enumerate(headers) if h in ("asnquantity", "quantity", "asnqty"))

    for row_idx, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
        external = normalize_id(row[external_col] if external_col < len(row) else None)
        qty_value = row[qty_col] if qty_col < len(row) else None

        if not external:
            continue
        # Skip summary rows such as TOTAL / CUMULATIVE / PO headings.
        if external.upper() in {"TOTAL", "CUMULATIVE", "PO", "ASIN"}:
            continue
        # Ignore rows where the quantity cell is blank or non-numeric.
        try:
            qty_float = float(qty_value)
            qty = int(qty_float)
            if qty_float != qty or qty < 0:
                raise ValueError
        except (TypeError, ValueError):
            missing_qty_rows.append((row_idx, external, qty_value))
            continue
        if qty == 0:
            continue
        totals[external] += qty
        source_rows[external].append(row_idx)

    wb.close()
    return dict(totals), dict(source_rows), missing_qty_rows, ws.title


def index_word_files(folder):
    """Index Word files recursively by filename stem, case-insensitively."""
    index = defaultdict(list)
    for path in Path(folder).rglob("*"):
        if path.is_file() and path.suffix.lower() in WORD_EXTENSIONS:
            index[path.stem.strip().casefold()].append(path)
    return index


def match_files(folder, quantities):
    file_index = index_word_files(folder)
    matches, missing, ambiguous = {}, [], []
    for external_id, qty in quantities.items():
        candidates = file_index.get(external_id.casefold(), [])
        if not candidates:
            missing.append((external_id, qty))
        elif len(candidates) > 1:
            ambiguous.append((external_id, qty, candidates))
        else:
            matches[external_id] = candidates[0]
    return matches, missing, ambiguous


def write_preview_report(output_folder, rows):
    report_path = Path(output_folder) / ("bulk_print_preview_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv")
    with report_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["External ID", "Copies", "File", "Status"])
        writer.writerows(rows)
    return report_path


class BulkPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("One-Click Word Bulk Printer")
        self.root.geometry("980x650")
        self.root.minsize(820, 520)

        self.excel_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.sheet_var = tk.StringVar(value="Active sheet")
        self.status_var = tk.StringVar(value="Choose your Excel workbook and the folder containing Word files.")
        self.rows = []
        self.matches = {}
        self.quantities = {}
        self.missing = []
        self.ambiguous = []
        self.sheet_name = None

        outer = ttk.Frame(root, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Excel workbook (.xlsx)", font=("", 10, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Entry(outer, textvariable=self.excel_var).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(outer, text="Browse…", command=self.pick_excel).grid(row=1, column=1, sticky="ew")

        ttk.Label(outer, text="Folder containing Word files (.docx / .doc / .docm)", font=("", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(outer, textvariable=self.folder_var).grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(outer, text="Browse…", command=self.pick_folder).grid(row=3, column=1, sticky="ew")

        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, columnspan=2, sticky="w", pady=12)
        ttk.Button(actions, text="1. Check files & build preview", command=self.build_preview).pack(side="left")
        self.print_button = ttk.Button(actions, text="2. Confirm & print", command=self.confirm_print, state="disabled")
        self.print_button.pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Open report folder", command=self.open_report_folder).pack(side="left", padx=(8, 0))

        ttk.Label(outer, text="Preview — each row is one file; duplicate External IDs are summed across all PO sections",
                  font=("", 10, "bold")).grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 6))

        table_frame = ttk.Frame(outer)
        table_frame.grid(row=6, column=0, columnspan=2, sticky="nsew")
        columns = ("external", "copies", "file", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)
        self.tree.heading("external", text="External ID")
        self.tree.heading("copies", text="Copies")
        self.tree.heading("file", text="Matched Word file")
        self.tree.heading("status", text="Status")
        self.tree.column("external", width=150, anchor="w")
        self.tree.column("copies", width=70, anchor="center")
        self.tree.column("file", width=480, anchor="w")
        self.tree.column("status", width=130, anchor="w")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        ttk.Label(outer, textvariable=self.status_var, wraplength=900).grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 0))
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(6, weight=1)

    def pick_excel(self):
        path = filedialog.askopenfilename(title="Select Excel workbook", filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.excel_var.set(path)

    def pick_folder(self):
        path = filedialog.askdirectory(title="Select folder containing Word files")
        if path:
            self.folder_var.set(path)

    def build_preview(self):
        excel = self.excel_var.get().strip()
        folder = self.folder_var.get().strip()
        if not excel or not Path(excel).is_file():
            messagebox.showerror("Excel file required", "Please choose a valid .xlsx workbook.")
            return
        if not folder or not Path(folder).is_dir():
            messagebox.showerror("Folder required", "Please choose a valid folder containing Word files.")
            return
        try:
            quantities, source_rows, bad_qty, sheet_name = read_quantities(excel)
            if not quantities:
                messagebox.showwarning("No quantities found", "No rows with a valid External Id and ASN quantity were found.")
                return
            matches, missing, ambiguous = match_files(folder, quantities)
        except Exception as e:
            messagebox.showerror("Could not build preview", str(e))
            return

        self.quantities = quantities
        self.matches = matches
        self.missing = missing
        self.ambiguous = ambiguous
        self.sheet_name = sheet_name
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.rows = []
        for external, qty in quantities.items():
            if external in matches:
                status = "Ready"
                filename = str(matches[external])
            elif any(x[0] == external for x in ambiguous):
                status = "AMBIGUOUS — skipped"
                filename = "; ".join(str(p) for p in next(x[2] for x in ambiguous if x[0] == external))
            else:
                status = "MISSING — skipped"
                filename = ""
            self.tree.insert("", "end", values=(external, qty, filename, status))
            self.rows.append([external, qty, filename, status])

        report_path = write_preview_report(Path(excel).parent, self.rows)
        total_copies = sum(quantities.values())
        ready_copies = sum(quantities[k] for k in matches)
        missing_count = len(missing)
        ambiguous_count = len(ambiguous)
        self.status_var.set(
            f"Sheet: {sheet_name} | Unique IDs: {len(quantities)} | Requested copies: {total_copies} | "
            f"Ready copies: {ready_copies} | Missing IDs: {missing_count} | Duplicate filename matches: {ambiguous_count} | "
            f"Preview report: {report_path}"
        )
        self.print_button.configure(state="normal" if matches else "disabled")
        if missing or ambiguous:
            details = []
            if missing:
                details.append("Missing files: " + ", ".join(f"{x[0]} ({x[1]} copies)" for x in missing))
            if ambiguous:
                details.append("Multiple files with the same name: " + ", ".join(x[0] for x in ambiguous))
            messagebox.showwarning("Preview built — some files need attention",
                                   "\n\n".join(details) + "\n\nThese rows will be skipped. Review the preview before printing.")
        else:
            messagebox.showinfo("Preview ready", f"All {len(quantities)} IDs matched a Word file.\nRequested copies: {total_copies}\n\nReview the list, then click 'Confirm & print'.")

    def open_report_folder(self):
        path = self.excel_var.get().strip()
        folder = str(Path(path).parent) if path else str(Path.home())
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Could not open folder", str(e))

    def confirm_print(self):
        if not self.matches:
            messagebox.showwarning("Nothing to print", "Build a preview first and make sure at least one file matched.")
            return
        total = sum(self.quantities[k] for k in self.matches)
        msg = (
            f"Send {len(self.matches)} Word files to the DEFAULT printer?\n\n"
            f"Copies requested for matched files: {total}\n"
            f"Missing IDs (skipped): {len(self.missing)}\n"
            f"Ambiguous IDs (skipped): {len(self.ambiguous)}\n\n"
            "Please make sure the correct printer is selected as your Windows default and that paper is loaded.\n"
            "Continue?"
        )
        if not messagebox.askyesno("Confirm bulk printing", msg):
            return
        self.print_button.configure(state="disabled")
        self.status_var.set("Printing started. Microsoft Word may open in the background; please wait…")
        threading.Thread(target=self.run_print, daemon=True).start()

    def run_print(self):
        try:
            import win32com.client
        except ImportError:
            self.root.after(0, lambda: messagebox.showerror(
                "Missing package", "Install pywin32 with this command:\npy -m pip install pywin32"))
            self.root.after(0, lambda: self.print_button.configure(state="normal"))
            return

        word = None
        printed_docs = 0
        printed_copies = 0
        errors = []
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for external, path in self.matches.items():
                copies = self.quantities[external]
                try:
                    doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False, Visible=False)
                    try:
                        # Background=False asks Word to finish submitting the job before continuing.
                        for copy_number in range(copies):
                            doc.PrintOut(Background=False)
                            printed_copies += 1
                        printed_docs += 1
                    finally:
                        doc.Close(SaveChanges=False)
                except Exception as e:
                    errors.append(f"{external}: {e}")
        except Exception as e:
            errors.append(f"Could not start Microsoft Word: {e}")
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

        def finish():
            self.print_button.configure(state="normal")
            if errors:
                self.status_var.set(f"Print attempt finished with {len(errors)} error(s). Submitted {printed_copies} copies across {printed_docs} files.")
                messagebox.showwarning("Printing completed with errors",
                                       f"Submitted {printed_copies} copies across {printed_docs} files.\n\nErrors:\n" + "\n".join(errors[:15]))
            else:
                self.status_var.set(f"Submitted {printed_copies} copies across {printed_docs} files to the default printer.")
                messagebox.showinfo("Print jobs submitted", f"Submitted {printed_copies} copies across {printed_docs} files to the default printer.\n\nNote: this confirms submission to Windows/Word, not physical completion by the printer.")
        self.root.after(0, finish)


def main():
    if os.name != "nt":
        print("This tool is intended for Windows with Microsoft Word installed.")
        return
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    BulkPrintApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
