# gui/app.py
"""
OFAC Scanner – Main GUI (revamp)
Tabbed interface: Scan, Extract Headers, Settings, History.
Job queue, optional mapping popup, UI polish.
Responsive layout with minimum pane sizes for small screens.
Flexible sidebar with separate company and password sections.
Password buttons fixed on the left, never overflow.
Now passes configured output folder to queued jobs.
Auto‑refresh file list every 5 seconds (optional).
Preview pane now supports switching between Excel sheets.
Drag‑and‑drop files into the file list to add them to the watch folder.
Confirmation before closing while a job is running.
Open Output Folder button for quick access to results.
Safe defaults – empty paths prevent crashes on inaccessible drives.
"""

import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import csv
import json
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import threading
from collections import defaultdict

import polars as pl
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinterdnd2 import TkinterDnD

# ---------- Backend imports ----------
from utils.file_handler import (
    get_all_files,
    is_file_extension,
    FILE_EXTENSIONS_EXCEL,
    FILE_EXTENSIONS_TEXT,
    FILE_EXTENSIONS_ARCHIVE,
)
from utils.company_headers import get_company_header, set_header_path
from engine.inspector import inspect_file
from engine.classifier import classify_columns
from engine.job import ExtractionJob
from engine.queue import JobQueue, QueuedJob

# ---------- Settings ----------
SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "ofac_settings.json"
)

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {}

def save_settings(settings_dict):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings_dict, f, indent=4)


# ==================== MAPPING POPUP ====================
class MappingPopup:
    """Popup window to review and edit column mappings."""
    def __init__(self, parent, watch_folder, company, passwords, output_folder, mapping_data):
        self.parent = parent
        self.watch_folder = watch_folder
        self.company = company
        self.passwords = passwords
        self.output_folder = output_folder
        self.mapping_data = mapping_data

        self.window = tb.Toplevel(parent.root)
        self.window.title("Column Mapping – " + company)
        self.window.geometry("950x600")
        self.window.attributes('-topmost', True)
        self.window.minsize(700, 400)

        self.build_ui()
        self.window.after(100, self.detect_headers)

    def build_ui(self):
        info_frame = tb.Frame(self.window, padding=10)
        info_frame.pack(fill=X)
        tb.Label(info_frame, text=f"Company: {self.company}", font=("Helvetica", 11, "bold")).pack(anchor=W)

        toolbar = tb.Frame(self.window)
        toolbar.pack(fill=X, padx=10, pady=(0,5))
        tb.Button(toolbar, text="Clear All Mappings", bootstyle="warning", command=self.clear_mappings).pack(side=LEFT, padx=5)

        tree_frame = tb.Frame(self.window)
        tree_frame.pack(fill=BOTH, expand=YES, padx=10, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree = tb.Treeview(tree_frame, bootstyle="primary",
                                columns=("file", "sheet", "table", "header", "category", "col_index"),
                                show="headings")
        self.tree.heading("file", text="File")
        self.tree.heading("sheet", text="Sheet")
        self.tree.heading("table", text="Table")
        self.tree.heading("header", text="Header")
        self.tree.heading("category", text="Category")
        self.tree.heading("col_index", text="ColIndex")
        self.tree.column("file", width=120, stretch=False)
        self.tree.column("sheet", width=100, stretch=False)
        self.tree.column("table", width=80, stretch=False)
        self.tree.column("header", width=150, stretch=False)
        self.tree.column("category", width=150, stretch=False)
        self.tree.column("col_index", width=0, stretch=False)

        scroll_y = tb.Scrollbar(tree_frame, orient=VERTICAL, command=self.tree.yview)
        scroll_x = tb.Scrollbar(tree_frame, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        self.combo_edit = tb.Combobox(
            tree_frame,
            values=["surname", "firstname", "fullname", "sex", "dob", "policynum", "ignore"],
            state="readonly",
            bootstyle="primary",
        )
        self.combo_edit.place_forget()
        self.tree.bind("<Double-1>", self.on_double_click)
        self.combo_edit.bind("<<ComboboxSelected>>", self.on_combo_selected)
        self.combo_edit.bind("<FocusOut>", lambda e: self.combo_edit.place_forget())

        btn_frame = tb.Frame(self.window, padding=10)
        btn_frame.pack(fill=X)
        tb.Button(btn_frame, text="Save & Close", bootstyle="success", command=self.save).pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="Cancel", bootstyle="secondary", command=self.cancel).pack(side=RIGHT, padx=5)

    def clear_mappings(self):
        if not self.tree.get_children():
            return
        if messagebox.askyesno("Confirm", "Remove all detected mappings for these files?", parent=self.window):
            self.tree.delete(*self.tree.get_children())
            selected = [f for f, var in self.parent.file_vars.items() if var.get() == 1]
            for fname in selected:
                filepath = os.path.join(self.watch_folder, fname)
                keys = [k for k in self.mapping_data if k.startswith(filepath + "|")]
                for k in keys:
                    del self.mapping_data[k]

    def detect_headers(self):
        selected = [f for f, var in self.parent.file_vars.items() if var.get() == 1]
        if not selected:
            messagebox.showwarning("Warning", "No files selected.", parent=self.window)
            self.window.destroy()
            return

        try:
            header_dict, _ = get_company_header(self.company)
        except Exception as e:
            messagebox.showerror("Error", f"Header dictionary error: {e}", parent=self.window)
            self.window.destroy()
            return

        for fname in selected:
            filepath = os.path.join(self.watch_folder, fname)
            keys_to_remove = [k for k in self.mapping_data if k.startswith(filepath + "|")]
            for k in keys_to_remove:
                del self.mapping_data[k]

        for fname in selected:
            filepath = os.path.join(self.watch_folder, fname)
            try:
                structure = inspect_file(filepath, self.passwords, self.output_folder)
            except Exception as e:
                messagebox.showwarning("Error", f"Inspector failed on {fname}: {e}", parent=self.window)
                continue
            if not structure.tables:
                continue

            for table in structure.tables:
                mapping = classify_columns(table, header_dict)
                table_key = f"{filepath}|{table.sheet_name}|{table.col_start}|{table.col_end}"
                self.mapping_data[table_key] = mapping.mappings

                for col_map in mapping.mappings:
                    self.tree.insert("", "end", values=(
                        fname,
                        table.sheet_name,
                        f"Table {table.col_start}-{table.col_end}",
                        col_map.header_raw,
                        col_map.category,
                        str(col_map.column_index),
                    ))

    def on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        col = self.tree.identify_column(event.x)
        if col != "#5":
            return
        vals = self.tree.item(item, "values")
        if not vals or len(vals) < 6:
            return
        bbox = self.tree.bbox(item, col)
        if bbox:
            x, y, w, h = bbox
            self.combo_edit.place(x=x, y=y+h, width=w, anchor=SW)
            self.combo_edit.set(vals[4])
            self.combo_edit.focus_set()
            self._edit_item = item

    def on_combo_selected(self, event):
        if not hasattr(self, '_edit_item'):
            return
        new_cat = self.combo_edit.get()
        vals = list(self.tree.item(self._edit_item, "values"))
        fname, sheet, tbl_str, header, _, col_index_str = vals
        vals[4] = new_cat
        self.tree.item(self._edit_item, values=vals)

        filepath = os.path.join(self.watch_folder, fname)
        col_start, col_end = map(int, tbl_str.replace("Table ", "").split("-"))
        table_key = f"{filepath}|{sheet}|{col_start}|{col_end}"
        col_index = int(col_index_str)

        if table_key in self.mapping_data:
            for cm in self.mapping_data[table_key]:
                if cm.column_index == col_index:
                    cm.category = new_cat
                    break

        self.combo_edit.place_forget()

    def save(self):
        self.window.destroy()

    def cancel(self):
        self.window.destroy()


# ==================== MAIN APP ====================
class OFACScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OFAC Scanner")
        self.root.geometry("1400x850")
        self.root.minsize(1100, 700)

        # Intercept window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.settings = load_settings()
        theme = self.settings.get("theme", "flatly")
        try:
            self.root.style.theme_use(theme)
        except:
            pass

        # Paths – empty by default to avoid startup crashes
        self.watch_folder = self.settings.get("folder", "")
        self.password_csv_path = self.settings.get("csv", "")
        self.output_folder = self.settings.get("output_folder", "")
        self.header_folder = self.settings.get("header_folder", "")
        set_header_path(self.header_folder)

        self.mapping_data = {}
        self.company_data = []

        self.queue = JobQueue(
            on_log=self.log,
            on_progress=self._update_progress,
            on_job_start=self._on_queue_job_start,
            on_job_finish=self._on_queue_job_finish,
            on_queue_empty=self._on_queue_empty,
        )

        # Create status variable early (needed for password list hints)
        self.status_var = tk.StringVar(value="Ready")

        # -------- Tabs --------
        self.notebook = tb.Notebook(self.root, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=YES, padx=10, pady=(10, 0))

        self.scan_tab = tb.Frame(self.notebook)
        self.extract_tab = tb.Frame(self.notebook)
        self.settings_tab = tb.Frame(self.notebook)
        self.history_tab = tb.Frame(self.notebook)

        self.notebook.add(self.scan_tab, text="Scan")
        self.notebook.add(self.extract_tab, text="Extract Headers")
        self.notebook.add(self.settings_tab, text="Settings")
        self.notebook.add(self.history_tab, text="History")

        self.build_scan_tab()
        self.build_extract_tab()
        self.build_settings_tab()
        self.build_history_tab()

        if self.password_csv_path and os.path.exists(self.password_csv_path):
            self.load_company_data()

        # -------- Bottom bar --------
        bottom_frame = tb.Frame(self.root)
        bottom_frame.pack(side=BOTTOM, fill=X, padx=10, pady=5)

        self.status_label = tb.Label(bottom_frame, textvariable=self.status_var, bootstyle="info", anchor=W)
        self.status_label.pack(side=LEFT, fill=X, expand=YES)

        self.progress = tb.Progressbar(bottom_frame, bootstyle="success", mode="determinate", maximum=100, value=0)
        self.progress.pack(side=RIGHT, padx=(20, 0))
        self.progress.pack_forget()

        self.refresh_file_list()
        self.refresh_password_list()
        self.ext_refresh_file_list()
        self.ext_refresh_password_list()

        # Start the auto‑refresh loop
        self._auto_refresh_loop()

    # ==================== CLOSE HANDLER ====================
    def on_close(self):
        """Ask for confirmation if a job is running."""
        if self.queue._running:
            if not messagebox.askyesno("Quit", "A scan is still running.\nAre you sure you want to exit?"):
                return
            self.queue.stop()
        self.root.destroy()

    # ==================== DRAG‑AND‑DROP ====================
    def _setup_drop_target(self, widget):
        """Enable drag‑and‑drop of files onto a widget."""
        def on_drop(event):
            files = self.root.tk.splitlist(event.data)
            added = 0
            for file_path in files:
                if os.path.isfile(file_path):
                    dest = os.path.join(self.watch_folder, os.path.basename(file_path))
                    try:
                        shutil.copy2(file_path, dest)
                        added += 1
                    except Exception as e:
                        self.log(f"Failed to copy {file_path}: {e}")
            if added:
                self.refresh_file_list()
                self.ext_refresh_file_list()
        widget.drop_target_register('*')
        widget.dnd_bind('<<Drop>>', on_drop)

    # ==================== SEARCHABLE COMBOBOX ====================
    def create_searchable_combobox(self, parent, label_text, values, variable):
        """Searchable combobox using a pop‑up listbox (only when focused)."""
        frame = tb.Frame(parent)
        tb.Label(frame, text=label_text).pack(side=tk.LEFT, padx=5)

        entry_var = tk.StringVar()
        entry = tb.Entry(frame, textvariable=entry_var, bootstyle="info")
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        frame.values = list(values)
        frame.entry = entry
        frame.entry_var = entry_var

        popup = None
        listbox = None

        def _create_popup():
            nonlocal popup, listbox
            if popup is not None:
                return
            popup = tk.Toplevel(self.root)
            popup.overrideredirect(True)
            popup.attributes('-topmost', True)
            popup_frame = tb.Frame(popup, bootstyle="light")
            popup_frame.pack(fill=tk.BOTH, expand=True)
            listbox = tk.Listbox(popup_frame, height=6, exportselection=False)
            scrollbar = tb.Scrollbar(popup_frame, orient="vertical", command=listbox.yview)
            listbox.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            def on_listbox_select(event):
                if listbox.curselection():
                    sel = listbox.get(listbox.curselection()[0])
                    entry_var.set(sel)
                    variable.set(sel)
                    _close_popup()
                    self.on_company_selected()

            listbox.bind("<<ListboxSelect>>", on_listbox_select)
            popup.bind("<FocusOut>", lambda e: _close_popup())
            listbox.bind("<FocusOut>", lambda e: _close_popup())

        def _close_popup():
            nonlocal popup, listbox
            if popup:
                popup.destroy()
                popup = None
                listbox = None

        def _update_listbox(*args):
            nonlocal popup, listbox
            search_term = entry_var.get().strip().lower()
            filtered = [v for v in frame.values if search_term in v.lower()] if search_term else frame.values
            if not filtered:
                _close_popup()
                return
            if popup is None:
                _create_popup()
            if popup is None:
                return
            listbox.delete(0, tk.END)
            for item in filtered:
                listbox.insert(tk.END, item)
            listbox.selection_set(0)
            listbox.see(0)
            # Show the popup ONLY if the entry currently has focus
            if entry.focus_get() == entry:
                x = entry.winfo_rootx()
                y = entry.winfo_rooty() + entry.winfo_height()
                popup.geometry(f"{entry.winfo_width()}x{max(6, len(filtered))*20+4}+{x}+{y}")
                popup.deiconify()
            else:
                _close_popup()

        def _on_entry_focus_out(event):
            self.root.after(150, lambda: _close_popup() if not (listbox and listbox.focus_get() == listbox) else None)

        def _on_key_release(event):
            _update_listbox()
            if event.keysym == 'Return' and listbox and listbox.size() == 1:
                sel = listbox.get(0)
                entry_var.set(sel)
                variable.set(sel)
                _close_popup()
                self.on_company_selected()

        entry_var.trace("w", lambda *a: _update_listbox())
        entry.bind("<KeyRelease>", _on_key_release)
        entry.bind("<FocusOut>", _on_entry_focus_out)

        return frame

    # ==================== DATA LOADING ====================
    def load_company_data(self):
        self.company_data = []
        if not self.password_csv_path or not os.path.exists(self.password_csv_path):
            return
        try:
            with open(self.password_csv_path, 'r', newline='', encoding='utf-8-sig') as f:
                sample = f.read(8192)
                f.seek(0)
                has_header = csv.Sniffer().has_header(sample)
                if has_header:
                    reader = csv.DictReader(f)
                    if reader.fieldnames:
                        code_key = next((k for k in reader.fieldnames if k.lower() == 'code'), None)
                        pwd_key = next((k for k in reader.fieldnames if k.lower() == 'password'), None)
                        if code_key and pwd_key:
                            for row in reader:
                                self.company_data.append({'Code': row[code_key].strip(), 'Password': row[pwd_key].strip()})
                        else:
                            f.seek(0)
                            next(reader, None)
                            for row in reader:
                                if len(row) >= 2:
                                    self.company_data.append({'Code': row[0].strip(), 'Password': row[1].strip()})
                else:
                    reader = csv.reader(f)
                    for row in reader:
                        if len(row) >= 2:
                            self.company_data.append({'Code': row[0].strip(), 'Password': row[1].strip()})
        except Exception as e:
            self.status_var.set(f"Cannot load password CSV: {e}")
            return

        codes = sorted({d['Code'] for d in self.company_data})
        self.company_combo.values = codes
        self.company_combo.entry_var.set("")
        self.company_combo.entry.event_generate('<KeyRelease>')
        if codes:
            self.company_combo.entry_var.set(codes[0])
            self.company_var.set(codes[0])
            self.on_company_selected()

        self.ext_populate_company_combo()

    def on_company_selected(self, event=None):
        self.refresh_password_list()
        self.update_queue_button_state()

    # ==================== SCAN TAB: PASSWORDS ====================
    def refresh_password_list(self):
        selected = self.company_var.get()
        all_pwds = [d['Password'] for d in self.company_data if d['Code'] == selected]
        search = self.pass_search_var.get().strip().lower()
        filtered = [p for p in all_pwds if search in p.lower()] if search else all_pwds

        for widget in self.pwd_inner.winfo_children():
            widget.destroy()
        self.pwd_vars.clear()
        if not filtered:
            tb.Label(self.pwd_inner, text="No passwords found", foreground="gray").pack(pady=10)
            self.status_var.set("Select at least one password to enable scanning")
        else:
            for pwd in filtered:
                var = tk.IntVar(value=0)
                self.pwd_vars[pwd] = var
                row = tb.Frame(self.pwd_inner)
                row.pack(fill=X, pady=2)
                tb.Checkbutton(row, variable=var, bootstyle="primary").pack(side=LEFT, padx=5)
                tb.Label(row, text=pwd, anchor="w").pack(side=LEFT, fill=X, expand=True)
            self.status_var.set("Ready")
        self.update_queue_button_state()

    def select_all_passwords(self):
        for var in self.pwd_vars.values():
            var.set(1)

    def clear_all_passwords(self):
        for var in self.pwd_vars.values():
            var.set(0)

    def add_new_password(self):
        code = self.company_var.get()
        if not code:
            messagebox.showwarning("Warning", "Select a company first.")
            return
        dialog = tb.Toplevel(self.root)
        dialog.title("Add Password")
        dialog.geometry("300x150")
        dialog.attributes('-topmost', True)
        tb.Label(dialog, text=f"New password for {code}:").pack(pady=5)
        entry = tb.Entry(dialog)
        entry.pack(pady=5)
        def save():
            pwd = entry.get().strip()
            if pwd:
                self.company_data.append({'Code': code, 'Password': pwd})
                try:
                    with open(self.password_csv_path, 'a', newline='', encoding='utf-8') as f:
                        writer = csv.DictWriter(f, fieldnames=['Code', 'Password'])
                        writer.writerow({'Code': code, 'Password': pwd})
                except Exception as e:
                    messagebox.showerror("Error", f"Could not write to password file: {e}")
                self.refresh_password_list()
                dialog.destroy()
        tb.Button(dialog, text="Save", command=save, bootstyle="success").pack(pady=10)
        dialog.transient(self.root)
        dialog.grab_set()
        self.root.wait_window(dialog)

    # ==================== SCAN TAB: FILE LIST & PREVIEW ====================
    def refresh_file_list(self):
        if not self.watch_folder or not os.path.isdir(self.watch_folder):
            return
        for widget in self.file_inner.winfo_children():
            widget.destroy()
        self.file_vars.clear()

        try:
            items = os.listdir(self.watch_folder)
        except Exception as e:
            self.log(f"Cannot read watch folder: {e}")
            return

        files = []
        for f in items:
            full = os.path.join(self.watch_folder, f)
            if os.path.isfile(full) and not f.endswith('.json'):
                ext = f.split('.')[-1].lower()
                if (ext in FILE_EXTENSIONS_EXCEL and not self.filter_excel.get()) or \
                   (ext in FILE_EXTENSIONS_TEXT and not self.filter_csv.get()) or \
                   (ext in FILE_EXTENSIONS_ARCHIVE and not self.filter_archive.get()):
                    continue
                size = os.path.getsize(full) // 1024
                mod_time = datetime.fromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d %H:%M")
                ftype = "Excel" if ext in FILE_EXTENSIONS_EXCEL else "Text" if ext in FILE_EXTENSIONS_TEXT else "Archive"
                files.append((f, size, mod_time, ftype))

        if not files:
            tb.Label(self.file_inner, text="No files found in watch folder.",
                     font=("Helvetica", 11), foreground="gray").pack(pady=20)
        else:
            files.sort(key=lambda x: x[0])
            for fname, size, mod_time, ftype in files:
                var = tk.IntVar(value=0)
                self.file_vars[fname] = var
                row = tb.Frame(self.file_inner)
                row.pack(fill=X, pady=2)
                cb = tb.Checkbutton(row, variable=var, bootstyle="primary")
                cb.pack(side=LEFT, padx=5)
                lbl = tb.Label(row, text=fname, font=("Helvetica", 10, "bold"), anchor="w")
                lbl.pack(side=LEFT, fill=X, expand=True, padx=5)
                lbl.bind("<Button-1>", lambda e, f=fname: self.show_preview(f))
                meta = tb.Label(row, text=f"{size} KB | {mod_time} | {ftype}", font=("Helvetica", 9), foreground="gray")
                meta.pack(side=RIGHT, padx=5)

        self.update_queue_button_state()

    def select_all_files(self):
        for var in self.file_vars.values():
            var.set(1)
        self.update_queue_button_state()

    def show_preview(self, filename, sheet_name=None):
        """Show preview of a file, optionally specifying the sheet for Excel files."""
        filepath = os.path.join(self.watch_folder, filename)
        self.preview_sheet_combo.pack_forget()
        self.preview_sheet_combo['values'] = []
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = ()
        if not os.path.exists(filepath):
            return
        try:
            ext = filename.split('.')[-1].lower()
            if ext in FILE_EXTENSIONS_EXCEL:
                all_sheets = pl.read_excel(filepath, sheet_id=0).keys()
                sheet_names = list(all_sheets)
                if not sheet_names:
                    return
                if len(sheet_names) > 1:
                    self.preview_sheet_combo['values'] = sheet_names
                    current_sheet = sheet_name if sheet_name in sheet_names else sheet_names[0]
                    self.preview_sheet_var.set(current_sheet)
                    self.preview_sheet_combo.pack(fill=X, padx=5, pady=(5,0), before=self.preview_tree)
                    self.preview_sheet_combo.bind('<<ComboboxSelected>>',
                        lambda e: self.show_preview(filename, self.preview_sheet_var.get()))
                else:
                    current_sheet = sheet_names[0]
                df = pl.read_excel(filepath, sheet_name=current_sheet, has_header=False)
                sample = df.head(20)
            elif ext in FILE_EXTENSIONS_TEXT:
                df = pl.read_csv(filepath, has_header=False, truncate_ragged_lines=True)
                sample = df.head(20)
            else:
                return

            col_names = [f"Col {i}" for i in range(sample.width)]
            self.preview_tree["columns"] = col_names
            self.preview_tree.heading("#0", text="Row")
            self.preview_tree.column("#0", width=50, stretch=False)
            for col in col_names:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100, stretch=False)
            for row_idx in range(sample.height):
                values = [str(sample[row_idx, i]) for i in range(sample.width)]
                self.preview_tree.insert("", "end", text=str(row_idx+1), values=values)
        except Exception as e:
            self.log(f"Preview error: {e}")

    # ==================== MAPPING POPUP OPENER ====================
    def open_mapping_popup(self):
        company = self.company_var.get()
        passwords = self.get_selected_passwords()
        if not company or not passwords:
            messagebox.showerror("Error", "Select a company and at least one password.")
            return
        MappingPopup(self, self.watch_folder, company, passwords, self.output_folder, self.mapping_data)

    def get_selected_passwords(self):
        return [pwd for pwd, var in self.pwd_vars.items() if var.get() == 1]

    # ==================== QUEUE BUTTON STATE ====================
    def update_queue_button_state(self):
        has_files = any(var.get() == 1 for var in self.file_vars.values())
        has_company = bool(self.company_var.get())
        has_passwords = bool(self.get_selected_passwords())
        if has_files and has_company and has_passwords:
            self.add_to_queue_btn.config(state=NORMAL)
        else:
            self.add_to_queue_btn.config(state=DISABLED)

    # ==================== QUEUE OPERATIONS ====================
    def add_to_queue(self):
        selected_files = [f for f, var in self.file_vars.items() if var.get() == 1]
        if not selected_files:
            messagebox.showwarning("Warning", "No files selected.")
            return
        company = self.company_var.get()
        passwords = self.get_selected_passwords()
        email_date = self.get_email_date()

        mapping_overrides = {}
        for table_key, mappings in self.mapping_data.items():
            filepath, sheet, col_start_str, col_end_str = table_key.rsplit("|", 3)
            fname = os.path.basename(filepath)
            if fname in selected_files:
                mapping_overrides.setdefault(filepath, {})[(sheet, int(col_start_str), int(col_end_str))] = mappings

        job = QueuedJob(
            company_code=company,
            passwords=passwords,
            email_received_date=email_date,
            file_names=selected_files,
            input_folder=self.watch_folder,
            mapping_overrides=mapping_overrides,
            output_folder=self.output_folder,
        )
        self.queue.add_job(job)
        self._refresh_queue_display()
        self.log(f"Queued {company} – {len(selected_files)} files")

    def run_queue(self):
        self.queue.start()
        self.run_queue_btn.config(state=DISABLED)
        self.stop_queue_btn.config(state=NORMAL)
        self.status_var.set("Queue running...")

    def stop_queue(self):
        self.queue.stop()
        self.stop_queue_btn.config(state=DISABLED)
        self.status_var.set("Stopping...")

    def remove_selected_job(self):
        selected = self.queue_tree.selection()
        if not selected:
            return
        if messagebox.askyesno("Confirm", "Remove the selected job from the queue?"):
            idx = self.queue_tree.index(selected[0])
            try:
                self.queue.remove_job(idx)
            except IndexError:
                pass
            self._refresh_queue_display()

    def _refresh_queue_display(self):
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        for job in self.queue.get_queue_snapshot():
            files_str = str(len(job.file_names))
            self.queue_tree.insert("", "end", values=(job.company_code, files_str, job.status))

    def _on_queue_job_start(self, job):
        self.root.after(0, lambda: self.log(f"Queue: Starting {job.company_code} ({len(job.file_names)} files)"))
        self._refresh_queue_display()

    def _on_queue_job_finish(self, job):
        self.root.after(0, lambda: self.log(f"Queue: Finished {job.company_code} – {job.status}"))
        self._refresh_queue_display()

    def _on_queue_empty(self):
        self.root.after(0, lambda: self.status_var.set("Queue completed"))
        self.root.after(0, lambda: self.run_queue_btn.config(state=NORMAL))
        self.root.after(0, lambda: self.stop_queue_btn.config(state=DISABLED))
        self.root.after(0, lambda: self.log("All queued jobs finished."))
        self.root.after(0, lambda: self._flash_status("Queue completed"))

    def _flash_status(self, text, count=3):
        flash_colors = [("success", "white"), ("primary", "white")]
        def step(remaining):
            if remaining > 0:
                idx = (count - remaining) % len(flash_colors)
                self.status_label.config(bootstyle=flash_colors[idx][0])
                self.status_var.set(text)
                self.root.after(300, step, remaining - 1)
            else:
                self.status_label.config(bootstyle="info")
                self.status_var.set("Ready")
        step(count)

    # ==================== OPEN OUTPUT FOLDER ====================
    def open_output_folder(self):
        """Open today's output folder for the selected company in File Explorer."""
        company = self.company_var.get()
        if not company:
            messagebox.showwarning("No Company", "Select a company first.")
            return
        date_str = datetime.today().strftime("%Y%m%d")
        path = os.path.join(self.output_folder, date_str, company)
        if not os.path.exists(path):
            messagebox.showinfo("Folder not found", f"The folder does not exist yet:\n{path}")
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open folder:\n{e}")

    # ==================== AUTO‑REFRESH LOOP ====================
    def _auto_refresh_loop(self):
        """Refresh file lists every 5 seconds if the checkbox is ticked."""
        if self.auto_refresh_var.get():
            self.refresh_file_list()
            self.ext_refresh_file_list()
        self.root.after(5000, self._auto_refresh_loop)

    # ==================== PROGRESS & DATE ====================
    def _update_progress(self, current, total):
        self.progress.pack(side=RIGHT, padx=(20, 0))
        self.progress['value'] = (current / total * 100) if total else 0
        self.status_var.set(f"Processing file {current} of {total}")

    def get_email_date(self):
        try:
            d = self.date_entry.get_date()
            if hasattr(d, 'strftime'):
                return d.strftime("%Y-%m-%d")
        except:
            pass
        raw = self.date_entry.get()
        try:
            parsed = datetime.strptime(raw, "%Y-%m-%d")
            return parsed.strftime("%Y-%m-%d")
        except:
            return datetime.today().strftime("%Y-%m-%d")

    # ==================== LOGGING ====================
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    # ==================== BUILD SCAN TAB ====================
    def build_scan_tab(self):
        paned = tk.PanedWindow(self.scan_tab, orient=tk.HORIZONTAL, sashwidth=4)
        paned.pack(fill=BOTH, expand=YES)

        # ---- LEFT SIDEBAR (with vertical split) ----
        self.sidebar = tb.Frame(paned, width=280)
        paned.add(self.sidebar, minsize=200)

        side_paned = tk.PanedWindow(self.sidebar, orient=tk.VERTICAL, sashwidth=4)
        side_paned.pack(fill=BOTH, expand=YES)

        # ---- Company section ----
        company_section = tb.Frame(side_paned, padding=10)
        side_paned.add(company_section, minsize=120)

        company_frame = tb.LabelFrame(company_section, text="Company")
        company_frame.pack(fill=X)
        self.company_var = tk.StringVar()
        self.company_combo = self.create_searchable_combobox(
            company_frame, "Company Code:", [], self.company_var
        )
        self.company_combo.pack(fill=X, pady=(2, 5))

        tb.Label(company_frame, text="Email Received Date:").pack(anchor=W)
        try:
            from ttkbootstrap.widgets import DateEntry
            self.date_entry = DateEntry(company_frame, bootstyle="primary")
            self.date_entry.set_date(datetime.today())
        except ImportError:
            self.date_entry = tb.Entry(company_frame, width=12)
            self.date_entry.insert(0, datetime.today().strftime("%Y-%m-%d"))
        self.date_entry.pack(fill=X, pady=(2, 5))

        # ---- Passwords section (buttons fixed on left) ----
        pwd_section = tb.Frame(side_paned, padding=10)
        side_paned.add(pwd_section, minsize=150)

        pwd_frame = tb.LabelFrame(pwd_section, text="Passwords")
        pwd_frame.pack(fill=BOTH, expand=YES)

        search_frame = tb.Frame(pwd_frame)
        search_frame.pack(fill=X, pady=(0, 5))
        tb.Label(search_frame, text="Search:").pack(side=LEFT)
        self.pass_search_var = tk.StringVar()
        self.pass_search_var.trace("w", lambda *a: self.refresh_password_list())
        tb.Entry(search_frame, textvariable=self.pass_search_var, bootstyle="info").pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        pwd_content = tb.Frame(pwd_frame)
        pwd_content.pack(fill=BOTH, expand=YES)

        btn_panel = tb.Frame(pwd_content, width=80)
        btn_panel.pack(side=LEFT, fill=Y, padx=(0, 5))
        btn_panel.pack_propagate(False)

        tb.Button(btn_panel, text="Select All", bootstyle="info",
                  command=self.select_all_passwords).pack(fill=X, pady=2)
        tb.Button(btn_panel, text="Clear All", bootstyle="secondary",
                  command=self.clear_all_passwords).pack(fill=X, pady=2)
        tb.Button(btn_panel, text="+ New", bootstyle="success",
                  command=self.add_new_password).pack(fill=X, pady=2)

        pwd_canvas = tk.Canvas(pwd_content, highlightthickness=0)
        pwd_scrollbar = tb.Scrollbar(pwd_content, orient=VERTICAL, command=pwd_canvas.yview)
        pwd_canvas.configure(yscrollcommand=pwd_scrollbar.set)
        pwd_scrollbar.pack(side=RIGHT, fill=Y)
        pwd_canvas.pack(side=LEFT, fill=BOTH, expand=YES)

        self.pwd_inner = tb.Frame(pwd_canvas)
        pwd_canvas.create_window((0, 0), window=self.pwd_inner, anchor=NW)
        self.pwd_inner.bind("<Configure>", lambda e: pwd_canvas.configure(scrollregion=pwd_canvas.bbox("all")))
        pwd_canvas.bind("<Configure>", lambda e: pwd_canvas.itemconfig(1, width=e.width))

        self.pwd_vars = {}

        # ---- RIGHT SIDE ----
        right_frame = tb.Frame(paned)
        paned.add(right_frame, minsize=400)

        right_paned = tk.PanedWindow(right_frame, orient=tk.VERTICAL, sashwidth=4)
        right_paned.pack(fill=BOTH, expand=YES)

        top_right = tb.Frame(right_paned)
        right_paned.add(top_right, minsize=200)

        filter_bar = tb.Frame(top_right)
        filter_bar.pack(fill=X, pady=(0, 5))
        self.filter_excel = tk.BooleanVar(value=True)
        self.filter_csv = tk.BooleanVar(value=True)
        self.filter_archive = tk.BooleanVar(value=True)
        tb.Checkbutton(filter_bar, text="Excel", variable=self.filter_excel, bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Checkbutton(filter_bar, text="CSV", variable=self.filter_csv, bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Checkbutton(filter_bar, text="Archives", variable=self.filter_archive, bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Button(filter_bar, text="Refresh", bootstyle="secondary", command=self.refresh_file_list).pack(side=LEFT, padx=10)
        tb.Button(filter_bar, text="Select All", bootstyle="info", command=self.select_all_files).pack(side=LEFT)

        # Auto‑refresh checkbox
        self.auto_refresh_var = tk.BooleanVar(value=False)
        tb.Checkbutton(filter_bar, text="Auto‑refresh (5s)", variable=self.auto_refresh_var,
                       bootstyle="primary").pack(side=LEFT, padx=10)

        file_preview_paned = tk.PanedWindow(top_right, orient=tk.VERTICAL, sashwidth=4)
        file_preview_paned.pack(fill=BOTH, expand=YES)

        file_list_frame = tb.LabelFrame(file_preview_paned, text="Files in Watch Folder")
        file_preview_paned.add(file_list_frame, minsize=100)

        self.file_canvas = tk.Canvas(file_list_frame, highlightthickness=0)
        file_scrollbar = tb.Scrollbar(file_list_frame, orient=VERTICAL, command=self.file_canvas.yview)
        self.file_canvas.configure(yscrollcommand=file_scrollbar.set)
        file_scrollbar.pack(side=RIGHT, fill=Y)
        self.file_canvas.pack(side=LEFT, fill=BOTH, expand=YES)

        self.file_inner = tb.Frame(self.file_canvas)
        self.file_canvas.create_window((0, 0), window=self.file_inner, anchor=NW)
        self.file_inner.bind("<Configure>", lambda e: self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all")))
        self.file_canvas.bind("<Configure>", lambda e: self.file_canvas.itemconfig(1, width=e.width))

        self.file_vars = {}

        # Enable drag‑and‑drop on file list areas
        self._setup_drop_target(self.file_canvas)
        self._setup_drop_target(self.file_inner)

        preview_frame = tb.LabelFrame(file_preview_paned, text="Preview (first 20 rows)")
        file_preview_paned.add(preview_frame, minsize=100)

        # Add sheet selector combobox (hidden by default)
        self.preview_sheet_var = tk.StringVar()
        self.preview_sheet_combo = tb.Combobox(preview_frame, textvariable=self.preview_sheet_var,
                                               bootstyle="secondary", state="readonly")

        preview_frame.grid_rowconfigure(1, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_tree = tb.Treeview(preview_frame, bootstyle="secondary")
        preview_scroll_y = tb.Scrollbar(preview_frame, orient=VERTICAL, command=self.preview_tree.yview)
        preview_scroll_x = tb.Scrollbar(preview_frame, orient=HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)
        self.preview_tree.grid(row=1, column=0, sticky="nsew")
        preview_scroll_y.grid(row=1, column=1, sticky="ns")
        preview_scroll_x.grid(row=2, column=0, sticky="ew")

        # ---- Bottom Right (controls, queue, log) ----
        bottom_right = tb.Frame(right_paned)
        right_paned.add(bottom_right, minsize=150)

        self.detect_btn = tb.Button(bottom_right, text="Detect Headers (Optional)", bootstyle="primary",
                                    command=self.open_mapping_popup)
        self.detect_btn.pack(anchor=W, pady=(0, 5))

        btn_row = tb.Frame(bottom_right)
        btn_row.pack(fill=X, pady=(5, 5))
        self.add_to_queue_btn = tb.Button(btn_row, text="Add to Queue", bootstyle="info",
                                         command=self.add_to_queue, state=DISABLED)
        self.add_to_queue_btn.pack(side=LEFT, padx=2)
        self.run_queue_btn = tb.Button(btn_row, text="Run Queue", bootstyle="success",
                                      command=self.run_queue)
        self.run_queue_btn.pack(side=LEFT, padx=2)
        self.stop_queue_btn = tb.Button(btn_row, text="Stop", bootstyle="danger",
                                       command=self.stop_queue, state=DISABLED)
        self.stop_queue_btn.pack(side=LEFT, padx=2)
        # Open Output Folder button
        self.open_folder_btn = tb.Button(btn_row, text="Open Output", bootstyle="secondary",
                                         command=self.open_output_folder)
        self.open_folder_btn.pack(side=LEFT, padx=2)

        queue_frame = tb.LabelFrame(bottom_right, text="Job Queue")
        queue_frame.pack(fill=X, pady=(5, 0))
        self.queue_tree = tb.Treeview(queue_frame, columns=("company", "files", "status"),
                                      show="headings", height=3, bootstyle="secondary")
        self.queue_tree.heading("company", text="Company")
        self.queue_tree.heading("files", text="Files")
        self.queue_tree.heading("status", text="Status")
        self.queue_tree.column("company", width=100, stretch=False)
        self.queue_tree.column("files", width=60, stretch=False)
        self.queue_tree.column("status", width=100, stretch=False)
        self.queue_tree.pack(fill=X, expand=YES, side=LEFT)
        queue_scroll = tb.Scrollbar(queue_frame, orient=VERTICAL, command=self.queue_tree.yview)
        queue_scroll.pack(side=RIGHT, fill=Y)
        self.queue_tree.configure(yscrollcommand=queue_scroll.set)
        remove_btn = tb.Button(queue_frame, text="Remove Selected", bootstyle="warning",
                               command=self.remove_selected_job)
        remove_btn.pack(side=BOTTOM, pady=2)

        log_frame = tb.LabelFrame(bottom_right, text="Log")
        log_frame.pack(fill=BOTH, expand=YES, pady=(5, 0))
        from ttkbootstrap.widgets.scrolled import ScrolledText
        self.log_text = ScrolledText(log_frame, height=6, wrap=WORD)
        self.log_text.pack(fill=BOTH, expand=YES)

    # ==================== BUILD EXTRACT HEADERS TAB ====================
    def build_extract_tab(self):
        paned = tk.PanedWindow(self.extract_tab, orient=tk.HORIZONTAL, sashwidth=4)
        paned.pack(fill=BOTH, expand=YES)

        sidebar = tb.Frame(paned, width=280)
        paned.add(sidebar, minsize=200)

        side_paned = tk.PanedWindow(sidebar, orient=tk.VERTICAL, sashwidth=4)
        side_paned.pack(fill=BOTH, expand=YES)

        company_section = tb.Frame(side_paned, padding=10)
        side_paned.add(company_section, minsize=120)

        company_frame = tb.LabelFrame(company_section, text="Company")
        company_frame.pack(fill=X)
        self.ext_company_var = tk.StringVar()
        self.ext_company_combo = self.create_searchable_combobox(
            company_frame, "Company Code:", [], self.ext_company_var
        )
        self.ext_company_combo.pack(fill=X, pady=(2, 5))

        tb.Label(company_frame, text="Extraction Date:").pack(anchor=W)
        try:
            from ttkbootstrap.widgets import DateEntry
            self.ext_date_entry = DateEntry(company_frame, bootstyle="primary")
            self.ext_date_entry.set_date(datetime.today())
        except ImportError:
            self.ext_date_entry = tb.Entry(company_frame, width=12)
            self.ext_date_entry.insert(0, datetime.today().strftime("%Y-%m-%d"))
        self.ext_date_entry.pack(fill=X, pady=(2, 5))

        pwd_section = tb.Frame(side_paned, padding=10)
        side_paned.add(pwd_section, minsize=150)

        pwd_frame = tb.LabelFrame(pwd_section, text="Passwords")
        pwd_frame.pack(fill=BOTH, expand=YES)

        search_frame = tb.Frame(pwd_frame)
        search_frame.pack(fill=X, pady=(0, 5))
        tb.Label(search_frame, text="Search:").pack(side=LEFT)
        self.ext_pass_search_var = tk.StringVar()
        self.ext_pass_search_var.trace("w", lambda *a: self.ext_refresh_password_list())
        tb.Entry(search_frame, textvariable=self.ext_pass_search_var, bootstyle="info").pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        pwd_content = tb.Frame(pwd_frame)
        pwd_content.pack(fill=BOTH, expand=YES)

        btn_panel = tb.Frame(pwd_content, width=80)
        btn_panel.pack(side=LEFT, fill=Y, padx=(0, 5))
        btn_panel.pack_propagate(False)

        tb.Button(btn_panel, text="Select All", bootstyle="info",
                  command=self.ext_select_all_passwords).pack(fill=X, pady=2)
        tb.Button(btn_panel, text="Clear All", bootstyle="secondary",
                  command=self.ext_clear_all_passwords).pack(fill=X, pady=2)

        pwd_canvas = tk.Canvas(pwd_content, highlightthickness=0)
        pwd_scrollbar = tb.Scrollbar(pwd_content, orient=VERTICAL, command=pwd_canvas.yview)
        pwd_canvas.configure(yscrollcommand=pwd_scrollbar.set)
        pwd_scrollbar.pack(side=RIGHT, fill=Y)
        pwd_canvas.pack(side=LEFT, fill=BOTH, expand=YES)

        self.ext_pwd_inner = tb.Frame(pwd_canvas)
        pwd_canvas.create_window((0, 0), window=self.ext_pwd_inner, anchor=NW)
        self.ext_pwd_inner.bind("<Configure>", lambda e: pwd_canvas.configure(scrollregion=pwd_canvas.bbox("all")))
        pwd_canvas.bind("<Configure>", lambda e: pwd_canvas.itemconfig(1, width=e.width))

        self.ext_pwd_vars = {}

        # ---- RIGHT SIDE ----
        right_frame = tb.Frame(paned)
        paned.add(right_frame, minsize=400)

        right_paned = tk.PanedWindow(right_frame, orient=tk.VERTICAL, sashwidth=4)
        right_paned.pack(fill=BOTH, expand=YES)

        top_right = tb.Frame(right_paned)
        right_paned.add(top_right, minsize=200)

        filter_bar = tb.Frame(top_right)
        filter_bar.pack(fill=X, pady=(5, 5))
        self.ext_filter_excel = tk.BooleanVar(value=True)
        self.ext_filter_csv = tk.BooleanVar(value=True)
        self.ext_filter_archive = tk.BooleanVar(value=True)
        tb.Checkbutton(filter_bar, text="Excel", variable=self.ext_filter_excel, bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Checkbutton(filter_bar, text="CSV", variable=self.ext_filter_csv, bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Checkbutton(filter_bar, text="Archives", variable=self.ext_filter_archive, bootstyle="primary").pack(side=LEFT, padx=5)
        tb.Button(filter_bar, text="Refresh", bootstyle="secondary", command=self.ext_refresh_file_list).pack(side=LEFT, padx=10)
        tb.Button(filter_bar, text="Select All", bootstyle="info", command=self.ext_select_all_files).pack(side=LEFT)

        file_preview_paned = tk.PanedWindow(top_right, orient=tk.VERTICAL, sashwidth=4)
        file_preview_paned.pack(fill=BOTH, expand=YES)

        file_list_frame = tb.LabelFrame(file_preview_paned, text="Files in Watch Folder")
        file_preview_paned.add(file_list_frame, minsize=100)

        self.ext_file_canvas = tk.Canvas(file_list_frame, highlightthickness=0)
        ext_file_scrollbar = tb.Scrollbar(file_list_frame, orient=VERTICAL, command=self.ext_file_canvas.yview)
        self.ext_file_canvas.configure(yscrollcommand=ext_file_scrollbar.set)
        ext_file_scrollbar.pack(side=RIGHT, fill=Y)
        self.ext_file_canvas.pack(side=LEFT, fill=BOTH, expand=YES)

        self.ext_file_inner = tb.Frame(self.ext_file_canvas)
        self.ext_file_canvas.create_window((0, 0), window=self.ext_file_inner, anchor=NW)
        self.ext_file_inner.bind("<Configure>", lambda e: self.ext_file_canvas.configure(scrollregion=self.ext_file_canvas.bbox("all")))
        self.ext_file_canvas.bind("<Configure>", lambda e: self.ext_file_canvas.itemconfig(1, width=e.width))

        self.ext_file_vars = {}

        # Enable drag‑and‑drop on extract tab file list
        self._setup_drop_target(self.ext_file_canvas)
        self._setup_drop_target(self.ext_file_inner)

        preview_frame = tb.LabelFrame(file_preview_paned, text="Preview (first 20 rows)")
        file_preview_paned.add(preview_frame, minsize=100)

        # Sheet selector for extract tab
        self.ext_preview_sheet_var = tk.StringVar()
        self.ext_preview_sheet_combo = tb.Combobox(preview_frame, textvariable=self.ext_preview_sheet_var,
                                                   bootstyle="secondary", state="readonly")

        preview_frame.grid_rowconfigure(1, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.ext_preview_tree = tb.Treeview(preview_frame, bootstyle="secondary")
        ext_preview_scroll_y = tb.Scrollbar(preview_frame, orient=VERTICAL, command=self.ext_preview_tree.yview)
        ext_preview_scroll_x = tb.Scrollbar(preview_frame, orient=HORIZONTAL, command=self.ext_preview_tree.xview)
        self.ext_preview_tree.configure(yscrollcommand=ext_preview_scroll_y.set, xscrollcommand=ext_preview_scroll_x.set)
        self.ext_preview_tree.grid(row=1, column=0, sticky="nsew")
        ext_preview_scroll_y.grid(row=1, column=1, sticky="ns")
        ext_preview_scroll_x.grid(row=2, column=0, sticky="ew")

        bottom_right = tb.Frame(right_paned)
        right_paned.add(bottom_right, minsize=100)

        action_frame = tb.Frame(bottom_right)
        action_frame.pack(fill=X, pady=(5, 5))
        self.extract_headers_btn = tb.Button(action_frame, text="Extract Headers", bootstyle="success",
                                             command=self.run_extract_headers)
        self.extract_headers_btn.pack(side=LEFT, padx=5)

        log_frame = tb.LabelFrame(bottom_right, text="Log")
        log_frame.pack(fill=BOTH, expand=YES)
        from ttkbootstrap.widgets.scrolled import ScrolledText
        self.ext_log_text = ScrolledText(log_frame, height=6, wrap=WORD)
        self.ext_log_text.pack(fill=BOTH, expand=YES)

    # ==================== EXTRACT TAB: HELPERS ====================
    def ext_populate_company_combo(self):
        codes = sorted({d['Code'] for d in self.company_data})
        self.ext_company_combo.values = codes
        if codes:
            self.ext_company_combo.entry_var.set(codes[0])
            self.ext_company_var.set(codes[0])
            self.ext_refresh_password_list()

    def ext_refresh_password_list(self):
        selected = self.ext_company_var.get()
        all_pwds = [d['Password'] for d in self.company_data if d['Code'] == selected]
        search = self.ext_pass_search_var.get().strip().lower()
        filtered = [p for p in all_pwds if search in p.lower()] if search else all_pwds

        for widget in self.ext_pwd_inner.winfo_children():
            widget.destroy()
        self.ext_pwd_vars.clear()
        if not filtered:
            tb.Label(self.ext_pwd_inner, text="No passwords found", foreground="gray").pack(pady=10)
            self.status_var.set("Select at least one password to enable scanning")
        else:
            for pwd in filtered:
                var = tk.IntVar(value=0)
                self.ext_pwd_vars[pwd] = var
                row = tb.Frame(self.ext_pwd_inner)
                row.pack(fill=X, pady=2)
                tb.Checkbutton(row, variable=var, bootstyle="primary").pack(side=LEFT, padx=5)
                tb.Label(row, text=pwd, anchor="w").pack(side=LEFT, fill=X, expand=True)
            self.status_var.set("Ready")

    def ext_select_all_passwords(self):
        for var in self.ext_pwd_vars.values():
            var.set(1)

    def ext_clear_all_passwords(self):
        for var in self.ext_pwd_vars.values():
            var.set(0)

    def ext_get_selected_passwords(self):
        return [pwd for pwd, var in self.ext_pwd_vars.items() if var.get() == 1]

    def ext_refresh_file_list(self):
        if not self.watch_folder or not os.path.isdir(self.watch_folder):
            return
        for widget in self.ext_file_inner.winfo_children():
            widget.destroy()
        self.ext_file_vars.clear()

        try:
            items = os.listdir(self.watch_folder)
        except Exception as e:
            self.ext_log(f"Cannot read watch folder: {e}")
            return

        files = []
        for f in items:
            full = os.path.join(self.watch_folder, f)
            if os.path.isfile(full) and not f.endswith('.json'):
                ext = f.split('.')[-1].lower()
                if (ext in FILE_EXTENSIONS_EXCEL and not self.ext_filter_excel.get()) or \
                   (ext in FILE_EXTENSIONS_TEXT and not self.ext_filter_csv.get()) or \
                   (ext in FILE_EXTENSIONS_ARCHIVE and not self.ext_filter_archive.get()):
                    continue
                size = os.path.getsize(full) // 1024
                mod_time = datetime.fromtimestamp(os.path.getmtime(full)).strftime("%Y-%m-%d %H:%M")
                ftype = "Excel" if ext in FILE_EXTENSIONS_EXCEL else "Text" if ext in FILE_EXTENSIONS_TEXT else "Archive"
                files.append((f, size, mod_time, ftype))

        if not files:
            tb.Label(self.ext_file_inner, text="No files found in watch folder.",
                     font=("Helvetica", 11), foreground="gray").pack(pady=20)
        else:
            files.sort(key=lambda x: x[0])
            for fname, size, mod_time, ftype in files:
                var = tk.IntVar(value=0)
                self.ext_file_vars[fname] = var
                row = tb.Frame(self.ext_file_inner)
                row.pack(fill=X, pady=2)
                cb = tb.Checkbutton(row, variable=var, bootstyle="primary")
                cb.pack(side=LEFT, padx=5)
                lbl = tb.Label(row, text=fname, font=("Helvetica", 10, "bold"), anchor="w")
                lbl.pack(side=LEFT, fill=X, expand=True, padx=5)
                lbl.bind("<Button-1>", lambda e, f=fname: self.ext_show_preview(f))
                meta = tb.Label(row, text=f"{size} KB | {mod_time} | {ftype}", font=("Helvetica", 9), foreground="gray")
                meta.pack(side=RIGHT, padx=5)

    def ext_select_all_files(self):
        for var in self.ext_file_vars.values():
            var.set(1)

    def ext_show_preview(self, filename, sheet_name=None):
        """Show preview of a file in Extract tab, optionally specifying the sheet for Excel files."""
        filepath = os.path.join(self.watch_folder, filename)
        self.ext_preview_sheet_combo.pack_forget()
        self.ext_preview_sheet_combo['values'] = []
        self.ext_preview_tree.delete(*self.ext_preview_tree.get_children())
        self.ext_preview_tree["columns"] = ()
        if not os.path.exists(filepath):
            return
        try:
            ext = filename.split('.')[-1].lower()
            if ext in FILE_EXTENSIONS_EXCEL:
                all_sheets = pl.read_excel(filepath, sheet_id=0).keys()
                sheet_names = list(all_sheets)
                if not sheet_names:
                    return
                if len(sheet_names) > 1:
                    self.ext_preview_sheet_combo['values'] = sheet_names
                    current_sheet = sheet_name if sheet_name in sheet_names else sheet_names[0]
                    self.ext_preview_sheet_var.set(current_sheet)
                    self.ext_preview_sheet_combo.pack(fill=X, padx=5, pady=(5,0), before=self.ext_preview_tree)
                    self.ext_preview_sheet_combo.bind('<<ComboboxSelected>>',
                        lambda e: self.ext_show_preview(filename, self.ext_preview_sheet_var.get()))
                else:
                    current_sheet = sheet_names[0]
                df = pl.read_excel(filepath, sheet_name=current_sheet, has_header=False)
                sample = df.head(20)
            elif ext in FILE_EXTENSIONS_TEXT:
                df = pl.read_csv(filepath, has_header=False, truncate_ragged_lines=True)
                sample = df.head(20)
            else:
                return

            col_names = [f"Col {i}" for i in range(sample.width)]
            self.ext_preview_tree["columns"] = col_names
            self.ext_preview_tree.heading("#0", text="Row")
            self.ext_preview_tree.column("#0", width=50, stretch=False)
            for col in col_names:
                self.ext_preview_tree.heading(col, text=col)
                self.ext_preview_tree.column(col, width=100, stretch=False)
            for row_idx in range(sample.height):
                values = [str(sample[row_idx, i]) for i in range(sample.width)]
                self.ext_preview_tree.insert("", "end", text=str(row_idx+1), values=values)
        except Exception as e:
            self.ext_log(f"Preview error: {e}")

    def ext_log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.ext_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.ext_log_text.see(tk.END)

    # ==================== RUN HEADER EXTRACTION ====================
    def run_extract_headers(self):
        selected_files = [f for f, var in self.ext_file_vars.items() if var.get() == 1]
        if not selected_files:
            messagebox.showwarning("Warning", "No files selected.")
            return
        company = self.ext_company_var.get()
        passwords = self.ext_get_selected_passwords()
        if not company or not passwords:
            messagebox.showerror("Error", "Select a company and at least one password.")
            return

        try:
            d = self.ext_date_entry.get_date()
            date_str = d.strftime("%m%d%Y") if hasattr(d, 'strftime') else d.strftime("%Y%m%d")
        except:
            raw = self.ext_date_entry.get()
            try:
                parsed = datetime.strptime(raw, "%Y-%m-%d")
                date_str = parsed.strftime("%Y%m%d")
            except:
                date_str = datetime.today().strftime("%Y%m%d")

        output_dir = os.path.join(self.output_folder, f"HD_extract_{date_str}_{company}")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"{date_str}_headerextract_{company}.xlsx")

        self.ext_log(f"Extracting headers for {company} – output: {output_file}")

        try:
            header_dict, _ = get_company_header(company)
        except Exception as e:
            self.ext_log(f"Header dictionary error: {e}")
            return

        all_headers = []
        for fname in selected_files:
            filepath = os.path.join(self.watch_folder, fname)
            self.ext_log(f"Inspecting {fname}...")
            try:
                structure = inspect_file(filepath, passwords, output_dir)
            except Exception as e:
                self.ext_log(f"Inspector failed: {e}")
                continue

            if not structure.tables:
                self.ext_log(f"No tables found in {fname}")
                continue

            for table in structure.tables:
                mapping = classify_columns(table, header_dict)
                for col_map in mapping.mappings:
                    all_headers.append({
                        'File Path': filepath,
                        'Sheet Name': table.sheet_name,
                        'Table': f"Table {table.col_start}-{table.col_end}",
                        'Header': col_map.header_raw,
                        'Column Index': col_map.column_index,
                        'Category': col_map.category,
                    })
                for i, raw_hdr in enumerate(table.headers):
                    if not any(cm.column_index == i for cm in mapping.mappings):
                        all_headers.append({
                            'File Path': filepath,
                            'Sheet Name': table.sheet_name,
                            'Table': f"Table {table.col_start}-{table.col_end}",
                            'Header': raw_hdr,
                            'Column Index': i,
                            'Category': 'Unknown',
                        })

        if all_headers:
            try:
                df = pl.DataFrame(all_headers)
                df = df.sort(['File Path', 'Sheet Name', 'Table', 'Column Index'])
                df.write_excel(output_file)
                self.ext_log(f"Extraction complete. {len(all_headers)} headers saved to {output_file}")
                messagebox.showinfo("Done", f"Headers extracted to:\n{output_file}")
            except Exception as e:
                csv_file = output_file.replace('.xlsx', '.csv')
                pl.DataFrame(all_headers).write_csv(csv_file)
                self.ext_log(f"Excel write failed: {e}. Saved as CSV: {csv_file}")
        else:
            self.ext_log("No headers found in any file.")

    # ==================== BUILD SETTINGS TAB ====================
    def build_settings_tab(self):
        frame = tb.Frame(self.settings_tab, padding=20)
        frame.pack(fill=BOTH, expand=YES)

        tb.Label(frame, text="Input (Watch) Folder:", font=("Helvetica", 11)).grid(row=0, column=0, sticky=W, pady=5)
        self.input_folder_var = tk.StringVar(value=self.watch_folder)
        tb.Entry(frame, textvariable=self.input_folder_var, width=60).grid(row=0, column=1, padx=5, pady=5)
        tb.Button(frame, text="Browse", command=lambda: self._browse_folder(self.input_folder_var)).grid(row=0, column=2, padx=5)

        tb.Label(frame, text="Company Passwords CSV:", font=("Helvetica", 11)).grid(row=1, column=0, sticky=W, pady=5)
        self.password_csv_var = tk.StringVar(value=self.password_csv_path)
        tb.Entry(frame, textvariable=self.password_csv_var, width=60).grid(row=1, column=1, padx=5, pady=5)
        tb.Button(frame, text="Browse", command=lambda: self._browse_file(self.password_csv_var, [("CSV Files", "*.csv")])).grid(row=1, column=2, padx=5)

        tb.Label(frame, text="Output Folder:", font=("Helvetica", 11)).grid(row=2, column=0, sticky=W, pady=5)
        self.output_folder_var = tk.StringVar(value=self.output_folder)
        tb.Entry(frame, textvariable=self.output_folder_var, width=60).grid(row=2, column=1, padx=5, pady=5)
        tb.Button(frame, text="Browse", command=lambda: self._browse_folder(self.output_folder_var)).grid(row=2, column=2, padx=5)

        tb.Label(frame, text="Header Files Folder:", font=("Helvetica", 11)).grid(row=3, column=0, sticky=W, pady=5)
        self.header_folder_var = tk.StringVar(value=self.header_folder)
        tb.Entry(frame, textvariable=self.header_folder_var, width=60).grid(row=3, column=1, padx=5, pady=5)
        tb.Button(frame, text="Browse", command=lambda: self._browse_folder(self.header_folder_var)).grid(row=3, column=2, padx=5)

        tb.Label(frame, text="Theme:", font=("Helvetica", 11)).grid(row=4, column=0, sticky=W, pady=5)
        self.theme_var = tk.StringVar(value=self.settings.get("theme", "flatly"))
        themes = ["flatly", "darkly", "cyborg", "solar", "superhero", "journal", "litera", "lumen", "minty", "pulse", "sandstone", "simplex", "spacelab", "united", "yeti"]
        tb.Combobox(frame, textvariable=self.theme_var, values=themes, bootstyle="primary").grid(row=4, column=1, padx=5, pady=5, sticky=W)

        tb.Button(frame, text="Save Settings", bootstyle="success", command=self.save_settings).grid(row=5, column=1, pady=20, sticky=E)

    def _browse_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _browse_file(self, var, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    def save_settings(self):
        new_settings = {
            "folder": self.input_folder_var.get(),
            "csv": self.password_csv_var.get(),
            "output_folder": self.output_folder_var.get(),
            "header_folder": self.header_folder_var.get(),
            "theme": self.theme_var.get(),
        }
        save_settings(new_settings)
        self.watch_folder = new_settings["folder"]
        self.password_csv_path = new_settings["csv"]
        self.output_folder = new_settings["output_folder"]
        self.header_folder = new_settings["header_folder"]
        set_header_path(self.header_folder)
        self.load_company_data()
        try:
            self.root.style.theme_use(new_settings["theme"])
        except:
            pass
        messagebox.showinfo("Settings", "Settings saved.")
        self.refresh_file_list()
        self.refresh_password_list()
        self.ext_refresh_file_list()
        self.ext_refresh_password_list()

    # ==================== BUILD HISTORY TAB ====================
    def build_history_tab(self):
        frame = tb.Frame(self.history_tab, padding=10)
        frame.pack(fill=BOTH, expand=YES)
        tb.Label(frame, text="Past Scans", font=("Helvetica", 14, "bold")).pack(anchor=W, pady=(0,5))
        columns = ("date", "company", "files", "output")
        self.history_tree = tb.Treeview(frame, columns=columns, show="headings", bootstyle="primary")
        self.history_tree.heading("date", text="Scan Date")
        self.history_tree.heading("company", text="Company")
        self.history_tree.heading("files", text="Files Processed")
        self.history_tree.heading("output", text="Output Folder")
        self.history_tree.pack(fill=BOTH, expand=YES)
        scroll = tb.Scrollbar(frame, orient=VERTICAL, command=self.history_tree.yview)
        scroll.pack(side=RIGHT, fill=Y)
        self.history_tree.configure(yscrollcommand=scroll.set)
        tb.Button(frame, text="Refresh", bootstyle="secondary", command=self.refresh_history).pack(pady=5)
        self.refresh_history()

    def refresh_history(self):
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        output_root = self.output_folder
        if not output_root or not os.path.exists(output_root):
            return
        try:
            items = os.listdir(output_root)
        except Exception:
            return
        for folder in items:
            folder_path = os.path.join(output_root, folder)
            if os.path.isdir(folder_path):
                log_file = os.path.join(folder_path, f"Log_{folder}.csv")
                if os.path.exists(log_file):
                    try:
                        df = pl.read_csv(log_file)
                        company = df['Company Code'].unique().to_list()[0] if 'Company Code' in df.columns else "Unknown"
                        file_count = df['File Name'].n_unique() if 'File Name' in df.columns else 0
                        self.history_tree.insert("", tk.END, values=(folder, company, file_count, folder_path))
                    except:
                        pass

if __name__ == "__main__":
    # Use TkinterDnD.Tk for drag‑and‑drop support
    root = TkinterDnD.Tk()
    # Create ttkbootstrap Style and attach it to root
    settings = load_settings()
    theme = settings.get("theme", "flatly")
    style = tb.Style(theme=theme)
    root.style = style
    # Start the app
    app = OFACScannerApp(root)
    root.mainloop()
