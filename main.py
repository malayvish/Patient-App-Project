import os
import sys
import shutil
import webbrowser
import subprocess
from datetime import datetime
import urllib.parse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import pandas as pd
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as ReportImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO

def resource_path(relative_path):
    
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# ========================= CONSTANTS =========================
DATA_FILE = resource_path("patients.xlsx")
PHOTO_DIR = resource_path("photos")
PAGE_SIZE = 24
DEFAULT_COLUMNS = [
    'SerialNo', 'PhotoPath', 'Name', 'Email', 'Gender', 'Age', 'Address',
    'PhoneNo', 'Occupation', 'AadharNo', 'Symptoms', 'Treatment',
    'StartDate', 'EndDate', 'Satisfied'
]


# ========================= HELPER FUNCTIONS =========================
def ensure_datafile():
    if not os.path.exists(DATA_FILE):
        os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
        df = pd.DataFrame(columns=DEFAULT_COLUMNS)
        df.to_excel(DATA_FILE, index=False)

def get_most_common(df, column):
    if df.empty or column not in df.columns:
        return 'N/A'
    counts = df[column].astype(str).value_counts()
    return counts.index[0] if not counts.empty else 'N/A'

def get_average_duration(df):
    if df.empty or 'StartDate' not in df.columns or 'EndDate' not in df.columns:
        return 'N/A'
    try:
        start_dates = pd.to_datetime(df['StartDate'], errors='coerce')
        end_dates = pd.to_datetime(df['EndDate'], errors='coerce')
        valid_durations = (end_dates - start_dates).dt.days
        mean_duration = valid_durations.mean()
        return f"{mean_duration:.1f} days" if pd.notna(mean_duration) else 'N/A'
    except Exception:
        return 'N/A'

def responsive_pack(widget, **kwargs):
    kwargs.setdefault('padx', 10)
    kwargs.setdefault('pady', 10)
    widget.pack(**kwargs)

# ========================= MAIN APPLICATION CLASS =========================
class PatientManagementSystem(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Navjeevan Arogya Sanstha - Patient Management System")
        self.geometry("1000x700")
        self.configure(bg='#f0f8ff')
        ensure_datafile()
        self.patients_df = self.load_patients()
        self.current_frame = None
        self.show_main_menu()

    def load_patients(self):
        try:
            df = pd.read_excel(DATA_FILE, dtype={'SerialNo': str, 'AadharNo': str, 'PhoneNo': str})
            for c in DEFAULT_COLUMNS:
                if c not in df.columns:
                    df[c] = ''
            return df[DEFAULT_COLUMNS]
        except Exception:
            return pd.DataFrame(columns=DEFAULT_COLUMNS)

    def save_patients(self):
        try:
            self.patients_df.to_excel(DATA_FILE, index=False)
            self.patients_df = self.load_patients() # Reload to maintain data type consistency
        except Exception as e:
            messagebox.showerror("Error", f"Could not save patients.xlsx: {e}")

    def clear_frame(self):
        if self.current_frame:
            self.current_frame.destroy()
            self.current_frame = None

    def show_main_menu(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=20, pady=20)
        self.current_frame = frame

        header_frame = tk.Frame(frame, bg='#2c5aa0')
        header_frame.pack(fill='x', pady=(0, 20))
        tk.Label(header_frame, text="Navjeevan Arogya Sanstha", font=("Arial", 28, "bold"),
                 fg='white', bg='#2c5aa0').pack(pady=20)
        tk.Label(header_frame, text="Patient Management System", font=("Arial", 14),
                 fg='white', bg='#2c5aa0').pack(pady=(0, 12))

        button_frame = tk.Frame(frame, bg='#f0f8ff')
        button_frame.pack(expand=True)

        buttons = [
            ("➕ New Patient Registration", self.show_new_patient),
            ("✏️ Edit Patients", self.show_edit_patients),
            ("🗑️ Delete Patients", self.show_delete_patients),
            ("📋 View All Records", self.show_view_records),
            ("📤 Share / Export Details", self.show_share_details),
            ("⚙️ Manage Patients", self.show_manage_patient),
            ("🔍 Manage Duplicates", self.show_duplicate_page),
            ("🚪 Exit", self.destroy)
        ]

        for i, (text, cmd) in enumerate(buttons):
            btn = tk.Button(button_frame, text=text, width=28, height=2,
                            font=("Arial", 12), bg='#4c72b0', fg='white',
                            command=cmd, relief='raised', bd=3)
            btn.grid(row=i // 2, column=i % 2, padx=20, pady=10)
            btn.bind("<Enter>", lambda e, b=btn: b.config(bg='#3a5a8a'))
            btn.bind("<Leave>", lambda e, b=btn: b.config(bg='#4c72b0'))

    def show_new_patient(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=20, pady=20)
        self.current_frame = frame

        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 10))
        tk.Label(header, text="New Patient Registration", font=("Arial", 18, "bold"),
                 fg='white', bg='#2c5aa0').pack(pady=10)

        canvas = tk.Canvas(frame, bg='#f0f8ff', highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f8ff')
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        try:
            numeric_serials = pd.to_numeric(self.patients_df['SerialNo'], errors='coerce')
            next_serial = int(numeric_serials.max()) + 1 if not numeric_serials.dropna().empty else 1
        except Exception:
            next_serial = 1

        fields = [
            ("Serial No:", 'serial', 'entry', str(next_serial)),
            ("Photo:", 'photo', 'button', None),
            ("Name *:", 'name', 'entry', ''),
            ("Email *:", 'email', 'entry', ''),
            ("Gender *:", 'gender', 'combo', ['Male', 'Female', 'Other']),
            ("Age *:", 'age', 'entry', ''),
            ("Address:", 'address', 'entry', ''),
            ("Phone No *:", 'phone', 'entry', ''),
            ("Occupation:", 'occupation', 'entry', ''),
            ("Aadhar No:", 'aadhar', 'entry', ''),
            ("Symptoms:", 'symptoms', 'text', ''),
            ("Treatment:", 'treatment', 'text', ''),
            ("Start Date (YYYY-MM-DD):", 'start_date', 'entry', datetime.today().strftime('%Y-%m-%d')),
            ("End Date (YYYY-MM-DD):", 'end_date', 'entry', ''),
            ("Satisfied with Treatment:", 'satisfied', 'combo', ['Yes', 'No', 'Not Sure'])
        ]

        self.new_vars = {}
        self.new_photo = None

        for label, key, typ, default in fields:
            row = tk.Frame(scrollable, bg='#f0f8ff')
            row.pack(fill='x', pady=6)
            tk.Label(row, text=label, width=28, anchor='w', bg='#f0f8ff').pack(side='left', padx=(0, 10))

            if typ == 'entry':
                var = tk.StringVar(value=default)
                tk.Entry(row, textvariable=var, width=36).pack(side='left', fill='x', expand=True)
                self.new_vars[key] = var
            elif typ == 'combo':
                var = tk.StringVar()
                ttk.Combobox(row, textvariable=var, values=default, width=34, state='readonly').pack(side='left', fill='x', expand=True)
                self.new_vars[key] = var
            elif typ == 'text':
                txt = scrolledtext.ScrolledText(row, height=4, width=36)
                txt.pack(side='left', fill='both', expand=True)
                self.new_vars[key] = txt
            elif typ == 'button':
                tk.Button(row, text="Add Photo", bg='#4c72b0', fg='white', command=self._new_add_photo).pack(side='left')

        btns = tk.Frame(frame, bg='#f0f8ff')
        btns.pack(fill='x', pady=10)
        tk.Button(btns, text="Submit", bg='#28a745', fg='white', width=12, command=self._new_submit).pack(side='left', padx=6, expand=True)
        tk.Button(btns, text="Clear", bg='#ffc107', width=12, command=self._new_clear).pack(side='left', padx=6, expand=True)
        tk.Button(btns, text="Back", bg='#6c757d', fg='white', width=12, command=self.show_main_menu).pack(side='left', padx=6, expand=True)

    def _new_add_photo(self):
        path = filedialog.askopenfilename(title="Select Photo", filetypes=[("Image", "*.jpg *.jpeg *.png *.bmp")])
        if path:
            self.new_photo = path
            messagebox.showinfo("Photo", "Photo selected")

    def _new_submit(self):
        required = ['name', 'email', 'gender', 'age', 'phone']
        for r in required:
            v = self.new_vars.get(r)
            if isinstance(v, tk.StringVar):
                if not v.get().strip():
                    messagebox.showerror("Error", f"Please fill {r}")
                    return
            elif isinstance(v, scrolledtext.ScrolledText):
                if not v.get("1.0", "end").strip():
                    messagebox.showerror("Error", f"Please fill {r}")
                    return
        
        try:
            int(self.new_vars['age'].get())
        except ValueError:
            messagebox.showerror("Error", "Age must be a numeric value.")
            return

        try:
            serial = int(self.new_vars['serial'].get())
            if str(serial) in self.patients_df['SerialNo'].values:
                messagebox.showerror("Error", f"Serial number {serial} already exists.")
                return
        except ValueError:
            messagebox.showerror("Error", "Serial must be numeric")
            return

        new_patient = {
            'SerialNo': serial, 'PhotoPath': '', 'Name': self.new_vars['name'].get().strip(),
            'Email': self.new_vars['email'].get().strip(), 'Gender': self.new_vars['gender'].get(),
            'Age': int(self.new_vars['age'].get()), 'Address': self.new_vars['address'].get().strip(),
            'PhoneNo': self.new_vars['phone'].get().strip(), 'Occupation': self.new_vars['occupation'].get().strip(),
            'AadharNo': self.new_vars['aadhar'].get().strip(),
            'Symptoms': self.new_vars['symptoms'].get("1.0", "end").strip(),
            'Treatment': self.new_vars['treatment'].get("1.0", "end").strip(),
            'StartDate': self.new_vars['start_date'].get(), 'EndDate': self.new_vars['end_date'].get(),
            'Satisfied': self.new_vars['satisfied'].get()
        }

        if self.new_photo:
            os.makedirs(PHOTO_DIR, exist_ok=True)
            fname = f"patient_{serial}.png"
            dest = os.path.join(PHOTO_DIR, fname)
            try:
                img = Image.open(self.new_photo)
                img.save(dest, format='PNG')
                new_patient['PhotoPath'] = dest
            except Exception as e:
                messagebox.showwarning("Photo", f"Could not copy photo: {e}")

        self.patients_df = pd.concat([self.patients_df, pd.DataFrame([new_patient])], ignore_index=True)
        self.save_patients()
        messagebox.showinfo("Success", "Patient registered")
        self._new_clear()

    def _new_clear(self):
        try:
            numeric_serials = pd.to_numeric(self.patients_df['SerialNo'], errors='coerce')
            next_serial = int(numeric_serials.max()) + 1 if not numeric_serials.dropna().empty else 1
        except Exception:
            next_serial = 1
        for k, v in self.new_vars.items():
            if isinstance(v, tk.StringVar): v.set('')
            elif isinstance(v, scrolledtext.ScrolledText): v.delete("1.0", "end")
        if 'serial' in self.new_vars: self.new_vars['serial'].set(str(next_serial))
        if 'start_date' in self.new_vars: self.new_vars['start_date'].set(datetime.today().strftime('%Y-%m-%d'))
        self.new_photo = None

    def show_view_records(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.current_frame = frame

        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 8))
        tk.Label(header, text="All Patient Records", font=("Arial", 18, "bold"), fg='white', bg='#2c5aa0').pack(pady=8)

        filter_frame = tk.Frame(frame, bg='#f0f8ff')
        filter_frame.pack(fill='x', pady=6)
        
        tk.Label(filter_frame, text="Search:", bg='#f0f8ff').pack(side='left', padx=6)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left')
        
        tk.Label(filter_frame, text="Gender Filter:", bg='#f0f8ff').pack(side='left', padx=(20, 6))
        self.gender_filter_var = tk.StringVar(value="All")
        gender_options = ["All"] + list(self.patients_df['Gender'].dropna().unique())
        gender_combo = ttk.Combobox(filter_frame, textvariable=self.gender_filter_var, values=gender_options, state='readonly')
        gender_combo.pack(side='left')
        
        tk.Label(filter_frame, text="Sort By:", bg='#f0f8ff').pack(side='left', padx=(20, 6))
        self.sort_by_var = tk.StringVar(value="SerialNo")
        sort_options = ['SerialNo', 'Name', 'Age', 'StartDate', 'EndDate']
        sort_combo = ttk.Combobox(filter_frame, textvariable=self.sort_by_var, values=sort_options, state='readonly')
        sort_combo.pack(side='left')

        action_frame = tk.Frame(frame, bg='#f0f8ff')
        action_frame.pack(fill='x', pady=5)
        
        cols = DEFAULT_COLUMNS.copy()
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=PAGE_SIZE)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100 if c != 'Name' else 180)
        tree.pack(fill='both', expand=True, pady=8)
        tree.bind('<Double-1>', lambda e: self._open_detail_from_tree(tree))

        self.current_page = {'i': 0}
        nav = tk.Frame(frame, bg='#f0f8ff')
        nav.pack(fill='x')
        page_label = tk.Label(nav, text="Page 1", bg='#f0f8ff')

        def refresh_view():
            self.search_var.set("")
            self.gender_filter_var.set("All")
            self.sort_by_var.set("SerialNo")
            self._view_refresh(tree, page_label)
            
        tk.Button(action_frame, text="Filter", command=lambda: self._view_refresh(tree, page_label)).pack(side='left', padx=6)
        tk.Button(action_frame, text="Refresh", command=refresh_view, bg='#17a2b8', fg='white').pack(side='left', padx=6)
        tk.Button(action_frame, text="Back", command=self.show_main_menu, bg='#6c757d', fg='white').pack(side='right')

        gender_combo.bind("<<ComboboxSelected>>", lambda e: self._view_refresh(tree, page_label))
        sort_combo.bind("<<ComboboxSelected>>", lambda e: self._view_refresh(tree, page_label))
        search_entry.bind("<Return>", lambda e: self._view_refresh(tree, page_label))

        def prev_page():
            if self.current_page['i'] > 0:
                self.current_page['i'] -= 1
                self._view_refresh(tree=tree, page_label=page_label)

        def next_page():
            df = self._get_filtered_df()
            total_pages = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
            if self.current_page['i'] < total_pages - 1:
                self.current_page['i'] += 1
                self._view_refresh(tree=tree, page_label=page_label)

        tk.Button(nav, text="<< Prev", command=prev_page).pack(side='left', padx=4)
        tk.Button(nav, text="Next >>", command=next_page).pack(side='left', padx=4)
        page_label.pack(side='left', padx=6)
        self._view_refresh(tree=tree, page_label=page_label)

    def _get_filtered_df(self):
        df = self.patients_df.copy()
        s = (self.search_var.get() or "").strip().lower()
        if s: df = df[df.apply(lambda r: any(s in str(cell).lower() for cell in r), axis=1)]
        gender_filter = self.gender_filter_var.get()
        if gender_filter != "All": df = df[df['Gender'] == gender_filter]
        sort_by = self.sort_by_var.get()
        if sort_by in df.columns:
            try:
                if sort_by in ['SerialNo', 'Age']: df[sort_by] = pd.to_numeric(df[sort_by], errors='coerce')
                df = df.sort_values(by=sort_by, na_position='last')
            except Exception: pass
        return df.reset_index(drop=True)

    def _view_refresh(self, tree, page_label):
        if tree is None: return
        for it in tree.get_children(): tree.delete(it)
        df = self._get_filtered_df()
        total_pages = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        if self.current_page['i'] >= total_pages: self.current_page['i'] = 0
        start = self.current_page['i'] * PAGE_SIZE
        for _, r in df.iloc[start:start + PAGE_SIZE].iterrows():
            vals = [r.get(c, '') for c in DEFAULT_COLUMNS]
            tree.insert('', 'end', values=vals)
        if page_label is not None: page_label.config(text=f"Page {self.current_page['i'] + 1} / {total_pages} (Total: {len(df)})")

    def _open_detail_from_tree(self, tree):
        sel = tree.selection()
        if not sel: return
        values = tree.item(sel[0])['values']
        patient = dict(zip(DEFAULT_COLUMNS, values))
        self._show_patient_detail(patient)

    def _show_patient_detail(self, patient):
        win = tk.Toplevel(self)
        win.title(f"Details - {patient.get('Name', '')}")
        win.geometry("480x640")
        top = tk.Frame(win, bg='#f0f8ff')
        top.pack(fill='both', expand=True, padx=10, pady=10)

        pp = patient.get('PhotoPath', '')
        if pp and os.path.exists(str(pp)):
            try:
                img = Image.open(pp)
                img.thumbnail((200, 200))
                photo = ImageTk.PhotoImage(img)
                lbl = tk.Label(top, image=photo, bg='#f0f8ff')
                lbl.image = photo
                lbl.pack(pady=8)
            except Exception:
                tk.Label(top, text="(Photo cannot be shown)", bg='#f0f8ff').pack()
        else:
            tk.Label(top, text="(No Photo)", bg='#f0f8ff').pack(pady=8)

        info = tk.Frame(top, bg='#f0f8ff')
        info.pack(fill='both', expand=True)
        for c in DEFAULT_COLUMNS:
            v = patient.get(c, '')
            if pd.isna(v): v = ''
            tk.Label(info, text=f"{c}: {v}", anchor='w', justify='left', wraplength=420, bg='#f0f8ff').pack(fill='x', padx=6, pady=2)

        btnf = tk.Frame(win, bg='#f0f8ff')
        btnf.pack(fill='x', pady=8)
        tk.Button(btnf, text="Export as PDF", bg='#28a745', fg='white', command=lambda: self._export_pdf_for_patient(patient)).pack(side='left', padx=6)
        tk.Button(btnf, text="Open Photo", bg='#4c72b0', fg='white', command=lambda: self._open_photo(patient)).pack(side='left', padx=6)
        tk.Button(btnf, text="Close", bg='#6c757d', fg='white', command=win.destroy).pack(side='right', padx=6)

    def _export_pdf_for_patient(self, patient, path=None):
        if path is None:
            fname = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"patient_{patient.get('SerialNo','')}.pdf")
            if not fname: return
        else:
            fname = path

        try:
            doc = SimpleDocTemplate(fname, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            story.append(Paragraph("<b>Patient Report</b>", styles['Title']))
            story.append(Spacer(1, 12))
            img_path = patient.get('PhotoPath', '')
            if img_path and os.path.exists(str(img_path)):
                story.append(ReportImage(img_path, width=200, height=200))
                story.append(Spacer(1, 12))
            for col in DEFAULT_COLUMNS:
                val = str(patient.get(col, ''))
                story.append(Paragraph(f"<b>{col}:</b> {val}", styles['Normal']))
            doc.build(story)
            if path is None: messagebox.showinfo("Export", f"Saved PDF: {fname}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export PDF: {e}")

    def _open_photo(self, patient):
        p = patient.get('PhotoPath', '')
        if not p or not os.path.exists(str(p)):
            messagebox.showinfo("No Photo", "No photo for this patient")
            return
        try:
            if os.name == 'nt': os.startfile(p)
            else: subprocess.run(['xdg-open', p], check=False)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open photo: {e}")

    def show_edit_patients(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.current_frame = frame
        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 8))
        tk.Label(header, text="Edit Patients", font=("Arial", 16, "bold"), fg='white', bg='#2c5aa0').pack(pady=8)
        filter_frame = tk.Frame(frame, bg='#f0f8ff')
        filter_frame.pack(fill='x', pady=6)
        tk.Label(filter_frame, text="Search:", bg='#f0f8ff').pack(side='left', padx=6)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left')
        tk.Label(filter_frame, text="Gender Filter:", bg='#f0f8ff').pack(side='left', padx=(20, 6))
        self.gender_filter_var = tk.StringVar(value="All")
        gender_options = ["All"] + list(self.patients_df['Gender'].dropna().unique())
        gender_combo = ttk.Combobox(filter_frame, textvariable=self.gender_filter_var, values=gender_options, state='readonly')
        gender_combo.pack(side='left')
        tk.Label(filter_frame, text="Sort By:", bg='#f0f8ff').pack(side='left', padx=(20, 6))
        self.sort_by_var = tk.StringVar(value="SerialNo")
        sort_options = ['SerialNo', 'Name', 'Age', 'StartDate', 'EndDate']
        sort_combo = ttk.Combobox(filter_frame, textvariable=self.sort_by_var, values=sort_options, state='readonly')
        sort_combo.pack(side='left')
        action_frame = tk.Frame(frame, bg='#f0f8ff')
        action_frame.pack(fill='x', pady=5)
        cols = DEFAULT_COLUMNS.copy()
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=PAGE_SIZE)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100 if c != 'Name' else 160)
        tree.pack(fill='both', expand=True, pady=8)
        tree.bind('<Double-1>', lambda e: self._edit_open_window(tree))
        self.current_page = {'i': 0}
        nav = tk.Frame(frame, bg='#f0f8ff')
        nav.pack(fill='x')
        page_label = tk.Label(frame, text="Page 1", bg='#f0f8ff')
        def refresh_edit():
            self.search_var.set("")
            self.gender_filter_var.set("All")
            self.sort_by_var.set("SerialNo")
            self._edit_refresh(tree, page_label)
        tk.Button(action_frame, text="Filter", command=lambda: self._edit_refresh(tree, page_label)).pack(side='left', padx=6)
        tk.Button(action_frame, text="Refresh", command=refresh_edit, bg='#17a2b8', fg='white').pack(side='left', padx=6)
        tk.Button(action_frame, text="Back", command=self.show_main_menu, bg='#6c757d', fg='white').pack(side='right')
        gender_combo.bind("<<ComboboxSelected>>", lambda e: self._edit_refresh(tree, page_label))
        sort_combo.bind("<<ComboboxSelected>>", lambda e: self._edit_refresh(tree, page_label))
        search_entry.bind("<Return>", lambda e: self._edit_refresh(tree, page_label))
        tk.Button(nav, text="Prev", command=lambda: self._edit_page_change(-1, tree, page_label)).pack(side='left', padx=6)
        tk.Button(nav, text="Next", command=lambda: self._edit_page_change(1, tree, page_label)).pack(side='left', padx=6)
        page_label.pack(side='left', padx=6)
        self._edit_refresh(tree=tree, page_label=page_label)
        
    def _edit_get_filtered_df(self):
        df = self.patients_df.copy()
        s = (self.search_var.get() or "").strip().lower()
        if s: df = df[df.apply(lambda r: any(s in str(cell).lower() for cell in r), axis=1)]
        gender_filter = self.gender_filter_var.get()
        if gender_filter != "All": df = df[df['Gender'] == gender_filter]
        sort_by = self.sort_by_var.get()
        if sort_by in df.columns:
            try:
                if sort_by in ['SerialNo', 'Age']: df[sort_by] = pd.to_numeric(df[sort_by], errors='coerce')
                df = df.sort_values(by=sort_by, na_position='last')
            except Exception: pass
        return df.reset_index(drop=True)

    def _edit_refresh(self, tree, page_label):
        if tree is None: return
        for it in tree.get_children(): tree.delete(it)
        df = self._edit_get_filtered_df()
        total = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        if self.current_page['i'] >= total: self.current_page['i'] = 0
        start = self.current_page['i'] * PAGE_SIZE
        for _, r in df.iloc[start:start + PAGE_SIZE].iterrows():
            vals = [r.get(c, '') for c in DEFAULT_COLUMNS]
            tree.insert('', 'end', values=vals)
        if page_label: page_label.config(text=f"Page {self.current_page['i'] + 1} / {total} (Total: {len(df)})")

    def _edit_page_change(self, delta, tree, page_label):
        df = self._edit_get_filtered_df()
        total = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        new = self.current_page['i'] + delta
        if 0 <= new < total:
            self.current_page['i'] = new
            self._edit_refresh(tree, page_label)

    def _edit_open_window(self, tree):
        sel = tree.selection()
        if not sel: return
        vals = tree.item(sel[0])['values']
        record = dict(zip(DEFAULT_COLUMNS, vals))
        original_serial = record.get('SerialNo')
        win = tk.Toplevel(self)
        win.title(f"Edit - {record.get('Name', '')}")
        win.geometry("650x680")
        canvas = tk.Canvas(win, bg='#f0f8ff', highlightthickness=0)
        scroll = ttk.Scrollbar(win, orient='vertical', command=canvas.yview)
        content = tk.Frame(canvas, bg='#f0f8ff')
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content, anchor='nw')
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')
        entries = {}
        for i, col in enumerate(['SerialNo', 'Name', 'Email', 'Gender', 'Age', 'Address', 'PhoneNo', 'Occupation', 'AadharNo', 'StartDate', 'EndDate', 'Satisfied']):
            row = tk.Frame(content, bg='#f0f8ff')
            row.pack(fill='x', pady=6, padx=8)
            tk.Label(row, text=f"{col}:", width=18, anchor='e', bg='#f0f8ff').pack(side='left', padx=(0, 8))
            val = record.get(col, '')
            var = tk.StringVar(value=str(val))
            ent = tk.Entry(row, textvariable=var, width=40)
            ent.pack(side='left', fill='x', expand=True)
            entries[col] = var
        row_photo = tk.Frame(content, bg='#f0f8ff'); row_photo.pack(fill='x', pady=6, padx=8)
        tk.Label(row_photo, text="Photo:", width=18, anchor='e', bg='#f0f8ff').pack(side='left', padx=(0, 8))
        tk.Button(row_photo, text="Add/Update Photo", bg='#4c72b0', fg='white', command=lambda: self._edit_add_photo(original_serial)).pack(side='left', padx=4)
        tk.Button(row_photo, text="Remove Photo", bg='#dc3545', fg='white', command=lambda: self._edit_remove_photo(original_serial)).pack(side='left', padx=4)
        row_symp = tk.Frame(content, bg='#f0f8ff'); row_symp.pack(fill='x', pady=6, padx=8)
        tk.Label(row_symp, text="Symptoms:", width=18, anchor='e', bg='#f0f8ff').pack(side='left', padx=(0, 8))
        sym = scrolledtext.ScrolledText(row_symp, height=4, width=46)
        sym.pack(side='left', fill='both', expand=True)
        sym.delete("1.0", "end"); sym.insert("1.0", record.get('Symptoms', ''))
        row_treat = tk.Frame(content, bg='#f0f8ff'); row_treat.pack(fill='x', pady=6, padx=8)
        tk.Label(row_treat, text="Treatment:", width=18, anchor='e', bg='#f0f8ff').pack(side='left', padx=(0, 8))
        treat = scrolledtext.ScrolledText(row_treat, height=4, width=46)
        treat.pack(side='left', fill='both', expand=True)
        treat.delete("1.0", "end"); treat.insert("1.0", record.get('Treatment', ''))
        def save_changes():
            try:
                idx_list = self.patients_df[self.patients_df['SerialNo'] == str(original_serial)].index
                if not idx_list.any():
                    messagebox.showerror("Error", "Could not find patient to update.")
                    return
                idx = idx_list[0]
                new_serial = entries['SerialNo'].get()
                if new_serial != str(original_serial) and new_serial in self.patients_df['SerialNo'].values:
                    messagebox.showerror("Error", f"Serial number {new_serial} already exists.")
                    return
                for col, var in entries.items(): self.patients_df.at[idx, col] = var.get()
                self.patients_df.at[idx, 'Symptoms'] = sym.get("1.0", "end").strip()
                self.patients_df.at[idx, 'Treatment'] = treat.get("1.0", "end").strip()
                self.save_patients()
                messagebox.showinfo("Saved", "Patient updated")
                win.destroy()
                self.show_edit_patients()
            except Exception as e: messagebox.showerror("Error", f"Could not save: {e}")
        btnf = tk.Frame(content, bg='#f0f8ff'); btnf.pack(pady=8)
        tk.Button(btnf, text="Save Changes", bg='#28a745', fg='white', command=save_changes).pack()

    def _edit_add_photo(self, serial):
        path = filedialog.askopenfilename(title="Select Photo", filetypes=[("Image", "*.jpg *.jpeg *.png *.bmp")])
        if not path: return
        os.makedirs(PHOTO_DIR, exist_ok=True)
        fname = f"patient_{serial}.png"
        dest = os.path.join(PHOTO_DIR, fname)
        try:
            img = Image.open(path)
            img.save(dest, format='PNG')
            idx = self.patients_df[self.patients_df['SerialNo'] == str(serial)].index[0]
            self.patients_df.at[idx, 'PhotoPath'] = dest
            self.save_patients()
            messagebox.showinfo("Photo", "Photo updated successfully.")
        except Exception as e: messagebox.showerror("Error", f"Could not update photo: {e}")

    def _edit_remove_photo(self, serial):
        idx = self.patients_df[self.patients_df['SerialNo'] == str(serial)].index[0]
        path = self.patients_df.at[idx, 'PhotoPath']
        if not path or not os.path.exists(str(path)):
            messagebox.showinfo("Info", "No photo to remove.")
            return
        if messagebox.askyesno("Confirm", "Are you sure you want to remove this patient's photo?"):
            try:
                os.remove(path)
                self.patients_df.at[idx, 'PhotoPath'] = ''
                self.save_patients()
                messagebox.showinfo("Photo", "Photo removed successfully.")
            except Exception as e: messagebox.showerror("Error", f"Could not remove photo: {e}")

    def show_delete_patients(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.current_frame = frame
        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 8))
        tk.Label(header, text="Delete Patients", font=("Arial", 16, "bold"), fg='white', bg='#2c5aa0').pack(pady=8)
        filter_frame = tk.Frame(frame, bg='#f0f8ff')
        filter_frame.pack(fill='x', pady=6)
        tk.Label(filter_frame, text="Search:", bg='#f0f8ff').pack(side='left', padx=6)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left')
        action_frame = tk.Frame(frame, bg='#f0f8ff')
        action_frame.pack(fill='x', pady=5)
        
        cols = ['Select'] + DEFAULT_COLUMNS
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=PAGE_SIZE, selectmode='extended')
        tree.heading('Select', text='✔')
        tree.column('Select', width=50, anchor='center')
        for c in DEFAULT_COLUMNS:
            tree.heading(c, text=c)
            tree.column(c, width=100 if c != 'Name' else 160)
        tree.pack(fill='both', expand=True, pady=8)
        
        page_label = tk.Label(frame, text="Page 1", bg='#f0f8ff')

        def refresh_delete():
            self.search_var.set("")
            self._delete_refresh(tree, page_label)
            
        def _delete_toggle_all():
            all_selected = all(self.checkbox_state.values()) if self.checkbox_state else False
            new_state = not all_selected
            for iid in self.checkbox_state.keys():
                self.checkbox_state[iid] = new_state
                tree.set(iid, 'Select', '☑' if new_state else '☐')

        tk.Button(action_frame, text="Filter", command=lambda: self._delete_refresh(tree, page_label)).pack(side='left', padx=6)
        tk.Button(action_frame, text="Refresh", command=refresh_delete, bg='#17a2b8', fg='white').pack(side='left', padx=6)
        tk.Button(action_frame, text="Toggle All Selections", command=_delete_toggle_all, bg='#007bff', fg='white').pack(side='left', padx=6)
        tk.Button(action_frame, text="Back", command=self.show_main_menu, bg='#6c757d', fg='white').pack(side='right')
        search_entry.bind("<Return>", lambda e: self._delete_refresh(tree, page_label))
        
        self.checkbox_state = {}
        def on_click(event):
            region = tree.identify_region(event.x, event.y)
            if region != "cell": return
            col = tree.identify_column(event.x)
            row = tree.identify_row(event.y)
            if not row: return
            if col == '#1':
                current = self.checkbox_state.get(row, False)
                self.checkbox_state[row] = not current
                tree.set(row, 'Select', '☑' if self.checkbox_state[row] else '☐')
        tree.bind('<Button-1>', on_click)
        
        self.current_page = {'i': 0}
        nav = tk.Frame(frame, bg='#f0f8ff')
        nav.pack(fill='x')
        tk.Button(nav, text="Prev", command=lambda: self._delete_page_change(-1, tree, page_label)).pack(side='left', padx=6)
        tk.Button(nav, text="Next", command=lambda: self._delete_page_change(1, tree, page_label)).pack(side='left', padx=6)
        tk.Button(nav, text="Delete Selected", bg='#dc3545', fg='white', command=lambda: self._delete_checked(tree)).pack(side='left', padx=8)
        page_label.pack(side='left', padx=6)
        self._delete_refresh(tree=tree, page_label=page_label)
        
    def _delete_get_filtered_df(self):
        df = self.patients_df.copy()
        s = (self.search_var.get() or "").strip().lower()
        if s: df = df[df.apply(lambda r: any(s in str(cell).lower() for cell in r), axis=1)]
        return df.reset_index(drop=True)

    def _delete_refresh(self, tree, page_label):
        if tree is None: return
        for it in tree.get_children(): tree.delete(it)
        self.checkbox_state.clear()
        df = self._delete_get_filtered_df()
        total = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        if self.current_page['i'] >= total: self.current_page['i'] = 0
        start = self.current_page['i'] * PAGE_SIZE
        for _, r in df.iloc[start:start + PAGE_SIZE].iterrows():
            vals = ['☐'] + [r.get(c, '') for c in DEFAULT_COLUMNS]
            iid = tree.insert('', 'end', values=vals)
            self.checkbox_state[iid] = False
        if page_label: page_label.config(text=f"Page {self.current_page['i'] + 1} / {total} (Total: {len(df)})")
            
    def _delete_page_change(self, delta, tree, page_label):
        df = self._delete_get_filtered_df()
        total = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        new = self.current_page['i'] + delta
        if 0 <= new < total:
            self.current_page['i'] = new
            self._delete_refresh(tree, page_label)

    def _delete_checked(self, tree):
        serials_to_delete = [str(tree.item(iid)['values'][1]) for iid, checked in self.checkbox_state.items() if checked]
        if not serials_to_delete:
            messagebox.showwarning("Warning", "No checked records to delete")
            return
        if not messagebox.askyesno("Confirm", f"Delete {len(serials_to_delete)} checked records?"): return
        self.patients_df = self.patients_df[~self.patients_df['SerialNo'].isin(serials_to_delete)].reset_index(drop=True)
        self.save_patients()
        messagebox.showinfo("Deleted", f"Deleted {len(serials_to_delete)} records")
        self.show_delete_patients()

    def show_share_details(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.current_frame = frame

        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 8))
        tk.Label(header, text="Share / Export Patient Details", font=("Arial", 16, "bold"), fg='white', bg='#2c5aa0').pack(pady=8)
        
        toolbar = tk.Frame(frame, bg='#f0f8ff')
        toolbar.pack(fill='x', pady=6)
        tk.Label(toolbar, text="Search:", bg='#f0f8ff').pack(side='left', padx=6)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar, textvariable=self.search_var, width=30)
        search_entry.pack(side='left')
        
        cols = ['Select'] + ['SerialNo', 'Name', 'Email', 'PhoneNo']
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=PAGE_SIZE, selectmode='extended')
        tree.heading('Select', text='✔')
        tree.column('Select', width=50, anchor='center')
        for c in ['SerialNo', 'Name', 'Email', 'PhoneNo']:
            tree.heading(c, text=c)
            tree.column(c, width=180)
        tree.pack(fill='both', expand=True, pady=8)

        page_label = tk.Label(frame, text="Page 1", bg='#f0f8ff')
        
        def refresh_share():
            self.search_var.set("")
            self._share_refresh(tree, page_label)
            
        tk.Button(toolbar, text="Filter", command=lambda: self._share_refresh(tree, page_label)).pack(side='left', padx=6)
        tk.Button(toolbar, text="Refresh", bg='#17a2b8', fg='white', command=refresh_share).pack(side='left', padx=6)
        tk.Button(toolbar, text="Back", command=self.show_main_menu, bg='#6c757d', fg='white').pack(side='right')

        search_entry.bind("<Return>", lambda e: self._share_refresh(tree, page_label))

        self.checkbox_state = {}
        def on_click(event):
            region = tree.identify_region(event.x, event.y)
            if region != "cell": return
            col = tree.identify_column(event.x)
            row = tree.identify_row(event.y)
            if not row: return
            if col == '#1':
                current = self.checkbox_state.get(row, False)
                self.checkbox_state[row] = not current
                tree.set(row, 'Select', '☑' if self.checkbox_state[row] else '☐')
        tree.bind('<Button-1>', on_click)
        
        self.current_page = {'i': 0}
        nav = tk.Frame(frame, bg='#f0f8ff')
        nav.pack(fill='x')
        tk.Button(nav, text="Prev", command=lambda: self._share_page_change(-1, tree, page_label)).pack(side='left', padx=6)
        tk.Button(nav, text="Next", command=lambda: self._share_page_change(1, tree, page_label)).pack(side='left', padx=6)
        page_label.pack(side='left', padx=6)
        
        btns = tk.Frame(frame, bg='#f0f8ff')
        btns.pack(fill='x', pady=6)

        def _share_toggle_all():
            all_selected = all(self.checkbox_state.values()) if self.checkbox_state else False
            new_state = not all_selected
            for iid in self.checkbox_state.keys():
                self.checkbox_state[iid] = new_state
                tree.set(iid, 'Select', '☑' if new_state else '☐')

        tk.Button(btns, text="Toggle All Selections", bg='#007bff', fg='white', command=_share_toggle_all).pack(side='left', padx=6)
        tk.Button(btns, text="Export Selected to Excel", bg='#28a745', fg='white', command=lambda: self._share_export_excel(tree)).pack(side='left', padx=6)
        tk.Button(btns, text="Export Selected as PDF", bg='#28a745', fg='white', command=lambda: self._share_export_pdf(tree)).pack(side='left', padx=6)
        tk.Button(btns, text="Show Statistics", bg='#17a2b8', fg='white', command=lambda: self._share_show_stats(tree)).pack(side='left', padx=6)
        
        self._share_refresh(tree=tree, page_label=page_label)
    
    def _share_get_filtered_df(self):
        df = self.patients_df.copy()
        s = (self.search_var.get() or "").strip().lower()
        if s:
            search_cols = ['SerialNo', 'Name', 'Email', 'PhoneNo']
            df = df[df.apply(lambda r: any(s in str(r.get(c, '')).lower() for c in search_cols), axis=1)]
        return df.reset_index(drop=True)

    def _share_refresh(self, tree, page_label):
        if tree is None: return
        for it in tree.get_children(): tree.delete(it)
        self.checkbox_state.clear()
        df = self._share_get_filtered_df()
        
        total = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        if self.current_page['i'] >= total: self.current_page['i'] = 0

        start = self.current_page['i'] * PAGE_SIZE
        for _, r in df.iloc[start:start + PAGE_SIZE].iterrows():
            vals = ['☐'] + [r.get(c, '') for c in ['SerialNo', 'Name', 'Email', 'PhoneNo']]
            iid = tree.insert('', 'end', values=vals)
            self.checkbox_state[iid] = False
        if page_label: page_label.config(text=f"Page {self.current_page['i'] + 1} / {total} (Total: {len(df)})")
            
    def _share_page_change(self, delta, tree, page_label):
        df = self._share_get_filtered_df()
        total = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
        new = self.current_page['i'] + delta
        if 0 <= new < total:
            self.current_page['i'] = new
            self._share_refresh(tree, page_label)

    def _share_get_selected_records(self, tree):
        serials = [str(tree.item(iid)['values'][1]) for iid, checked in self.checkbox_state.items() if checked]
        if not serials:
            messagebox.showwarning("Warning", "No records selected.")
            return pd.DataFrame()
        return self.patients_df[self.patients_df['SerialNo'].isin(serials)]

    def _share_export_excel(self, tree):
        df = self._share_get_selected_records(tree)
        if df.empty: return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path: return
        try:
            df.to_excel(path, index=False)
            messagebox.showinfo("Exported", f"Exported {len(df)} record(s) to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export: {e}")
            
    def _share_export_pdf(self, tree):
        df = self._share_get_selected_records(tree)
        if df.empty: return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not path: return
        try:
            doc = SimpleDocTemplate(path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            for i, patient in df.iterrows():
                if story: story.append(PageBreak())
                story.append(Paragraph(f"<b>Patient Report: {patient.get('Name')}</b>", styles['Title']))
                story.append(Spacer(1, 12))
                img_path = patient.get('PhotoPath', '')
                if img_path and os.path.exists(str(img_path)):
                    story.append(ReportImage(img_path, width=150, height=150))
                    story.append(Spacer(1, 12))
                for col in DEFAULT_COLUMNS:
                    val = str(patient.get(col, ''))
                    story.append(Paragraph(f"<b>{col}:</b> {val}", styles['Normal']))
            doc.build(story)
            messagebox.showinfo("Export", f"Saved PDF: {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not export PDF: {e}")

    def _share_show_stats(self, tree):
        df = self._share_get_selected_records(tree)
        if df.empty: return
        win = tk.Toplevel(self)
        win.title("Selected Patients Statistics")
        win.geometry("500x300")
        win.configure(bg='#f0f8ff')
        total = len(df)
        df['Age'] = pd.to_numeric(df['Age'], errors='coerce')
        males, females = int((df["Gender"] == "Male").sum()), int((df["Gender"] == "Female").sum())
        others, avg_age = total - males - females, df['Age'].mean()
        tk.Label(win, text=f"Total Patients Selected: {total}", font=("Arial", 12, "bold"), bg='#f0f8ff').pack(pady=4)
        tk.Label(win, text=f"Male: {males} ({males / total * 100:.1f}%)", font=("Arial", 10), bg='#f0f8ff').pack()
        tk.Label(win, text=f"Female: {females} ({females / total * 100:.1f}%)", font=("Arial", 10), bg='#f0f8ff').pack()
        tk.Label(win, text=f"Other: {others} ({others / total * 100:.1f}%)", font=("Arial", 10), bg='#f0f8ff').pack()
        tk.Label(win, text=f"Average Age: {avg_age:.1f}", font=("Arial", 10), bg='#f0f8ff').pack(pady=4)
        tk.Label(win, text=f"Most Common Symptom: {get_most_common(df, 'Symptoms')}", font=("Arial", 10), bg='#f0f8ff').pack()
        tk.Label(win, text=f"Average Treatment Duration: {get_average_duration(df)}", font=("Arial", 10), bg='#f0f8ff').pack(pady=4)

    def show_manage_patient(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.current_frame = frame
        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 8))
        tk.Label(header, text="Manage Patients", font=("Arial", 16, "bold"), fg='white', bg='#2c5aa0').pack(pady=8)
        btns_frame = tk.Frame(frame, bg='#f0f8ff')
        btns_frame.pack(pady=20)
        buttons = [("Backup Database", self._backup_db, '#4c72b0', 'white'), ("Import Database", self._import_db, '#ffc107', 'black'),
                   ("Export All to Excel", self._export_all, '#28a745', 'white'), ("Show Overall Stats", self._show_stats, '#17a2b8', 'white')]
        for i, (text, cmd, bg_color, fg_color) in enumerate(buttons):
            btn = tk.Button(btns_frame, text=text, width=22, height=2, font=("Arial", 11), command=cmd, bg=bg_color, fg=fg_color, relief='raised', bd=2)
            btn.grid(row=i // 2, column=i % 2, padx=10, pady=8)
        tk.Button(frame, text="Back to Main Menu", width=48, height=2, bg='#6c757d', fg='white', command=self.show_main_menu).pack(pady=20)

    def _backup_db(self):
        if not os.path.exists(DATA_FILE):
            messagebox.showwarning("Warning", "No data file to backup")
            return
        dst = f"patients_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        try:
            shutil.copy2(DATA_FILE, dst)
            messagebox.showinfo("Backup", f"Backup created: {dst}")
        except Exception as e: messagebox.showerror("Error", f"Could not backup: {e}")

    def _import_db(self):
        path = filedialog.askopenfilename(title="Select file to import", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        if not messagebox.askyesno("Confirm Import", "This will add records from the selected file and assign them new serial numbers. Are you sure?"): return
        try:
            numeric_serials = pd.to_numeric(self.patients_df['SerialNo'], errors='coerce').dropna()
            max_existing_serial = int(numeric_serials.max()) if not numeric_serials.empty else 0
            
            new_df = pd.read_excel(path)
            for col in DEFAULT_COLUMNS:
                if col not in new_df.columns: new_df[col] = ''
            new_df = new_df[DEFAULT_COLUMNS]

            new_df['SerialNo'] = range(max_existing_serial + 1, max_existing_serial + 1 + len(new_df))

            self.patients_df = pd.concat([self.patients_df, new_df], ignore_index=True)
            self.save_patients()
            messagebox.showinfo("Import", f"Database imported successfully. {len(new_df)} new records were added with updated serial numbers.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not import: {e}")

    def _export_all(self):
        if self.patients_df.empty:
            messagebox.showinfo("No Data", "No patient records to export")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="patients_export.xlsx")
        if not path: return
        try:
            self.patients_df.to_excel(path, index=False)
            messagebox.showinfo("Exported", f"Exported all records to {path}")
        except Exception as e: messagebox.showerror("Error", f"Could not export: {e}")

    def _show_stats(self):
        win = tk.Toplevel(self)
        win.title("Overall Patient Statistics")
        win.geometry("500x300")
        win.configure(bg='#f0f8ff')
        df = self.patients_df.copy()
        if df.empty:
            tk.Label(win, text="No patient data available for statistics.", font=("Arial", 12), bg='#f0f8ff').pack(pady=20)
            return
        total = len(df)
        df['Age'] = pd.to_numeric(df['Age'], errors='coerce')
        males, females = int((df["Gender"] == "Male").sum()), int((df["Gender"] == "Female").sum())
        others, avg_age = total - males - females, df['Age'].mean()
        content_frame = tk.Frame(win, bg='#f0f8ff', padx=20, pady=20)
        content_frame.pack(expand=True, fill='both')
        stats = [("Total Patients:", total), ("Male:", f"{males} ({males/total*100:.1f}%)" if total > 0 else 0),
                 ("Female:", f"{females} ({females/total*100:.1f}%)" if total > 0 else 0), ("Other:", f"{others} ({others/total*100:.1f}%)" if total > 0 else 0),
                 ("Average Age:", f"{avg_age:.1f}" if pd.notna(avg_age) else "N/A"),
                 ("Most Common Symptom:", get_most_common(df, 'Symptoms')), ("Average Treatment Duration:", get_average_duration(df))]
        for i, (label, value) in enumerate(stats):
            tk.Label(content_frame, text=label, font=("Arial", 11, "bold"), bg='#f0f8ff', anchor='w').grid(row=i, column=0, sticky='w', pady=4)
            tk.Label(content_frame, text=value, font=("Arial", 11), bg='#f0f8ff', anchor='w').grid(row=i, column=1, sticky='w', padx=10)

    def show_duplicate_page(self):
        self.clear_frame()
        frame = tk.Frame(self, bg='#f0f8ff')
        frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.current_frame = frame
        
        header = tk.Frame(frame, bg='#2c5aa0')
        header.pack(fill='x', pady=(0, 8))
        tk.Label(header, text="Manage Duplicates", font=("Arial", 16, "bold"), fg='white', bg='#2c5aa0').pack(pady=8)
        
        # **FIX:** Using a more practical set of columns to identify duplicates.
        # Checking against all columns is too strict and will miss most real duplicates.
        dup_cols = ['Name', 'Email', 'PhoneNo', 'AadharNo']
        
        # Drop records where all key identifiers are missing before checking for duplicates
        df_for_check = self.patients_df.dropna(subset=dup_cols, how='all')
        
        df_dups = df_for_check[df_for_check.duplicated(subset=dup_cols, keep=False)].sort_values(by=['Name', 'SerialNo'])

        if df_dups.empty:
            tk.Label(frame, text="No duplicate patients found.", bg='#f0f8ff', font=("Arial", 12)).pack(pady=20)
            tk.Button(frame, text="Back", command=self.show_main_menu, bg='#6c757d', fg='white').pack()
            return

        cols = ['Select'] + DEFAULT_COLUMNS
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=PAGE_SIZE)
        tree.heading('Select', text='✔')
        tree.column('Select', width=50, anchor='center')
        for c in DEFAULT_COLUMNS:
            tree.heading(c, text=c)
            tree.column(c, width=100)
        tree.pack(fill='both', expand=True, pady=8)
        
        self.checkbox_state = {}
        for _, r in df_dups.iterrows():
            vals = ['☐'] + [r.get(c, '') for c in DEFAULT_COLUMNS]
            iid = tree.insert('', 'end', values=vals)
            self.checkbox_state[iid] = False

        def on_click(event):
            region = tree.identify_region(event.x, event.y)
            if region != "cell": return
            col = tree.identify_column(event.x)
            row = tree.identify_row(event.y)
            if not row or col != '#1': return
            current = self.checkbox_state.get(row, False)
            self.checkbox_state[row] = not current
            tree.set(row, 'Select', '☑' if self.checkbox_state[row] else '☐')
        tree.bind('<Button-1>', on_click)

        def _edit_duplicate_serial(event):
            row_id = tree.identify_row(event.y)
            column_id = tree.identify_column(event.x)
            if not row_id or column_id != '#2': return

            original_serial = tree.item(row_id)['values'][1]
            new_serial = simpledialog.askstring("Edit Serial No", f"Enter new Serial No:", initialvalue=original_serial)
            
            if not new_serial or new_serial == original_serial: return
            if not new_serial.isdigit():
                messagebox.showerror("Error", "Serial number must be a numeric value.")
                return
            if new_serial in self.patients_df['SerialNo'].values:
                messagebox.showerror("Error", f"Serial number {new_serial} already exists.")
                return

            idx = self.patients_df[self.patients_df['SerialNo'] == str(original_serial)].index[0]
            self.patients_df.at[idx, 'SerialNo'] = new_serial
            self.save_patients()
            messagebox.showinfo("Success", "Serial number updated.")
            self.show_duplicate_page()
        tree.bind('<Double-1>', _edit_duplicate_serial)

        def get_selected_serials():
            return [str(tree.item(iid)['values'][1]) for iid, checked in self.checkbox_state.items() if checked]

        def delete_selected_duplicates():
            serials_to_delete = get_selected_serials()
            if not serials_to_delete:
                messagebox.showwarning("Warning", "Select records to delete.")
                return
            if not messagebox.askyesno("Confirm Delete", f"Delete {len(serials_to_delete)} selected records?"): return
            self.patients_df = self.patients_df[~self.patients_df['SerialNo'].isin(serials_to_delete)].reset_index(drop=True)
            self.save_patients()
            messagebox.showinfo("Success", "Selected duplicates deleted.")
            self.show_duplicate_page()

        def mark_as_not_duplicate():
            iids_to_remove = [iid for iid, checked in self.checkbox_state.items() if checked]
            if not iids_to_remove:
                messagebox.showwarning("Action", "Select records to mark as not duplicate.")
                return
            for iid in iids_to_remove:
                tree.delete(iid)
                del self.checkbox_state[iid]
            messagebox.showinfo("Action", f"{len(iids_to_remove)} record(s) hidden from this view.")

        btnf = tk.Frame(frame, bg='#f0f8ff')
        btnf.pack(pady=10)
        tk.Button(btnf, text="Delete Selected", bg='#dc3545', fg='white', command=delete_selected_duplicates).pack(side='left', padx=10)
        tk.Button(btnf, text="Mark as Not Duplicate (Hide)", bg='#ffc107', command=mark_as_not_duplicate).pack(side='left', padx=10)
        tk.Button(btnf, text="Back", command=self.show_main_menu, bg='#6c757d', fg='white').pack(side='left', padx=10)


if __name__ == "__main__":
    app = PatientManagementSystem()

    app.mainloop()
