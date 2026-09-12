import os
import json
import sqlite3
import shutil
import re
import time
import hashlib
import secrets
import logging
import tkinter as tk
import tkinter.font as tkFont
import customtkinter as ctk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox, Toplevel, simpledialog
import pandas as pd
import jdatetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
try:
    import docx
except ImportError:
    pass

# تنظیمات پایه
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

MONTHS = ['فروردین','اردیبهشت','خرداد','تیر','مرداد','شهریور','مهر','آبان','آذر','دی','بهمن','اسفند']

def to_persian_num(text):
    if pd.isna(text) or text is None: return ""
    return str(text).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))

def convert_persian_to_english_number(text):
    if not text: return text
    return str(text).translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789'))

def clean_text_advanced(val):
    if pd.isna(val) or val is None: return ""
    text = str(val).strip()
    chars_to_replace = {
        '۰':'0', '۱':'1', '۲':'2', '۳':'3', '۴':'4', '۵':'5', '۶':'6', '۷':'7', '۸':'8', '۹':'9',
        '٠':'0', '١':'1', '٢':'2', '٣':'3', '٤':'4', '٥':'5', '٦':'6', '٧':'7', '٨':'8', '٩':'9',
        'ي':'ی', 'ك':'ک', '\u200c':'', '\u200e':'', '\u200f':'', '\xa0':' '
    }
    for k, v in chars_to_replace.items(): text = text.replace(k, v)
    return text.strip()

def clean_contract_number(val):
    text = clean_text_advanced(val)
    for char in [' ', '/', '\\', '-', '_']: text = text.replace(char, '')
    return text

def clean_number(val):
    try:
        return int(re.sub(r'\D', '', clean_text_advanced(val)))
    except Exception:
        logging.debug(f"clean_number: مقدار قابل تبدیل به عدد نیست: {val!r}")
        return 0

def resolve_column(df, config_name, purpose=""):
    if not config_name: return None
    candidates = [c for c in df.columns if c == config_name or re.match(rf"^{re.escape(config_name)}\.\d+$", str(c))]
    if not candidates: return None
    if len(candidates) == 1: return candidates[0]

    best_col, max_score = candidates[0], -1
    for c in candidates:
        valid_vals = df[c].dropna().astype(str).str.strip()
        valid_vals = valid_vals[valid_vals != ""]
        if purpose == "contract":
            score = valid_vals.str.contains(r'\d+/\d+', regex=True).sum()
            if score == 0: score = valid_vals.str.contains(r'\d', regex=True).sum()
        else: score = len(valid_vals)
        if score > max_score: max_score, best_col = score, c
    return best_col

# ================= امنیت (هش رمز عبور مدیریت) =================
def hash_password(password, salt=None):
    """رمز عبور را با PBKDF2-HMAC-SHA256 هش می‌کند و (salt, hash) هگزادسیمال برمی‌گرداند."""
    if salt is None:
        salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), bytes.fromhex(salt), 100_000)
    return salt, digest.hex()

def verify_password(password, salt, expected_hash):
    if not salt or not expected_hash:
        return False
    _, digest = hash_password(password, salt)
    return secrets.compare_digest(digest, expected_hash)

class MammutInsuranceApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("سیستم جامع مالی بیمه ماموت")
        self.geometry("1450x900")

        self.base_dir = r"C:\Bime Ba Ma\برنامه مالی ماموت"
        self.config_path = os.path.join(self.base_dir, "config.json")
        self.db_dir = os.path.join(self.base_dir, "Database")
        self.archive_dir = os.path.join(self.base_dir, "Archive")
        self.template_dir = os.path.join(self.base_dir, "Data", "Template")
        self.backup_dir = os.path.join(self.base_dir, "Backups")
        self.payments_excel_path = os.path.join(self.archive_dir, "لیست_کل_پرداخت_ها.xlsx")

        self.main_font = ctk.CTkFont(family="Vazir", size=13)
        self.title_font = ctk.CTkFont(family="Vazir", size=16, weight="bold")
        self.big_font = ctk.CTkFont(family="Vazir", size=24, weight="bold")

        self.option_add("*TCombobox*Listbox.font", ("Vazir", 13))
        self.option_add("*TMenu*font", ("Vazir", 13))

        self.setup_directories()
        self.setup_logging()
        self.load_config()
        self.init_db()
        self.check_login()
        self.ensure_admin_password()
        self.setup_treeview_style()

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#0F172A")
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(7, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="بیمه ماموت", font=self.big_font, text_color="#38BDF8")
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 20))
        self.lbl_user = ctk.CTkLabel(self.sidebar_frame, text=f"کاربر: {self.current_user}", font=self.main_font, text_color="#94A3B8")
        self.lbl_user.grid(row=1, column=0, pady=(0, 20))

        menus = [("داشبورد و پرداخت‌ها", self.show_dashboard), ("پردازش هوشمند", self.show_processing),
                 ("تنظیمات سیستم", self.show_settings), ("تغییر کاربر", self.change_user)]
        for i, (txt, cmd) in enumerate(menus, 2):
            ctk.CTkButton(self.sidebar_frame, text=txt, font=self.main_font, command=cmd).grid(row=i, column=0, padx=20, pady=10)
        ctk.CTkButton(self.sidebar_frame, text="حذف کلی اطلاعات", font=self.main_font, fg_color="#E11D48", hover_color="#BE123C", command=self.reset_data).grid(row=6, column=0, padx=20, pady=10)

        self.main_frame = ctk.CTkFrame(self, corner_radius=15, fg_color="#1E293B")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.pending_preview_data = {}
        self.show_dashboard()

    def setup_directories(self):
        for path in [self.base_dir, self.db_dir, self.archive_dir, self.template_dir, self.backup_dir]: os.makedirs(path, exist_ok=True)

    def setup_logging(self):
        log_dir = os.path.join(self.base_dir, "Logs")
        os.makedirs(log_dir, exist_ok=True)
        logging.basicConfig(
            filename=os.path.join(log_dir, "app.log"),
            level=logging.INFO,
            format="%(asctime)s [%(levelname)s] %(message)s",
            encoding="utf-8"
        )
        logging.info("=== برنامه اجرا شد ===")

    def load_config(self):
        default_config = {
            "s_type": "نوع بیمه گذار", "s_company": "طرف قرارداد", "s_contract": "شماره قرارداد",
            "s_date": "تاریخ صدور", "s_premium": "حق بیمه", "s_name": "بیمه گذار", "s_policy": "شماره بیمه نامه",
            "b_type": "نوع بیمه گذار", "b_company": "طرف قرارداد", "b_contract": "شماره قرارداد",
            "b_date": "تاریخ صدور", "b_premium": "حق بیمه", "b_name": "بیمه گذار", "b_policy": "شماره بیمه نامه",
            "target_contracts": ["1404/30000/5051"], "cutoff_day": "25", "enforce_sequential_payment": "فعال",
            "admin_password_salt": "", "admin_password_hash": ""
        }
        if not os.path.exists(self.config_path):
            with open(self.config_path, 'w', encoding='utf-8') as f: json.dump(default_config, f, indent=4, ensure_ascii=False)
        with open(self.config_path, 'r', encoding='utf-8') as f:
            self.config = {k: str(v).strip() if isinstance(v, str) else v for k, v in json.load(f).items()}
            for k, v in default_config.items():
                if k not in self.config: self.config[k] = v

    def persist_config(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def init_db(self):
        self.conn = sqlite3.connect(os.path.join(self.db_dir, "insurance_db.sqlite"))
        cursor = self.conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS sys_user (id INTEGER PRIMARY KEY, username TEXT)")
        cursor.execute('''CREATE TABLE IF NOT EXISTS policies (
            id INTEGER PRIMARY KEY AUTOINCREMENT, policy_no TEXT UNIQUE, person_name TEXT,
            company_name TEXT, premium INTEGER, issue_date TEXT, period TEXT, type TEXT)''')

        cursor.execute('''CREATE TABLE IF NOT EXISTS installments (
            id INTEGER PRIMARY KEY AUTOINCREMENT, policy_id INTEGER, inst_num INTEGER, amount INTEGER,
            status TEXT DEFAULT 'پرداخت نشده', due_date TEXT, payment_type TEXT, description TEXT, receipt_path TEXT, payment_date TEXT,
            FOREIGN KEY(policy_id) REFERENCES policies(id))''')

        for col in ["due_date", "payment_type", "description", "receipt_path", "payment_date"]:
            try: cursor.execute(f"ALTER TABLE installments ADD COLUMN {col} TEXT")
            except sqlite3.OperationalError: pass  # ستون از قبل وجود دارد
        self.conn.commit()

    def check_login(self):
        cursor = self.conn.cursor()
        user = cursor.execute("SELECT username FROM sys_user ORDER BY id DESC LIMIT 1").fetchone()
        if user: self.current_user = user[0]
        else:
            self.current_user = simpledialog.askstring("ورود", "لطفا نام کاربری خود را وارد کنید:") or "مدیر_سیستم"
            cursor.execute("INSERT INTO sys_user (username) VALUES (?)", (self.current_user,))
            self.conn.commit()

    def change_user(self):
        if not self.verify_admin("برای تغییر کاربر، رمز مدیریت را وارد کنید:"):
            return
        new_u = simpledialog.askstring("تغییر کاربر", "نام کاربری جدید:")
        if new_u:
            self.current_user = new_u
            self.conn.cursor().execute("INSERT INTO sys_user (username) VALUES (?)", (new_u,))
            self.conn.commit()
            self.lbl_user.configure(text=f"کاربر: {self.current_user}")
            logging.info(f"کاربر به «{new_u}» تغییر کرد.")

    # ================= مدیریت رمز عبور مدیریت =================
    def ensure_admin_password(self):
        if self.config.get("admin_password_hash"):
            return
        while True:
            pw1 = simpledialog.askstring("تنظیم رمز مدیریت", "این اولین اجرای برنامه است.\nبرای انجام عملیات حساس (تغییر کاربر، حذف کلی اطلاعات) یک رمز مدیریت تعیین کنید:", show="*", parent=self)
            if not pw1:
                messagebox.showwarning("هشدار", "بدون تعیین رمز مدیریت، عملیات حساس تا زمان تنظیم رمز از بخش «تنظیمات سیستم» غیرفعال می‌ماند.")
                return
            pw2 = simpledialog.askstring("تایید رمز", "رمز عبور را دوباره وارد کنید:", show="*", parent=self)
            if pw1 == pw2:
                salt, h = hash_password(pw1)
                self.config["admin_password_salt"] = salt
                self.config["admin_password_hash"] = h
                self.persist_config()
                logging.info("رمز مدیریت برای اولین بار تنظیم شد.")
                messagebox.showinfo("موفق", "رمز مدیریت با موفقیت ثبت شد.")
                return
            messagebox.showerror("خطا", "رمزهای وارد شده مطابقت ندارند، دوباره تلاش کنید.")

    def verify_admin(self, prompt="رمز عبور مدیریت را وارد کنید:"):
        if not self.config.get("admin_password_hash"):
            messagebox.showerror("خطا", "رمز مدیریت هنوز تنظیم نشده است. ابتدا از «تنظیمات سیستم» یک رمز تعیین کنید.")
            return False
        pw = simpledialog.askstring("تایید هویت مدیر", prompt, show="*", parent=self)
        if pw is None:
            return False
        if verify_password(pw, self.config.get("admin_password_salt", ""), self.config.get("admin_password_hash", "")):
            return True
        logging.warning("تلاش ناموفق برای احراز هویت مدیر.")
        messagebox.showerror("خطا", "رمز اشتباه است.")
        return False

    def change_admin_password(self):
        if self.config.get("admin_password_hash") and not self.verify_admin("رمز فعلی مدیریت را وارد کنید:"):
            return
        pw1 = simpledialog.askstring("رمز جدید", "رمز جدید مدیریت را وارد کنید:", show="*", parent=self)
        if not pw1:
            return
        pw2 = simpledialog.askstring("تایید رمز جدید", "رمز جدید را دوباره وارد کنید:", show="*", parent=self)
        if pw1 != pw2:
            return messagebox.showerror("خطا", "رمزهای وارد شده مطابقت ندارند.")
        salt, h = hash_password(pw1)
        self.config["admin_password_salt"] = salt
        self.config["admin_password_hash"] = h
        self.persist_config()
        logging.info("رمز مدیریت تغییر کرد.")
        messagebox.showinfo("موفق", "رمز مدیریت با موفقیت تغییر کرد.")

    def reset_data(self):
        if not self.verify_admin("برای حذف کلی اطلاعات، رمز مدیریت را وارد کنید:"):
            return
        if not messagebox.askyesno("اخطار امنیتی", "حذف تمامی اطلاعات مالی و بایگانی‌ها؟\nپیش از حذف، یک نسخه پشتیبان کامل گرفته می‌شود."):
            return
        try:
            backup_path = self.create_backup()
            self.conn.cursor().execute("DELETE FROM installments")
            self.conn.cursor().execute("DELETE FROM policies")
            self.conn.commit()
            for item in os.listdir(self.archive_dir):
                p = os.path.join(self.archive_dir, item)
                os.remove(p) if os.path.isfile(p) else shutil.rmtree(p)
            logging.info(f"حذف کلی اطلاعات توسط «{self.current_user}» انجام شد. نسخه پشتیبان: {backup_path}")
            messagebox.showinfo("موفقیت", f"داده‌ها با موفقیت پاکسازی شدند.\nنسخه پشتیبان قبل از حذف در این مسیر ذخیره شد:\n{backup_path}")
            self.show_dashboard()
        except Exception as e:
            logging.exception("خطا در حذف کلی اطلاعات")
            messagebox.showerror("خطا", str(e))

    def create_backup(self):
        """قبل از عملیات مخرب، از دیتابیس و بایگانی یک فایل zip پشتیبان می‌سازد."""
        ts = jdatetime.date.today().strftime('%Y%m%d') + "_" + time.strftime("%H%M%S")
        stage_dir = os.path.join(self.backup_dir, f"backup_{ts}")
        os.makedirs(stage_dir, exist_ok=True)
        db_path = os.path.join(self.db_dir, "insurance_db.sqlite")
        if os.path.exists(db_path):
            self.conn.commit()
            shutil.copy2(db_path, os.path.join(stage_dir, "insurance_db.sqlite"))
        if os.path.isdir(self.archive_dir) and os.listdir(self.archive_dir):
            shutil.copytree(self.archive_dir, os.path.join(stage_dir, "Archive"))
        zip_path = shutil.make_archive(stage_dir, 'zip', stage_dir)
        shutil.rmtree(stage_dir, ignore_errors=True)
        return zip_path

    def setup_treeview_style(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview",
                        background="#0F172A",
                        foreground="#F1F5F9",
                        rowheight=40,
                        fieldbackground="#0F172A",
                        borderwidth=0,
                        font=("Vazir", 13))
        style.map('Treeview', background=[('selected', '#1e40af')]) # رنگ بک‌گراند ردیف انتخاب شده (آبی تیره)
        style.configure("Treeview.Heading",
                        background="#1e293b",
                        foreground="#38bdf8",
                        font=("Vazir", 14, "bold"),
                        borderwidth=1,
                        relief="flat")
        style.map("Treeview.Heading", background=[('active', '#334155')])

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children(): widget.destroy()

    def create_glass_card(self, parent, title, value, color):
        f = ctk.CTkFrame(parent, fg_color="#334155", corner_radius=15, border_width=1, border_color=color)
        f.pack(side="left", fill="both", expand=True, padx=10)
        ctk.CTkLabel(f, text=title, font=self.main_font, text_color="#CBD5E1").pack(pady=(15, 5))
        ctk.CTkLabel(f, text=value, font=self.big_font, text_color=color).pack(pady=(0, 15))

    def get_period(self, d_str):
        cutoff = int(self.config.get("cutoff_day", 25))
        try:
            y, m, d = [int(x) for x in str(d_str).split()[0].replace("-", "/").split("/")]
            if d >= cutoff:
                m += 1
                if m > 12: m, y = 1, y + 1
            return f"{MONTHS[m-1]} {y}"
        except Exception:
            logging.warning(f"get_period: تاریخ قابل تجزیه نبود: {d_str!r}")
            return "نامشخص"

    def get_due_date(self, period_str, inst_num):
        try:
            m_name, y_str = period_str.split()
            m_idx = MONTHS.index(m_name) + 1
            y, tm = int(y_str), m_idx + int(inst_num)
            while tm > 12: tm, y = tm - 12, y + 1
            return f"{y}/{tm:02d}/15"
        except Exception:
            logging.warning(f"get_due_date: ورودی نامعتبر period={period_str!r} inst_num={inst_num!r}")
            return ""

    def auto_resize_columns(self, tree, columns, data_rows):
        font = tkFont.Font(family="Vazir", size=13)
        for col_idx, col_name in enumerate(columns):
            max_w = font.measure(col_name) + 50
            if col_name == "انتخاب": max_w = 70
            for row in data_rows:
                w = font.measure(str(row[col_idx])) + 50
                if w > max_w: max_w = w
            tree.column(col_name, width=max_w, minwidth=max_w, stretch=False)

    def handle_tree_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        if region == "cell":
            col = tree.identify_column(event.x)
            if col == "#1":
                iid = tree.identify_row(event.y)
                if iid:
                    vals = list(tree.item(iid, "values"))
                    vals[0] = "☑" if vals[0] == "☐" else "☐"
                    tree.item(iid, values=vals)

    def handle_right_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        if region == "cell":
            iid = tree.identify_row(event.y)
            col_id = tree.identify_column(event.x)
            col_idx = int(col_id.replace('#', '')) - 1
            if iid and col_idx >= 0:
                val = tree.item(iid, "values")[col_idx]
                menu = tk.Menu(tree, tearoff=0, font=("Vazir", 12))
                menu.add_command(label="کپی مقدار این سلول", command=lambda: self.clipboard_clear() or self.clipboard_append(str(val)))
                menu.post(event.x_root, event.y_root)

    def is_date_in_range(self, date_str, start_str, end_str):
        if not date_str: return False
        eng_date = convert_persian_to_english_number(date_str)
        eng_start = convert_persian_to_english_number(start_str)
        eng_end = convert_persian_to_english_number(end_str)

        m_date = re.search(r'(\d{4})/(\d{1,2})/(\d{1,2})', eng_date)
        if not m_date: return False

        try:
            d_val = jdatetime.date(int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3)))

            s_val = None
            if eng_start:
                ms = re.search(r'(\d{4})/(\d{1,2})/(\d{1,2})', eng_start)
                if ms: s_val = jdatetime.date(int(ms.group(1)), int(ms.group(2)), int(ms.group(3)))

            e_val = None
            if eng_end:
                me = re.search(r'(\d{4})/(\d{1,2})/(\d{1,2})', eng_end)
                if me: e_val = jdatetime.date(int(me.group(1)), int(me.group(2)), int(me.group(3)))

            if s_val and d_val < s_val: return False
            if e_val and d_val > e_val: return False
            return True
        except Exception:
            logging.debug(f"is_date_in_range: تاریخ نامعتبر date={date_str!r} start={start_str!r} end={end_str!r}")
            return False

    # ================= صدور ورد =================
    def open_invoice_dialog(self):
        diag = Toplevel(self)
        diag.title("صدور صورت‌حساب")
        diag.geometry("450x350")
        diag.configure(bg="#1E293B")
        diag.transient(self)
        diag.grab_set()

        templates = [f for f in os.listdir(self.template_dir) if f.endswith(".docx")]
        if not templates: return ctk.CTkLabel(diag, text="قالبی یافت نشد!", text_color="red", font=self.main_font).pack()

        stemp = ctk.StringVar(value=templates[0])
        itype = ctk.StringVar(value="گروهی (کل ماموت)")

        ctk.CTkLabel(diag, text="انتخاب قالب ورد:", font=self.main_font).pack(pady=(20, 5))
        cb_temp = ctk.CTkOptionMenu(diag, variable=stemp, values=templates, font=self.main_font)
        cb_temp.pack(pady=5)

        ctk.CTkLabel(diag, text="نوع صدور:", font=self.main_font).pack(pady=(10, 5))
        cb_type = ctk.CTkOptionMenu(diag, variable=itype, values=["گروهی (کل ماموت)", "تفکیکی"], font=self.main_font)
        cb_type.pack(pady=5)

        def generate():
            try:
                self.generate_word_invoice(stemp.get(), itype.get())
                diag.destroy()
            except Exception as e:
                logging.exception("generate_word_invoice: خطا در صدور صورت‌حساب")
                messagebox.showerror("خطا", str(e), parent=diag)
        ctk.CTkButton(diag, text="تولید فایل", font=self.main_font, fg_color="#EA580C", command=generate).pack(pady=30)

    def generate_word_invoice(self, template_name, inv_type):
        if 'docx' not in globals(): return messagebox.showerror("خطا", "نصب کتابخانه python-docx الزامیست.")
        template_path = os.path.join(self.template_dir, template_name)
        cursor = self.conn.cursor()

        if inv_type == "گروهی (کل ماموت)":
            doc = docx.Document(template_path)
            data = cursor.execute("SELECT company_name, COUNT(*), SUM(premium) FROM policies GROUP BY company_name").fetchall()
            g_tot = sum(r[2] for r in data)
            for p in doc.paragraphs:
                p.text = p.text.replace('[نام_شرکت]', 'گروه صنعتی ماموت').replace('[مبلغ_کل]', to_persian_num(f"{g_tot:,}"))

            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            h = table.rows[0].cells
            h[0].text, h[1].text, h[2].text, h[3].text = "ردیف", "نام شرکت", "تعداد بیمه", "جمع حق بیمه (ریال)"
            for i, r in enumerate(data, 1):
                c = table.add_row().cells
                c[0].text, c[1].text, c[2].text, c[3].text = to_persian_num(str(i)), r[0], to_persian_num(str(r[1])), to_persian_num(f"{r[2]:,}")
            sp = os.path.join(self.archive_dir, f"صورتحساب_گروهی_{jdatetime.date.today().strftime('%Y%m%d')}.docx")
            doc.save(sp)
            messagebox.showinfo("موفق", f"صورت‌حساب صادر شد:\n{sp}")
        else:
            comps = [r[0] for r in cursor.execute("SELECT DISTINCT company_name FROM policies").fetchall()]
            for comp in comps:
                c_doc = docx.Document(template_path)
                pols = cursor.execute("SELECT person_name, policy_no, type, premium FROM policies WHERE company_name=?", (comp,)).fetchall()
                t_prem = sum(p[3] for p in pols)
                for p in c_doc.paragraphs:
                    p.text = p.text.replace('[نام_شرکت]', comp).replace('[مبلغ_کل]', to_persian_num(f"{t_prem:,}"))

                table = c_doc.add_table(rows=1, cols=5)
                table.style = 'Table Grid'
                h = table.rows[0].cells
                h[0].text, h[1].text, h[2].text, h[3].text, h[4].text = "ردیف", "نام", "شماره بیمه", "نوع", "مبلغ (ریال)"
                for i, r in enumerate(pols, 1):
                    c = table.add_row().cells
                    c[0].text, c[1].text, c[2].text, c[3].text, c[4].text = to_persian_num(str(i)), r[0], r[1], r[2], to_persian_num(f"{r[3]:,}")
                cdir = os.path.join(self.archive_dir, "Invoices_Individual")
                os.makedirs(cdir, exist_ok=True)
                c_doc.save(os.path.join(cdir, f"صورتحساب_{comp}.docx"))
            messagebox.showinfo("موفق", "صورت‌حساب‌های تفکیکی صادر شد.")

    # ================= داشبورد و جداول =================
    def show_dashboard(self):
        self.clear_main_frame()
        cursor = self.conn.cursor()

        tot_pols = cursor.execute("SELECT COUNT(*) FROM policies").fetchone()[0]
        cur_period = self.get_period(jdatetime.date.today().strftime("%Y/%m/%d"))
        cur_pols = cursor.execute("SELECT COUNT(*) FROM policies WHERE period=?", (cur_period,)).fetchone()[0]
        tot_comps = cursor.execute("SELECT COUNT(DISTINCT company_name) FROM policies").fetchone()[0]

        tf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        tf.pack(fill="x", pady=10, padx=10)
        self.create_glass_card(tf, "کل بیمه‌نامه‌ها", to_persian_num(str(tot_pols)), "#38BDF8")
        self.create_glass_card(tf, "بیمه‌های دوره جاری", to_persian_num(str(cur_pols)), "#4ADE80")
        self.create_glass_card(tf, "شرکت‌های فعال", to_persian_num(str(tot_comps)), "#FACC15")

        bf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bf.pack(fill="x", pady=5, padx=10)
        ctk.CTkButton(bf, text="ایجاد صورت‌حساب", fg_color="#EA580C", font=self.main_font, command=self.open_invoice_dialog).pack(side="right", padx=5)
        ctk.CTkButton(bf, text="تسویه اقساط انتخابی", fg_color="#16A34A", font=self.main_font, command=self.open_payment_dialog).pack(side="right", padx=5)
        ctk.CTkButton(bf, text="واریز پاسارگاد", fg_color="#9333EA", font=self.main_font, command=self.mark_pasargad).pack(side="right", padx=5)
        ctk.CTkButton(bf, text="اکسل پرداخت‌ها", fg_color="#0284C7", font=self.main_font, command=lambda: os.startfile(self.payments_excel_path) if os.path.exists(self.payments_excel_path) else None).pack(side="left", padx=5)

        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=5)
        self.tabview._segmented_button.configure(font=self.main_font)

        self.tab_comp = self.tabview.add("گروهی (شرکتی)")
        self.tab_det = self.tabview.add("جزئی (تفکیکی)")

        # فیلترهای بالا
        ff = ctk.CTkFrame(self.main_frame, fg_color="#0F172A", corner_radius=10)
        ff.pack(fill="x", padx=10, pady=5)

        self.status_filter = ctk.StringVar(value="پرداخت نشده")
        ctk.CTkLabel(ff, text="وضعیت:", font=self.main_font).pack(side="right", padx=5, pady=5)
        cb_sf = ctk.CTkOptionMenu(ff, variable=self.status_filter, values=["همه", "پرداخت نشده", "پرداخت شده", "پرداخت به پاسارگاد"], command=self.load_tables, font=self.main_font)
        cb_sf.pack(side="right", padx=5)

        self.period_filter = ctk.StringVar(value="همه دوره‌ها")
        ctk.CTkLabel(ff, text="بازه زمانی:", font=self.main_font).pack(side="right", padx=(20,0))
        cb_pf = ctk.CTkOptionMenu(ff, variable=self.period_filter, values=["همه دوره‌ها", "دوره جاری (این ماه)"], command=self.load_tables, font=self.main_font)
        cb_pf.pack(side="right", padx=5)

        # فیلتر تاریخ سررسید دوگانه
        ctk.CTkLabel(ff, text="سررسید از:", font=self.main_font).pack(side="right", padx=(20,0))
        self.due_start_var = ctk.StringVar()
        self.due_start_var.trace_add("write", lambda *args: self.load_tables())
        ctk.CTkEntry(ff, textvariable=self.due_start_var, placeholder_text="1405/01/01", font=self.main_font, width=100).pack(side="right", padx=5)

        ctk.CTkLabel(ff, text="تا:", font=self.main_font).pack(side="right", padx=(5,0))
        self.due_end_var = ctk.StringVar()
        self.due_end_var.trace_add("write", lambda *args: self.load_tables())
        ctk.CTkEntry(ff, textvariable=self.due_end_var, placeholder_text="1405/12/29", font=self.main_font, width=100).pack(side="right", padx=5)

        # ---------------- پنل جستجوی مجزا بالای جدول گروهی ----------------
        sf_comp = ctk.CTkFrame(self.tab_comp, fg_color="transparent")
        sf_comp.pack(fill="x", pady=2)
        self.comp_vars = {}
        for col in ["شرکت", "دوره", "شماره قسط"]:
            v = ctk.StringVar()
            v.trace_add("write", lambda *args: self.load_comp_table())
            self.comp_vars[col] = v
            # Placeholder با نام ستون
            ctk.CTkEntry(sf_comp, textvariable=v, placeholder_text=f"جستجو {col}", font=self.main_font, width=150, text_color="white", placeholder_text_color="gray").pack(side="right", padx=5)

        self.cols_comp = ("انتخاب", "شرکت", "دوره", "شماره قسط", "سررسید", "وضعیت", "مبلغ کل (ریال)")
        tree_f_comp = ctk.CTkFrame(self.tab_comp)
        tree_f_comp.pack(fill="both", expand=True)
        self.tree_comp = ttk.Treeview(tree_f_comp, columns=self.cols_comp, show="headings", selectmode="extended")
        vs_comp = ttk.Scrollbar(tree_f_comp, orient="vertical", command=self.tree_comp.yview)
        hs_comp = ttk.Scrollbar(tree_f_comp, orient="horizontal", command=self.tree_comp.xview)
        self.tree_comp.configure(yscrollcommand=vs_comp.set, xscrollcommand=hs_comp.set)
        hs_comp.pack(side="bottom", fill="x")
        vs_comp.pack(side="right", fill="y")
        self.tree_comp.pack(side="left", fill="both", expand=True)

        # ---------------- پنل جستجوی مجزا بالای جدول تفکیکی ----------------
        sf_det = ctk.CTkFrame(self.tab_det, fg_color="transparent")
        sf_det.pack(fill="x", pady=2)
        self.det_vars = {}
        for col in ["شرکت", "بیمه‌گذار", "شماره بیمه"]:
            v = ctk.StringVar()
            v.trace_add("write", lambda *args: self.load_det_table())
            self.det_vars[col] = v
            # Placeholder با نام ستون
            ctk.CTkEntry(sf_det, textvariable=v, placeholder_text=f"جستجو {col}", font=self.main_font, width=180, text_color="white", placeholder_text_color="gray").pack(side="right", padx=5)

        self.cols_det = ("انتخاب", "شناسه", "شرکت", "دوره", "بیمه‌گذار", "شماره بیمه", "قسط", "سررسید", "مبلغ (ریال)", "وضعیت")
        tree_f_det = ctk.CTkFrame(self.tab_det)
        tree_f_det.pack(fill="both", expand=True)
        self.tree_det = ttk.Treeview(tree_f_det, columns=self.cols_det, show="headings", selectmode="extended")
        vs_det = ttk.Scrollbar(tree_f_det, orient="vertical", command=self.tree_det.yview)
        hs_det = ttk.Scrollbar(tree_f_det, orient="horizontal", command=self.tree_det.xview)
        self.tree_det.configure(yscrollcommand=vs_det.set, xscrollcommand=hs_det.set)
        hs_det.pack(side="bottom", fill="x")
        vs_det.pack(side="right", fill="y")
        self.tree_det.pack(side="left", fill="both", expand=True)

        for t in [self.tree_comp, self.tree_det]:
            # رنگ‌های جدید و روشن‌تر برای خوانایی بهتر روی پس زمینه تیره
            t.tag_configure("unpaid", foreground="#FF6B6B") # قرمز روشن
            t.tag_configure("paid", foreground="#4ADE80") # سبز روشن
            t.tag_configure("pasargad", foreground="#C084FC") # بنفش
            t.bind("<ButtonRelease-1>", lambda e, tree=t: self.handle_tree_click(e, tree))
            t.bind("<Button-3>", lambda e, tree=t: self.handle_right_click(e, tree))

        self.tree_comp.bind("<Double-1>", self.show_selected_comp_details)

        for c in self.cols_comp: self.tree_comp.heading(c, text=c)
        for c in self.cols_det: self.tree_det.heading(c, text=c)

        self.load_tables()

    def load_tables(self, *args):
        self.load_comp_table()
        self.load_det_table()

    def get_sql_conditions(self):
        conds, params = [], []
        if self.status_filter.get() != "همه":
            conds.append("i.status = ?")
            params.append(self.status_filter.get())
        if self.period_filter.get() == "دوره جاری (این ماه)":
            cur_per = self.get_period(jdatetime.date.today().strftime("%Y/%m/%d"))
            conds.append("p.period = ?")
            params.append(cur_per)
        where = " WHERE " + " AND ".join(conds) if conds else ""
        return where, params

    def passes_date_range(self, due_str):
        start_str = self.due_start_var.get().strip()
        end_str = self.due_end_var.get().strip()
        if not start_str and not end_str: return True
        return self.is_date_in_range(due_str, start_str, end_str)

    def load_comp_table(self):
        for item in self.tree_comp.get_children(): self.tree_comp.delete(item)

        q = '''SELECT p.company_name, p.period, i.inst_num, i.status, SUM(i.amount), i.due_date
               FROM installments i JOIN policies p ON i.policy_id = p.id'''
        where, params = self.get_sql_conditions()
        q += where
        q += " GROUP BY p.company_name, p.period, i.inst_num, i.status, i.due_date ORDER BY p.company_name, i.inst_num"

        rows = self.conn.cursor().execute(q, params).fetchall()

        search_comp = self.comp_vars["شرکت"].get().strip().lower()
        search_per = self.comp_vars["دوره"].get().strip().lower()
        search_in = convert_persian_to_english_number(self.comp_vars["شماره قسط"].get().strip())

        final_rows = []
        for r in rows:
            comp, per, inum, stat, amt, due = r
            due_f = due if due else self.get_due_date(per, inum)

            if not self.passes_date_range(due_f): continue
            if search_comp and search_comp not in str(comp).lower(): continue
            if search_per and search_per not in str(per).lower(): continue
            if search_in and search_in != str(inum): continue

            vals = ("☐", comp, per, to_persian_num(str(inum)), to_persian_num(due_f), stat, to_persian_num(f"{amt:,}"))
            tag = "paid" if stat == "پرداخت شده" else "pasargad" if stat == "پرداخت به پاسارگاد" else "unpaid"
            iid = f"{comp}|{per}|{inum}|{stat}"
            self.tree_comp.insert("", "end", iid=iid, values=vals, tags=(tag,))
            final_rows.append(vals)

        self.auto_resize_columns(self.tree_comp, self.cols_comp, final_rows)

    def show_selected_comp_details(self, event=None):
        selected = self.tree_comp.selection()
        if not selected: return messagebox.showwarning("هشدار", "یک ردیف را دابل کلیک کنید.")
        iid = selected[0]
        comp, period, i_num, status = iid.split('|')

        diag = Toplevel(self)
        diag.title(f"ریز اقساط: {comp} - قسط {to_persian_num(str(i_num))}")
        diag.geometry("900x450")
        diag.configure(bg="#1E293B")
        diag.transient(self)

        cols = ("بیمه‌گذار", "شماره بیمه", "سررسید", "مبلغ (ریال)", "وضعیت")
        ts = ttk.Treeview(diag, columns=cols, show="headings")
        for c in cols: ts.heading(c, text=c)
        ts.tag_configure("unpaid", foreground="#FF6B6B")
        ts.tag_configure("paid", foreground="#4ADE80")
        ts.tag_configure("pasargad", foreground="#C084FC")

        sb = ttk.Scrollbar(diag, orient="vertical", command=ts.yview)
        ts.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        ts.pack(fill="both", expand=True, padx=10, pady=10)
        ts.bind("<Button-3>", lambda e, t=ts: self.handle_right_click(e, t))

        q = '''SELECT p.person_name, p.policy_no, i.amount, i.status, i.due_date
               FROM installments i JOIN policies p ON i.policy_id = p.id
               WHERE p.company_name=? AND p.period=? AND i.inst_num=? AND i.status=?'''
        rows = self.conn.cursor().execute(q, (comp, period, int(i_num), status)).fetchall()

        data_rows = []
        for r in rows:
            due = to_persian_num(r[4] if r[4] else self.get_due_date(period, int(i_num)))
            v = (r[0], r[1], due, to_persian_num(f"{r[2]:,}"), r[3])
            tag = "paid" if r[3] == "پرداخت شده" else "pasargad" if r[3] == "پرداخت به پاسارگاد" else "unpaid"
            ts.insert("", "end", values=v, tags=(tag,))
            data_rows.append(v)
        self.auto_resize_columns(ts, cols, data_rows)

    def load_det_table(self):
        for item in self.tree_det.get_children(): self.tree_det.delete(item)

        q = '''SELECT i.id, p.company_name, p.period, p.person_name, p.policy_no, i.inst_num, i.due_date, i.amount, i.status
               FROM installments i JOIN policies p ON i.policy_id = p.id'''
        where, params = self.get_sql_conditions()
        q += where
        q += " ORDER BY p.company_name, p.id, i.inst_num LIMIT 1000"

        rows = self.conn.cursor().execute(q, params).fetchall()

        s_comp = self.det_vars["شرکت"].get().strip().lower()
        s_name = self.det_vars["بیمه‌گذار"].get().strip().lower()
        s_pol = self.det_vars["شماره بیمه"].get().strip().lower()

        final_rows = []
        for r in rows:
            _id, comp, per, name, pol, inum, due, amt, stat = r
            due_f = due if due else self.get_due_date(per, inum)

            if not self.passes_date_range(due_f): continue
            if s_comp and s_comp not in str(comp).lower(): continue
            if s_name and s_name not in str(name).lower(): continue
            if s_pol and s_pol not in str(pol).lower(): continue

            vals = ("☐", _id, comp, per, name, pol, to_persian_num(str(inum)), to_persian_num(due_f), to_persian_num(f"{amt:,}"), stat)
            tag = "paid" if stat == "پرداخت شده" else "pasargad" if stat == "پرداخت به پاسارگاد" else "unpaid"
            self.tree_det.insert("", "end", iid=str(_id), values=vals, tags=(tag,))
            final_rows.append(vals)

        self.auto_resize_columns(self.tree_det, self.cols_det, final_rows)

    def verify_sequential_payment(self, cursor, policy_id, inst_num):
        if self.config.get("enforce_sequential_payment", "فعال") != "فعال":
            return True, ""

        if inst_num == 1:
            return True, ""

        q = "SELECT inst_num, status FROM installments WHERE policy_id = ? AND inst_num < ? ORDER BY inst_num ASC"
        prev_insts = cursor.execute(q, (policy_id, inst_num)).fetchall()

        for p_num, p_stat in prev_insts:
            if p_stat == "پرداخت نشده":
                return False, f"برای پرداخت قسط {inst_num}، ابتدا باید قسط {p_num} تسویه شود."
        return True, ""

    def open_payment_dialog(self):
        active = self.tabview.get()
        cursor = self.conn.cursor()

        selected_ids_to_pay = []

        if active == "گروهی (شرکتی)":
            selected_rows = [i for i in self.tree_comp.get_children() if self.tree_comp.item(i, "values")[0] == "☑"]
            if not selected_rows: return messagebox.showwarning("هشدار", "تیک حداقل یک گروه را فعال کنید.")

            for row_iid in selected_rows:
                comp, per, inum, stat = row_iid.split('|')
                res = cursor.execute("SELECT i.id, p.id FROM installments i JOIN policies p ON i.policy_id=p.id WHERE p.company_name=? AND p.period=? AND i.inst_num=? AND i.status=?", (comp, per, int(inum), stat)).fetchall()
                for r_id, p_id in res:
                    ok, msg = self.verify_sequential_payment(cursor, p_id, int(inum))
                    if not ok:
                        messagebox.showerror("خطای تسویه پیوسته", f"در شرکت {comp}:\n{msg}")
                        return
                    selected_ids_to_pay.append(r_id)
        else:
            selected_rows = [i for i in self.tree_det.get_children() if self.tree_det.item(i, "values")[0] == "☑"]
            if not selected_rows: return messagebox.showwarning("هشدار", "تیک حداقل یک قسط را فعال کنید.")

            for row_iid in selected_rows:
                r_id = int(row_iid)
                p_id, i_num = cursor.execute("SELECT policy_id, inst_num FROM installments WHERE id=?", (r_id,)).fetchone()
                ok, msg = self.verify_sequential_payment(cursor, p_id, i_num)
                if not ok:
                    messagebox.showerror("خطای تسویه پیوسته", msg)
                    return
                selected_ids_to_pay.append(r_id)

        placeholders = ",".join("?" * len(selected_ids_to_pay))
        tot_amt = cursor.execute(f"SELECT SUM(amount) FROM installments WHERE id IN ({placeholders})", selected_ids_to_pay).fetchone()[0]

        diag = Toplevel(self)
        diag.title("ثبت پرداختی و تسویه")
        diag.geometry("550x550")
        diag.configure(bg="#1E293B")
        diag.transient(self)
        diag.grab_set()

        ctk.CTkLabel(diag, text=f"مبلغ کل انتخابی: {to_persian_num(f'{tot_amt:,}')} ریال", font=self.title_font).pack(pady=20)

        pay_mode = ctk.StringVar(value="full")
        def toggle_ent(): ent_amt.configure(state="normal" if pay_mode.get() == "partial" else "disabled")

        frb = ctk.CTkFrame(diag, fg_color="transparent")
        frb.pack(fill="x")
        ctk.CTkRadioButton(frb, text="تسویه کامل", variable=pay_mode, value="full", font=self.main_font, command=toggle_ent).pack(side="right", padx=10)
        ctk.CTkRadioButton(frb, text="مبلغ دلخواه (ثبت کسری)", variable=pay_mode, value="partial", font=self.main_font, command=toggle_ent).pack(side="right", padx=10)

        ent_amt = ctk.CTkEntry(diag, font=self.main_font, placeholder_text="مبلغ پرداختی (ریال)", state="disabled")
        ent_amt.pack(pady=10, fill="x", padx=40)

        ctk.CTkLabel(diag, text="نوع پرداخت:", font=self.main_font).pack(pady=(10,0))
        p_type = ctk.StringVar(value="فیش بانکی")
        cb_pt = ctk.CTkOptionMenu(diag, variable=p_type, values=["فیش بانکی", "نقدی", "چک", "کسر از حقوق"], font=self.main_font)
        cb_pt.pack(pady=5)

        ctk.CTkLabel(diag, text="توضیحات:", font=self.main_font).pack(pady=(10,0))
        desc = ctk.CTkEntry(diag, font=self.main_font, width=300)
        desc.pack(pady=5)

        self.cur_receipts = [] # پشتیبانی از چند فیش
        def sel_file():
            paths = filedialog.askopenfilenames()
            if paths:
                self.cur_receipts.extend(paths)
                lbl_file.configure(text=f"{len(self.cur_receipts)} فایل انتخاب شد")
        ctk.CTkButton(diag, text="انتخاب تصویر فیش/رسید", font=self.main_font, fg_color="#0284C7", command=sel_file).pack(pady=10)
        lbl_file = ctk.CTkLabel(diag, text="سندی انتخاب نشده", font=self.main_font, text_color="gray")
        lbl_file.pack()

        def commit():
            amt = tot_amt
            if pay_mode.get() == "partial":
                if active == "گروهی (شرکتی)":
                    diag.destroy()
                    messagebox.showerror("خطا", "برای پرداخت ناقص گروهی، باید وارد بخش «ریز اقساط» شوید.")
                    return
                amt = clean_number(ent_amt.get())
                if amt <= 0 or amt >= tot_amt: return messagebox.showerror("خطا", "مبلغ نامعتبر است.", parent=diag)

            saved_receipt_paths = []
            if self.cur_receipts:
                sample_id = selected_ids_to_pay[0]
                i_num_str = cursor.execute("SELECT inst_num FROM installments WHERE id=?", (sample_id,)).fetchone()[0]
                comp = cursor.execute("SELECT p.company_name FROM installments i JOIN policies p ON i.policy_id=p.id WHERE i.id=?", (sample_id,)).fetchone()[0]
                per = cursor.execute("SELECT p.period FROM installments i JOIN policies p ON i.policy_id=p.id WHERE i.id=?", (sample_id,)).fetchone()[0]

                r_dir = os.path.join(self.archive_dir, per, comp, "Receipts")
                os.makedirs(r_dir, exist_ok=True)

                for idx, r_path in enumerate(self.cur_receipts, 1):
                    ext = os.path.splitext(r_path)[1]
                    t_sh = jdatetime.date.today().strftime('%Y.%m.%d')
                    fname = f"{t_sh} - پرداخت قسط {i_num_str} - {comp} - دوره {per}_{idx}{ext}".replace(' ', '_')
                    sp = os.path.join(r_dir, fname)
                    shutil.copy(r_path, sp)
                    saved_receipt_paths.append(sp)

            final_receipt_str = ";".join(saved_receipt_paths)
            now_d = jdatetime.date.today().strftime('%Y/%m/%d')
            placeholders = ",".join("?" * len(selected_ids_to_pay))

            if pay_mode.get() == "full":
                cursor.execute(f"UPDATE installments SET status='پرداخت شده', payment_type=?, description=?, receipt_path=?, payment_date=? WHERE id IN ({placeholders})",
                               (p_type.get(), desc.get(), final_receipt_str, now_d, *selected_ids_to_pay))
            else:
                cursor.execute(f"SELECT id, policy_id, inst_num, amount FROM installments WHERE id IN ({placeholders}) ORDER BY id", selected_ids_to_pay)
                for r_id, p_id, i_n, i_a in cursor.fetchall():
                    if amt >= i_a:
                        cursor.execute("UPDATE installments SET status='پرداخت شده', payment_type=?, description=?, receipt_path=?, payment_date=? WHERE id=?", (p_type.get(), desc.get(), final_receipt_str, now_d, r_id))
                        amt -= i_a
                    elif amt > 0:
                        cursor.execute("UPDATE installments SET amount=?, status='پرداخت شده', payment_type=?, description=?, receipt_path=?, payment_date=? WHERE id=?", (amt, p_type.get(), desc.get(), final_receipt_str, now_d, r_id))
                        rem = i_a - amt
                        cursor.execute("INSERT INTO installments (policy_id, inst_num, amount, status) VALUES (?, ?, ?, 'پرداخت نشده')", (p_id, i_n, rem))
                        amt = 0
                    else: break
            self.conn.commit()
            logging.info(f"پرداخت توسط «{self.current_user}» ثبت شد: {len(selected_ids_to_pay)} قسط، حالت={pay_mode.get()}")
            self.generate_excel_reports()
            self.load_tables()
            diag.destroy()
            messagebox.showinfo("موفق", "پرداختی ثبت شد.")

        ctk.CTkButton(diag, text="تایید نهایی", fg_color="#16A34A", font=self.title_font, command=commit).pack(pady=20)

    def mark_pasargad(self):
        active = self.tabview.get()
        tree = self.tree_comp if active == "گروهی (شرکتی)" else self.tree_det
        ids = [i for i in tree.get_children() if tree.item(i, "values")[0] == "☑"]
        if not ids: return messagebox.showwarning("هشدار", "هیچ تیکی در جدول فعال نیست.")

        cursor = self.conn.cursor()
        real_ids = []
        if active == "گروهی (شرکتی)":
            for key in ids:
                c, p, i_n, s = key.split('|')
                res = cursor.execute("SELECT i.id, i.status FROM installments i JOIN policies p ON i.policy_id=p.id WHERE p.company_name=? AND p.period=? AND i.inst_num=? AND i.status=?", (c, p, int(i_n), s)).fetchall()
                for r in res:
                    if r[1] != 'پرداخت شده': return messagebox.showerror("خطا", "فقط اقساط تسویه شده به پاسارگاد می‌روند.")
                    real_ids.append(r[0])
        else:
            real_ids = [int(tree.item(i, "values")[1]) for i in ids]
            placeholders = ",".join("?" * len(real_ids))
            sts = cursor.execute(f"SELECT status FROM installments WHERE id IN ({placeholders})", real_ids).fetchall()
            if any(s[0] != 'پرداخت شده' for s in sts): return messagebox.showerror("خطا", "فقط اقساط تسویه شده به پاسارگاد می‌روند.")

        placeholders = ",".join("?" * len(real_ids))
        cursor.execute(f"UPDATE installments SET status='پرداخت به پاسارگاد' WHERE id IN ({placeholders})", real_ids)
        self.conn.commit()
        self.generate_excel_reports()
        self.load_tables()
        logging.info(f"{len(real_ids)} قسط به «پرداخت به پاسارگاد» تغییر یافت. کاربر: {self.current_user}")
        messagebox.showinfo("موفق", "تغییر وضعیت انجام شد.")

    # ================= خروجی اکسل =================
    def generate_excel_reports(self):
        cursor = self.conn.cursor()
        data = cursor.execute('''SELECT p.company_name, p.period, p.person_name, p.policy_no,
                          i.inst_num, i.amount, i.due_date, i.status, i.payment_type, i.description, i.payment_date, i.receipt_path
                          FROM installments i JOIN policies p ON i.policy_id = p.id ORDER BY p.company_name, p.id, i.inst_num''').fetchall()
        if not data: return
        df = pd.DataFrame(data, columns=["شرکت", "دوره", "بیمه‌گذار", "شماره بیمه", "قسط", "مبلغ قسط", "سررسید", "وضعیت", "نوع پرداخت", "توضیحات", "تاریخ پرداخت", "فیش"])
        df['کاربر ثبت‌کننده'] = self.current_user
        df.to_excel(self.payments_excel_path, index=False)
        self.apply_b_nazanin(self.payments_excel_path)

        comps = df['شرکت'].unique()
        for comp in comps:
            c_df = df[df['شرکت'] == comp]
            period = c_df['دوره'].iloc[0]
            c_dir = os.path.join(self.archive_dir, period, comp)
            os.makedirs(c_dir, exist_ok=True)

            pay_path = os.path.join(c_dir, f"گزارش_اقساط_{comp}.xlsx")
            c_df.to_excel(pay_path, index=False)
            self.apply_b_nazanin(pay_path, hyperlink_col="فیش")

            pol_path = os.path.join(c_dir, f"بیمه‌های_صادره_{comp}.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "بیمه‌های صادره و وضعیت"

            headers = ["شماره بیمه", "بیمه‌گذار", "تاریخ صدور", "حق بیمه کل"]
            for i in range(1, 10): headers.extend([f"مبلغ قسط {i}", f"سررسید {i}"])
            ws.append(headers)

            pols = cursor.execute("SELECT id, policy_no, person_name, issue_date, premium FROM policies WHERE company_name=? AND period=?", (comp, period)).fetchall()
            green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

            for pol in pols:
                row_data = [pol[1], pol[2], pol[3], pol[4]]
                insts = cursor.execute("SELECT inst_num, amount, due_date, status FROM installments WHERE policy_id=? ORDER BY inst_num", (pol[0],)).fetchall()

                for i in range(1, 10):
                    inst = next((x for x in insts if x[0] == i), None)
                    if inst: row_data.extend([inst[1], inst[2] if inst[2] else self.get_due_date(period, i)])
                    else: row_data.extend(["", ""])
                ws.append(row_data)

                current_row = ws.max_row
                for i in range(1, 10):
                    inst = next((x for x in insts if x[0] == i), None)
                    if inst and inst[3] == "پرداخت شده":
                        col_amt, col_due = 4 + (i*2) - 1, 4 + (i*2)
                        ws.cell(row=current_row, column=col_amt).fill = green
                        ws.cell(row=current_row, column=col_due).fill = green
                        ws.cell(row=current_row, column=col_amt).hyperlink = f"گزارش_اقساط_{comp}.xlsx"
                        ws.cell(row=current_row, column=col_amt).font = Font(color="0000FF", underline="single")

            for row in ws.iter_rows():
                for cell in row:
                    if cell.value:
                        font_k = {'name': 'B Nazanin'}
                        if cell.row == 1: font_k['bold'] = True
                        if cell.hyperlink:
                            font_k['color'] = "0000FF"
                            font_k['underline'] = "single"
                        cell.font = Font(**font_k)
            wb.save(pol_path)

    def apply_b_nazanin(self, path, hyperlink_col=None):
        try:
            wb = load_workbook(path)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                hl_idx = -1
                if hyperlink_col:
                    for c in ws[1]:
                        if c.value == hyperlink_col: hl_idx = c.column
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value:
                            cell.font = Font(name='B Nazanin', bold=(cell.row == 1))
                            if cell.column == hl_idx and cell.row > 1 and str(cell.value).endswith(('.jpg','.png','.pdf')):
                                cell.hyperlink = str(cell.value)
                                cell.font = Font(name='B Nazanin', color="0000FF", underline="single")
            wb.save(path)
        except Exception:
            logging.exception(f"apply_b_nazanin: قالب‌بندی فایل ناموفق بود: {path}")

    # ================= پردازش هوشمند =================
    def show_processing(self):
        self.clear_main_frame()
        self.b_path = self.s_path = None

        ctk.CTkLabel(self.main_frame, text="آپلود و پردازش هوشمند", font=self.title_font).pack(pady=20)

        for ftype, txt in [('b', "فایل بدنه"), ('s', "فایل ثالث")]:
            f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
            f.pack(pady=5)
            ctk.CTkButton(f, text=txt, font=self.main_font, command=lambda t=ftype: self.sel_file(t)).pack(side="right", padx=10)
            lbl = ctk.CTkLabel(f, text="انتخاب نشده", font=self.main_font)
            lbl.pack(side="right")
            if ftype == 'b': self.lbl_b_file = lbl
            else: self.lbl_s_file = lbl

        bf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bf.pack(pady=15)
        ctk.CTkButton(bf, text="پیش‌پردازش فایل‌ها", font=self.title_font, fg_color="#EA580C", command=self.run_preview).pack(side="right", padx=10)
        ctk.CTkButton(bf, text="افزودن دستی", font=self.title_font, fg_color="#0284C7", command=self.add_manual_row).pack(side="left", padx=10)

        self.preview_scroll = ctk.CTkScrollableFrame(self.main_frame, fg_color="#0F172A")
        self.preview_scroll.pack(fill="both", expand=True, padx=20, pady=10)
        self.btn_commit = ctk.CTkButton(self.main_frame, text="پردازش نهایی", font=self.title_font, fg_color="#16A34A", command=self.commit_processing, state="disabled")
        self.btn_commit.pack(pady=10)
        self.render_preview_ui()

    def sel_file(self, ftype):
        p = filedialog.askopenfilename()
        if p:
            if ftype == 'b':
                self.b_path = p
                self.lbl_b_file.configure(text=to_persian_num(os.path.basename(p)))
            else:
                self.s_path = p
                self.lbl_s_file.configure(text=to_persian_num(os.path.basename(p)))

    def add_manual_row(self):
        diag = Toplevel(self)
        diag.title("افزودن ردیف")
        diag.geometry("450x550")
        diag.configure(bg="#1E293B")
        diag.transient(self)

        fields = [("نام شرکت:", "company"), ("بیمه‌گذار:", "name"), ("شماره بیمه:", "policy_no"), ("حق بیمه:", "premium"), ("تاریخ صدور:", "date")]
        ents = {}
        for lbl, k in fields:
            f = ctk.CTkFrame(diag, fg_color="transparent")
            f.pack(fill="x", padx=20, pady=5)
            ctk.CTkLabel(f, text=lbl, font=self.main_font, width=120, anchor="e").pack(side="right", padx=10)
            e = ctk.CTkEntry(f, font=self.main_font)
            e.pack(side="right", fill="x", expand=True)
            ents[k] = e

        ft = ctk.CTkFrame(diag, fg_color="transparent")
        ft.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(ft, text="نوع بیمه:", font=self.main_font, width=120, anchor="e").pack(side="right", padx=10)
        t_var = ctk.StringVar(value="ثالث")
        cb_type = ctk.CTkOptionMenu(ft, variable=t_var, values=["ثالث", "بدنه"], font=self.main_font)
        cb_type.pack(side="right", fill="x", expand=True)

        def save():
            c, p = ents["company"].get().strip(), ents["policy_no"].get().strip()
            if not c or not p: return messagebox.showerror("خطا", "اطلاعات ناقص است.", parent=diag)
            p_c = clean_text_advanced(p)
            if self.conn.cursor().execute("SELECT id FROM policies WHERE policy_no=?", (p_c,)).fetchone():
                return messagebox.showerror("خطا", "بیمه‌نامه موجود است!", parent=diag)
            if c not in self.pending_preview_data: self.pending_preview_data[c] = []
            self.pending_preview_data[c].append({
                "policy_no": p_c, "name": clean_text_advanced(ents["name"].get()),
                "premium": clean_number(ents["premium"].get()), "date": clean_text_advanced(ents["date"].get()), "type": t_var.get()
            })
            self.render_preview_ui()
            diag.destroy()
        ctk.CTkButton(diag, text="ثبت", fg_color="#16A34A", font=self.main_font, command=save).pack(pady=20)

    def edit_preview_company(self, comp):
        diag = Toplevel(self)
        diag.title(f"ویرایش: {comp}")
        diag.geometry("800x600")
        diag.configure(bg="#1E293B")
        diag.transient(self)

        sf = ctk.CTkScrollableFrame(diag, fg_color="#0F172A")
        sf.pack(fill="both", expand=True, padx=10, pady=10)

        elist = []
        for i, item in enumerate(self.pending_preview_data[comp]):
            f = ctk.CTkFrame(sf, fg_color="transparent")
            f.pack(fill="x", pady=5)
            en = ctk.CTkEntry(f, font=self.main_font, width=150)
            en.insert(0, item['name']); en.pack(side="right", padx=2)
            ep = ctk.CTkEntry(f, font=self.main_font, width=150)
            ep.insert(0, item['policy_no']); ep.pack(side="right", padx=2)
            epr = ctk.CTkEntry(f, font=self.main_font, width=120)
            epr.insert(0, str(item['premium'])); epr.pack(side="right", padx=2)
            ed = ctk.CTkEntry(f, font=self.main_font, width=100)
            ed.insert(0, item['date']); ed.pack(side="right", padx=2)
            elist.append((en, ep, epr, ed))

        def save_e():
            for idx, (en, ep, epr, ed) in enumerate(elist):
                self.pending_preview_data[comp][idx]['name'] = clean_text_advanced(en.get())
                self.pending_preview_data[comp][idx]['policy_no'] = clean_text_advanced(ep.get())
                self.pending_preview_data[comp][idx]['premium'] = clean_number(epr.get())
                self.pending_preview_data[comp][idx]['date'] = clean_text_advanced(ed.get())
            self.render_preview_ui()
            diag.destroy()
        ctk.CTkButton(diag, text="ذخیره", fg_color="#16A34A", font=self.main_font, command=save_e).pack(pady=10)

    def render_preview_ui(self):
        for w in self.preview_scroll.winfo_children(): w.destroy()
        if not self.pending_preview_data:
            ctk.CTkLabel(self.preview_scroll, text="رکوردی یافت نشد.", font=self.main_font).pack(pady=20)
            self.btn_commit.configure(state="disabled")
            return
        self.btn_commit.configure(state="normal")
        for comp, items in list(self.pending_preview_data.items()):
            if not items:
                del self.pending_preview_data[comp]; continue
            f = ctk.CTkFrame(self.preview_scroll, fg_color="#334155")
            f.pack(fill="x", pady=5, padx=5)
            t_prem = sum(i['premium'] for i in items)
            ctk.CTkLabel(f, text=f"{to_persian_num(comp)} | {to_persian_num(str(len(items)))} مورد | {to_persian_num(f'{t_prem:,}')} ریال", font=self.main_font, text_color="#4ADE80").pack(side="right", padx=10, pady=10)
            ctk.CTkButton(f, text="حذف", fg_color="#E11D48", width=60, font=self.main_font, command=lambda c=comp: [self.pending_preview_data.pop(c), self.render_preview_ui()]).pack(side="left", padx=5)
            ctk.CTkButton(f, text="ویرایش", fg_color="#0284C7", width=60, font=self.main_font, command=lambda c=comp: self.edit_preview_company(c)).pack(side="left", padx=5)

    def run_preview(self):
        tc = [clean_contract_number(t) for t in self.config.get("target_contracts", [])]
        dl = []
        if getattr(self, 'b_path', None): dl.append((self.b_path, 'b', 'بدنه'))
        if getattr(self, 's_path', None): dl.append((self.s_path, 's', 'ثالث'))
        if not dl: return messagebox.showwarning("هشدار", "فایلی انتخاب نشده")

        epols = {r[0] for r in self.conn.cursor().execute("SELECT policy_no FROM policies").fetchall()}

        for path, pfx, t_name in dl:
            try:
                df = pd.read_excel(path)
                df.columns = df.columns.astype(str).str.strip()
                ct = resolve_column(df, self.config.get(f"{pfx}_type"), "type")
                cc = resolve_column(df, self.config.get(f"{pfx}_company"), "company")
                cco = resolve_column(df, self.config.get(f"{pfx}_contract"), "contract")
                cd = resolve_column(df, self.config.get(f"{pfx}_date"), "date")
                cpr = resolve_column(df, self.config.get(f"{pfx}_premium"), "premium")
                cn = resolve_column(df, self.config.get(f"{pfx}_name"), "name")
                cp = resolve_column(df, self.config.get(f"{pfx}_policy"), "policy")
                if None in [ct, cc, cco, cd, cpr, cn, cp]: continue

                dft = df[ct].apply(clean_text_advanced)
                dfco = df[cco].apply(clean_contract_number)
                for _, row in df[(dft.str.contains("حقیقی", na=False)) & (dfco.isin(tc))].iterrows():
                    pol = clean_text_advanced(row.get(cp, ""))
                    comp = clean_text_advanced(row.get(cc, ""))
                    if not pol or pol in epols or not comp or comp == "nan": continue

                    dup = False
                    for c_items in self.pending_preview_data.values():
                        if any(i['policy_no'] == pol for i in c_items): dup = True; break
                    if dup: continue

                    if comp not in self.pending_preview_data: self.pending_preview_data[comp] = []
                    self.pending_preview_data[comp].append({
                        "policy_no": pol, "name": clean_text_advanced(row.get(cn, "")),
                        "premium": clean_number(row.get(cpr, 0)), "date": clean_text_advanced(row.get(cd, "")), "type": t_name
                    })
            except Exception as e:
                logging.exception(f"run_preview: خطا در پردازش فایل {path}")
                messagebox.showerror("خطا", str(e))
        self.render_preview_ui()

    def commit_processing(self):
        cursor = self.conn.cursor()
        for comp, items in self.pending_preview_data.items():
            for it in items:
                per = self.get_period(it['date'])
                cursor.execute('''INSERT OR IGNORE INTO policies (policy_no, person_name, company_name, premium, issue_date, period, type)
                                  VALUES (?, ?, ?, ?, ?, ?, ?)''', (it['policy_no'], it['name'], comp, it['premium'], it['date'], per, it['type']))
                if cursor.rowcount > 0:
                    pid = cursor.lastrowid
                    base, rem = it['premium'] // 9, it['premium'] % 9
                    for i in range(1, 10):
                        due = self.get_due_date(per, i)
                        cursor.execute("INSERT INTO installments (policy_id, inst_num, amount, due_date) VALUES (?, ?, ?, ?)", (pid, i, base + rem if i == 1 else base, due))
        self.conn.commit()
        logging.info(f"پردازش هوشمند توسط «{self.current_user}» انجام شد: {sum(len(v) for v in self.pending_preview_data.values())} بیمه‌نامه.")
        self.generate_excel_reports()
        messagebox.showinfo("موفقیت", "پردازش انجام شد.")
        self.show_dashboard()

    # ================= تنظیمات =================
    def show_settings(self):
        self.clear_main_frame()
        s = ctk.CTkScrollableFrame(self.main_frame, fg_color="transparent")
        s.pack(fill="both", expand=True)
        ctk.CTkLabel(s, text="تنظیمات سیستم", font=self.title_font).pack(pady=10)

        # بخش الزام تسویه پیوسته
        f_sq = ctk.CTkFrame(s, fg_color="#334155", corner_radius=10)
        f_sq.pack(fill="x", pady=10, padx=20)
        ctk.CTkLabel(f_sq, text="قانون تسویه پیوسته اقساط:", font=self.main_font).pack(side="right", padx=10, pady=10)
        self.var_sq = ctk.StringVar(value=self.config.get("enforce_sequential_payment", "فعال"))
        cb_sq = ctk.CTkOptionMenu(f_sq, variable=self.var_sq, values=["فعال", "غیرفعال"], font=self.main_font)
        cb_sq.pack(side="right", padx=10, pady=10)

        self.entries = {}
        fields = [("بیمه‌گذار ثالث:", "s_type"), ("طرف قرارداد ثالث:", "s_company"), ("قرارداد ثالث:", "s_contract"),
                  ("تاریخ صدور ثالث:", "s_date"), ("حق بیمه ثالث:", "s_premium"), ("نام ثالث:", "s_name"), ("شماره ثالث:", "s_policy"),
                  ("بیمه‌گذار بدنه:", "b_type"), ("طرف قرارداد بدنه:", "b_company"), ("قرارداد بدنه:", "b_contract"),
                  ("تاریخ صدور بدنه:", "b_date"), ("حق بیمه بدنه:", "b_premium"), ("نام بدنه:", "b_name"), ("شماره بدنه:", "b_policy"),
                  ("روز کات‌آف چرخه ماهانه:", "cutoff_day")]

        for lbl, key in fields:
            f = ctk.CTkFrame(s, fg_color="transparent")
            f.pack(fill="x", pady=2)
            ctk.CTkLabel(f, text=lbl, font=self.main_font, width=200, anchor="e").pack(side="right", padx=10)
            e = ctk.CTkEntry(f, font=self.main_font, width=200)
            e.pack(side="right")
            e.insert(0, self.config.get(key, ""))
            self.entries[key] = e

        ctk.CTkLabel(s, text="قراردادهای هدف:", font=self.main_font).pack(pady=10)
        self.txt_cont = ctk.CTkTextbox(s, font=self.main_font, height=80)
        self.txt_cont.pack(fill="x", padx=50)
        self.txt_cont.insert("0.0", "\n".join(self.config.get("target_contracts", [])))
        ctk.CTkButton(s, text="ذخیره تنظیمات", font=self.main_font, command=self.save_config).pack(pady=20)

        f_pw = ctk.CTkFrame(s, fg_color="#334155", corner_radius=10)
        f_pw.pack(fill="x", pady=10, padx=20)
        ctk.CTkLabel(f_pw, text="رمز مدیریت (برای تغییر کاربر و حذف کلی اطلاعات):", font=self.main_font).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(f_pw, text="تغییر رمز مدیریت", font=self.main_font, fg_color="#EA580C", command=self.change_admin_password).pack(side="right", padx=10, pady=10)

    def save_config(self):
        for k, e in self.entries.items(): self.config[k] = e.get().strip()
        self.config["target_contracts"] = [c.strip() for c in self.txt_cont.get("0.0", "end").split("\n") if c.strip()]
        self.config["enforce_sequential_payment"] = self.var_sq.get()
        self.persist_config()
        messagebox.showinfo("", "ذخیره شد.")

if __name__ == "__main__":
    app = MammutInsuranceApp()
    app.mainloop()
