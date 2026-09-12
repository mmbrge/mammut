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
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None

# تنظیمات پایه
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

MONTHS = ['فروردین','اردیبهشت','خرداد','تیر','مرداد','شهریور','مهر','آبان','آذر','دی','بهمن','اسفند']
NAV_NORMAL = "#1E2344"
NAV_ACTIVE = "#00C2A8"

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

# ================= روش‌های تقسیم اقساط =================
def split_installments_simple(premium, n):
    """محاسبات راحت: مبلغ تقسیم بر n، باقیمانده تقسیم صحیح به قسط اول اضافه می‌شود."""
    n = max(1, int(n))
    premium = int(premium)
    base = premium // n
    rem = premium - base * n
    return [base + rem if i == 0 else base for i in range(n)]

def split_installments_mammut(premium, n):
    """محاسبات ماموت: هر قسط به نزدیک‌ترین هزار به پایین گرد می‌شود و مجموع باقیمانده‌ها (۳ رقم آخر هر قسط) به قسط اول اضافه می‌شود."""
    raw = split_installments_simple(premium, n)
    tails = [r % 1000 for r in raw]
    truncated = [r - t for r, t in zip(raw, tails)]
    truncated[0] += sum(tails)
    return truncated

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

        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#141936")
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(9, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="🐘 بیمه ماموت", font=self.big_font, text_color="#00D9C0")
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 0))
        ctk.CTkLabel(self.sidebar_frame, text="سامانه حسابداری بیمه", font=self.main_font, text_color="#64749A").grid(row=1, column=0, pady=(0, 15))
        self.lbl_user = ctk.CTkLabel(self.sidebar_frame, text=f"👤 کاربر جاری: {self.current_user}", font=self.main_font, text_color="#A8B7CE")
        self.lbl_user.grid(row=2, column=0, pady=(0, 20))

        menus = [
            ("dash", "📊 داشبورد و پرداخت‌ها", self.show_dashboard),
            ("proc", "📥 وارد کردن اطلاعات", self.show_processing),
            ("manage", "🗂️ مدیریت بیمه‌نامه‌ها", self.show_policy_manager),
            ("log", "🕒 تاریخچه فعالیت‌ها", self.show_activity_log),
            ("set", "⚙️ تنظیمات سیستم", self.show_settings),
        ]
        self.nav_buttons = {}
        for i, (key, txt, cmd) in enumerate(menus, 3):
            btn = ctk.CTkButton(self.sidebar_frame, text=txt, font=self.main_font, anchor="e", command=cmd)
            btn.grid(row=i, column=0, padx=20, pady=8, sticky="ew")
            self.nav_buttons[key] = btn

        ctk.CTkButton(self.sidebar_frame, text="🔁 تغییر کاربر", font=self.main_font, anchor="e",
                      fg_color="#2A2F5C", hover_color="#64749A", command=self.change_user).grid(row=8, column=0, padx=20, pady=(20, 8), sticky="ew")
        ctk.CTkButton(self.sidebar_frame, text="🗑️ حذف کلی اطلاعات", font=self.main_font, anchor="e",
                      fg_color="#FF3B5C", hover_color="#E0294F", command=self.reset_data).grid(row=10, column=0, padx=20, pady=10, sticky="ews")

        self.main_frame = ctk.CTkFrame(self, corner_radius=15, fg_color="#1E2344")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.pending_preview_data = {}
        self.show_dashboard()

    # ================= زیرساخت (پوشه‌ها، لاگ، تنظیمات) =================
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
            "s_personnel_name": "نام پرسنل", "s_personnel_no": "شماره پرسنلی", "s_national_id": "شماره ملی",
            "b_type": "نوع بیمه گذار", "b_company": "طرف قرارداد", "b_contract": "شماره قرارداد",
            "b_date": "تاریخ صدور", "b_premium": "حق بیمه", "b_name": "بیمه گذار", "b_policy": "شماره بیمه نامه",
            "b_personnel_name": "نام پرسنل", "b_personnel_no": "شماره پرسنلی", "b_national_id": "شماره ملی",
            "target_contracts": ["1404/30000/5051"], "cutoff_day": "25", "enforce_sequential_payment": "فعال",
            "admin_password_salt": "", "admin_password_hash": "",
            "installment_calc_method": "راحت", "invoice_prefix": "22562"
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

        cursor.execute('''CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, username TEXT, action TEXT, details TEXT)''')

        cursor.execute('''CREATE TABLE IF NOT EXISTS invoice_counter (
            year TEXT PRIMARY KEY, last_seq INTEGER)''')

        cursor.execute('''CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT, invoice_no TEXT UNIQUE, inv_type TEXT, company_name TEXT,
            period TEXT, total_amount INTEGER, row_count INTEGER, created_by TEXT, created_at TEXT,
            word_path TEXT, pdf_path TEXT)''')

        for col in ["due_date", "payment_type", "description", "receipt_path", "payment_date", "updated_by"]:
            try: cursor.execute(f"ALTER TABLE installments ADD COLUMN {col} TEXT")
            except sqlite3.OperationalError: pass  # ستون از قبل وجود دارد

        for col in ["personnel_name", "personnel_no", "national_id", "extra_fields", "created_by", "created_at"]:
            try: cursor.execute(f"ALTER TABLE policies ADD COLUMN {col} TEXT")
            except sqlite3.OperationalError: pass  # ستون از قبل وجود دارد

        self.conn.commit()

    def log_action(self, action, details=""):
        """ثبت یک ردیف در تاریخچه فعالیت‌ها (audit_log) + فایل لاگ، برای رهگیری اینکه چه کسی چه‌کاری کرده."""
        ts = jdatetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        try:
            self.conn.cursor().execute("INSERT INTO audit_log (ts, username, action, details) VALUES (?, ?, ?, ?)",
                                        (ts, self.current_user, action, details))
            self.conn.commit()
        except Exception:
            logging.exception("ثبت گزارش فعالیت (audit_log) ناموفق بود")
        logging.info(f"[{self.current_user}] {action}: {details}")

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
            old_u = self.current_user
            self.current_user = new_u
            self.conn.cursor().execute("INSERT INTO sys_user (username) VALUES (?)", (new_u,))
            self.conn.commit()
            self.lbl_user.configure(text=f"👤 کاربر جاری: {self.current_user}")
            self.log_action("تغییر کاربر", f"از «{old_u}» به «{new_u}»")

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
        self.log_action("تغییر رمز مدیریت")
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
            self.log_action("حذف کلی اطلاعات", f"نسخه پشتیبان: {backup_path}")
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

    def set_active_nav(self, key):
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.configure(fg_color=NAV_ACTIVE, hover_color="#00A88F")
            else:
                btn.configure(fg_color=NAV_NORMAL, hover_color="#2A2F5C")

    def setup_treeview_style(self):
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Treeview",
                        background="#141936",
                        foreground="#F1F5F9",
                        rowheight=40,
                        fieldbackground="#141936",
                        borderwidth=0,
                        font=("Vazir", 13))
        style.map('Treeview', background=[('selected', '#3B6FE0')]) # رنگ بک‌گراند ردیف انتخاب شده (آبی تیره)
        style.configure("Treeview.Heading",
                        background="#1E2344",
                        foreground="#00D9C0",
                        font=("Vazir", 14, "bold"),
                        borderwidth=1,
                        relief="flat")
        style.map("Treeview.Heading", background=[('active', '#2A2F5C')])

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children(): widget.destroy()

    def create_glass_card(self, parent, icon, title, value, color):
        f = ctk.CTkFrame(parent, fg_color="#2A2F5C", corner_radius=15, border_width=1, border_color=color)
        f.pack(side="left", fill="both", expand=True, padx=10)
        ctk.CTkLabel(f, text=f"{icon} {title}", font=self.main_font, text_color="#CBD5E1").pack(pady=(15, 5))
        ctk.CTkLabel(f, text=value, font=self.big_font, text_color=color).pack(pady=(0, 15))

    def apply_row_stripes(self, tree):
        tree.tag_configure("stripe_even", background="#141936")
        tree.tag_configure("stripe_odd", background="#1C2142")

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

    def compute_installments(self, premium, n):
        n = max(1, int(n))
        if self.config.get("installment_calc_method", "راحت") == "ماموت":
            return split_installments_mammut(premium, n)
        return split_installments_simple(premium, n)

    def auto_resize_columns(self, tree, columns, data_rows, sample_limit=150):
        """عرض ستون‌ها را بر اساس محتوا تنظیم می‌کند؛ برای سرعت، فقط روی یک نمونه از ردیف‌ها اندازه‌گیری می‌شود، نه کل جدول."""
        font = tkFont.Font(family="Vazir", size=13)
        sample = data_rows[:sample_limit]
        for col_idx, col_name in enumerate(columns):
            max_w = font.measure(col_name) + 50
            if col_name == "انتخاب": max_w = 70
            for row in sample:
                w = font.measure(str(row[col_idx])) + 50
                if w > max_w: max_w = w
            tree.column(col_name, width=max_w, minwidth=max_w, stretch=False)

    def debounce(self, attr_name, delay_ms, func):
        """اجرای func را delay_ms میلی‌ثانیه به تعویق می‌اندازد و زمان‌بندی قبلی (در صورت وجود) را لغو می‌کند —
        برای جلوگیری از اجرای مکرر کوئری هنگام تایپ در فیلدهای جستجو."""
        existing = getattr(self, attr_name, None)
        if existing:
            try: self.after_cancel(existing)
            except Exception: pass
        setattr(self, attr_name, self.after(delay_ms, func))

    def status_badge(self, status):
        return {"پرداخت شده": "🟢 پرداخت شده", "پرداخت نشده": "🔴 پرداخت نشده",
                "پرداخت به پاسارگاد": "🟣 پرداخت به پاسارگاد"}.get(status, status)

    def make_sortable(self, tree, columns):
        """با کلیک روی هر عنوان ستون، جدول (بدون کوئری مجدد از دیتابیس) بر اساس آن ستون مرتب می‌شود."""
        state = {"col": None, "reverse": False}
        base_headers = list(columns)

        def sort_by(col):
            items = [(tree.set(k, col), k) for k in tree.get_children("")]
            reverse = state["col"] == col and not state["reverse"]

            def key_fn(pair):
                val = convert_persian_to_english_number(pair[0])
                num = re.sub(r'[^\d.\-]', '', val or "")
                try: return (0, float(num)) if num not in ("", "-", ".") else (1, val)
                except Exception: return (1, val)

            items.sort(key=key_fn, reverse=reverse)
            for idx, (_, k) in enumerate(items):
                tree.move(k, "", idx)
                tags = [t for t in tree.item(k, "tags") if not str(t).startswith("stripe_")]
                tags.append("stripe_even" if idx % 2 == 0 else "stripe_odd")
                tree.item(k, tags=tuple(tags))
            state["col"], state["reverse"] = col, reverse
            for c in base_headers: tree.heading(c, text=c)
            tree.heading(col, text=col + (" ▼" if reverse else " ▲"))

        for c in base_headers:
            tree.heading(c, command=lambda c=c: sort_by(c))

    def handle_tree_click(self, event, tree):
        """با کلیک روی هر جای ردیف (نه فقط ستون تیک)، وضعیت انتخاب (☐/☑) همان ردیف را عوض می‌کند."""
        region = tree.identify_region(event.x, event.y)
        if region == "cell":
            iid = tree.identify_row(event.y)
            if iid:
                vals = list(tree.item(iid, "values"))
                vals[0] = "☑" if vals[0] == "☐" else "☐"
                tree.item(iid, values=vals)

    def select_row_under_cursor(self, event, tree):
        """پیش از باز کردن منوی کلیک راست، ردیف زیر ماوس را (اگر جدول قابلیت انتخاب داشته باشد) هایلایت می‌کند."""
        iid = tree.identify_row(event.y)
        if iid and str(tree.cget("selectmode")) != "none":
            try:
                tree.selection_set(iid)
                tree.focus(iid)
            except Exception:
                pass
        return iid

    def handle_right_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        iid = self.select_row_under_cursor(event, tree)
        if region == "cell":
            col_id = tree.identify_column(event.x)
            col_idx = int(col_id.replace('#', '')) - 1
            if iid and col_idx >= 0:
                val = tree.item(iid, "values")[col_idx]
                menu = tk.Menu(tree, tearoff=0, font=("Vazir", 12))
                menu.add_command(label="📋 کپی مقدار این سلول", command=lambda: self.clipboard_clear() or self.clipboard_append(str(val)))
                menu.post(event.x_root, event.y_root)

    def handle_right_click_checkbox(self, event, tree):
        """کلیک راست روی جدول‌هایی که ستون تیک (☐/☑) دارند: علاوه بر کپی مقدار سلول، امکان تیک زدن/برداشتن تیک همهٔ ردیف‌های نمایش‌داده‌شده را می‌دهد."""
        region = tree.identify_region(event.x, event.y)
        iid = self.select_row_under_cursor(event, tree)
        menu = tk.Menu(tree, tearoff=0, font=("Vazir", 12))
        if region == "cell" and iid:
            col_id = tree.identify_column(event.x)
            col_idx = int(col_id.replace('#', '')) - 1
            if col_idx >= 0:
                val = tree.item(iid, "values")[col_idx]
                menu.add_command(label="📋 کپی مقدار این سلول", command=lambda: self.clipboard_clear() or self.clipboard_append(str(val)))
                menu.add_separator()
        menu.add_command(label="☑ تیک‌زدن همهٔ ردیف‌های نمایش‌داده‌شده", command=lambda: self.set_all_checks(tree, True))
        menu.add_command(label="☐ برداشتن تیک از همه", command=lambda: self.set_all_checks(tree, False))
        menu.post(event.x_root, event.y_root)

    def set_all_checks(self, tree, checked):
        mark = "☑" if checked else "☐"
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            vals[0] = mark
            tree.item(iid, values=vals)

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

    # ================= صدور صورت‌حساب (Word + PDF) =================
    def format_invoice_no(self, seq, year2):
        prefix = self.config.get("invoice_prefix", "22562")
        return f"{prefix}-{year2}-{seq:04d}"

    def get_next_invoice_no(self, cursor):
        year2 = f"{jdatetime.date.today().year % 100:02d}"
        row = cursor.execute("SELECT last_seq FROM invoice_counter WHERE year=?", (year2,)).fetchone()
        seq = (row[0] + 1) if row else 1
        if row:
            cursor.execute("UPDATE invoice_counter SET last_seq=? WHERE year=?", (seq, year2))
        else:
            cursor.execute("INSERT INTO invoice_counter (year, last_seq) VALUES (?, ?)", (year2, seq))
        return self.format_invoice_no(seq, year2)

    def fill_template_placeholders(self, doc, mapping):
        for p in doc.paragraphs:
            for key, val in mapping.items():
                if key in p.text: p.text = p.text.replace(key, val)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for key, val in mapping.items():
                            if key in p.text: p.text = p.text.replace(key, val)

    def build_and_place_table(self, doc, headers, rows, placeholder='[جدول]'):
        """جدول را می‌سازد و پر می‌کند. اگر پاراگرافی حاوی placeholder در قالب پیدا شود،
        جدول درست همان‌جا درج می‌شود (و متن placeholder حذف می‌شود)؛ در غیر این صورت
        جدول مثل قبل به انتهای سند اضافه می‌شود."""
        anchor = None
        for p in doc.paragraphs:
            if placeholder in p.text:
                anchor = p
                break

        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(headers): hdr_cells[i].text = h
        for row_vals in rows:
            cells = table.add_row().cells
            for i, v in enumerate(row_vals): cells[i].text = str(v)

        if anchor is not None:
            anchor.text = anchor.text.replace(placeholder, '').strip()
            anchor._p.addnext(table._tbl)
            if not anchor.text:
                anchor._p.getparent().remove(anchor._p)
        return table

    def convert_docx_to_pdf(self, docx_path):
        if docx2pdf_convert is None:
            logging.warning(f"docx2pdf در دسترس نیست؛ فقط فایل Word ساخته شد: {docx_path}")
            return None
        try:
            pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
            docx2pdf_convert(docx_path, pdf_path)
            return pdf_path
        except Exception:
            logging.exception(f"تبدیل PDF ناموفق بود: {docx_path}")
            return None

    def open_invoice_dialog(self):
        diag = Toplevel(self)
        diag.title("صدور صورت‌حساب")
        diag.geometry("480x520")
        diag.configure(bg="#1E2344")
        diag.transient(self)
        diag.grab_set()

        templates = [f for f in os.listdir(self.template_dir) if f.endswith(".docx")]
        if not templates:
            ctk.CTkLabel(diag, text=f"ابتدا یک قالب Word (.docx) در این پوشه قرار دهید:\n{self.template_dir}",
                         text_color="red", font=self.main_font, wraplength=420, justify="center").pack(pady=40, padx=10)
            return

        periods = [r[0] for r in self.conn.cursor().execute("SELECT DISTINCT period FROM policies ORDER BY period").fetchall()]
        if not periods:
            ctk.CTkLabel(diag, text="هیچ بیمه‌نامه‌ای ثبت نشده است.", text_color="red", font=self.main_font).pack(pady=40)
            return

        stemp = ctk.StringVar(value=templates[0])
        itype = ctk.StringVar(value="تفکیکی (یک شرکت)")
        speriod = ctk.StringVar(value=periods[-1])
        scomp = ctk.StringVar(value="")

        ctk.CTkLabel(diag, text="قالب Word:", font=self.main_font).pack(pady=(20, 5))
        ctk.CTkOptionMenu(diag, variable=stemp, values=templates, font=self.main_font).pack(pady=5)

        ctk.CTkLabel(diag, text="دوره (ماه):", font=self.main_font).pack(pady=(10, 5))

        comp_frame = ctk.CTkFrame(diag, fg_color="transparent")
        ctk.CTkLabel(comp_frame, text="شرکت:", font=self.main_font).pack(side="right", padx=5)
        cb_comp = ctk.CTkOptionMenu(comp_frame, variable=scomp, values=["-"], font=self.main_font)
        cb_comp.pack(side="right", padx=5)

        def refresh_companies(*_):
            comps = [r[0] for r in self.conn.cursor().execute(
                "SELECT DISTINCT company_name FROM policies WHERE period=? ORDER BY company_name", (speriod.get(),)).fetchall()]
            if not comps: comps = ["-"]
            cb_comp.configure(values=comps)
            scomp.set(comps[0])

        ctk.CTkOptionMenu(diag, variable=speriod, values=periods, font=self.main_font, command=refresh_companies).pack(pady=5)

        ctk.CTkLabel(diag, text="نوع صورت‌حساب:", font=self.main_font).pack(pady=(10, 5))

        def on_type_change(*_):
            if itype.get() == "تفکیکی (یک شرکت)":
                comp_frame.pack(pady=5)
            else:
                comp_frame.pack_forget()
        ctk.CTkOptionMenu(diag, variable=itype, values=["تفکیکی (یک شرکت)", "گروهی (همه شرکت‌ها)"], font=self.main_font, command=on_type_change).pack(pady=5)

        refresh_companies()
        on_type_change()

        def generate():
            try:
                comp = scomp.get() if itype.get() == "تفکیکی (یک شرکت)" else None
                result = self.generate_word_invoice(stemp.get(), itype.get(), speriod.get(), comp)
                if result:
                    diag.destroy()
                    messagebox.showinfo("موفق", f"صورت‌حساب صادر شد:\n{result}")
            except Exception as e:
                logging.exception("خطا در صدور صورت‌حساب")
                messagebox.showerror("خطا", str(e), parent=diag)
        ctk.CTkButton(diag, text="🧾 تولید فایل (Word + PDF)", font=self.title_font, fg_color="#FF5F1F", command=generate).pack(pady=30)

    def period_to_year_month(self, period):
        """رشتهٔ دوره مثل 'شهریور 1404' را به (سال, ماه) تفکیک می‌کند تا در بایگانی پوشهٔ سال/ماه ساخته شود."""
        try:
            month_name, year = period.rsplit(" ", 1)
            return year, month_name
        except Exception:
            return "نامشخص", period

    def period_dir(self, period):
        year, month = self.period_to_year_month(period)
        return os.path.join(self.archive_dir, year, month)

    def generate_word_invoice(self, template_name, inv_type, period, company=None):
        if 'docx' not in globals():
            messagebox.showerror("خطا", "نصب کتابخانه python-docx الزامیست.")
            return None
        template_path = os.path.join(self.template_dir, template_name)
        cursor = self.conn.cursor()
        today_str = jdatetime.date.today().strftime('%Y/%m/%d')
        invoice_no = self.get_next_invoice_no(cursor)

        if inv_type == "گروهی (همه شرکت‌ها)":
            data = cursor.execute("SELECT company_name, COUNT(*), SUM(premium) FROM policies WHERE period=? GROUP BY company_name", (period,)).fetchall()
            if not data:
                self.conn.rollback()
                messagebox.showwarning("هشدار", "برای این دوره بیمه‌نامه‌ای یافت نشد.")
                return None
            g_tot = sum(r[2] for r in data)
            doc = docx.Document(template_path)
            mapping = {
                '[نام_شرکت]': 'گروه صنعتی ماموت', '[دوره]': period, '[شماره_صورتحساب]': invoice_no,
                '[تاریخ_صدور]': to_persian_num(today_str), '[مبلغ_کل]': to_persian_num(f"{g_tot:,}"),
                '[تعداد_ردیف]': to_persian_num(str(len(data))),
            }
            self.fill_template_placeholders(doc, mapping)
            rows_data = [(to_persian_num(str(i)), r[0], to_persian_num(str(r[1])), to_persian_num(f"{r[2]:,}")) for i, r in enumerate(data, 1)]
            self.build_and_place_table(doc, ["ردیف", "نام شرکت", "تعداد بیمه", "جمع حق بیمه (ریال)"], rows_data)
            p_dir = self.period_dir(period)
            os.makedirs(p_dir, exist_ok=True)
            base = os.path.join(p_dir, f"صورتحساب_گروهی_{period}_{invoice_no}")
            doc.save(base + ".docx")
            pdf_path = self.convert_docx_to_pdf(base + ".docx")
            cursor.execute('''INSERT INTO invoices (invoice_no, inv_type, company_name, period, total_amount, row_count, created_by, created_at, word_path, pdf_path)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (invoice_no, "گروهی", "همه شرکت‌ها", period, g_tot, len(data), self.current_user, today_str, base + ".docx", pdf_path))
            self.conn.commit()
            self.generate_excel_reports()
            self.log_action("صدور صورتحساب گروهی", f"دوره={period}, شماره={invoice_no}, مبلغ={g_tot}")
            return base + ".docx" + (" + PDF" if pdf_path else "")
        else:
            if not company or company == "-":
                messagebox.showerror("خطا", "شرکت انتخاب نشده است.")
                return None
            pols = cursor.execute('''SELECT type, person_name, personnel_name, personnel_no, national_id, premium
                                      FROM policies WHERE company_name=? AND period=?''', (company, period)).fetchall()
            if not pols:
                self.conn.rollback()
                messagebox.showwarning("هشدار", "برای این شرکت در این دوره بیمه‌نامه‌ای یافت نشد.")
                return None
            t_prem = sum(p[5] for p in pols)
            c_doc = docx.Document(template_path)
            mapping = {
                '[نام_شرکت]': company, '[دوره]': period, '[شماره_صورتحساب]': invoice_no,
                '[تاریخ_صدور]': to_persian_num(today_str), '[مبلغ_کل]': to_persian_num(f"{t_prem:,}"),
                '[تعداد_ردیف]': to_persian_num(str(len(pols))),
            }
            self.fill_template_placeholders(c_doc, mapping)
            rows_data = [(ptype, pname, pers_name or "", pers_no or "", nid or "", to_persian_num(f"{premium:,}"))
                         for ptype, pname, pers_name, pers_no, nid, premium in pols]
            self.build_and_place_table(c_doc, ["نوع بیمه‌نامه", "نام بیمه‌گذار", "نام پرسنل", "شماره پرسنلی", "شماره ملی", "مبلغ حق بیمه"], rows_data)
            c_dir = os.path.join(self.period_dir(period), company)
            os.makedirs(c_dir, exist_ok=True)
            base = os.path.join(c_dir, f"صورتحساب_{company}_{period}_{invoice_no}")
            c_doc.save(base + ".docx")
            pdf_path = self.convert_docx_to_pdf(base + ".docx")
            cursor.execute('''INSERT INTO invoices (invoice_no, inv_type, company_name, period, total_amount, row_count, created_by, created_at, word_path, pdf_path)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (invoice_no, "تفکیکی", company, period, t_prem, len(pols), self.current_user, today_str, base + ".docx", pdf_path))
            self.conn.commit()
            self.generate_excel_reports()
            self.log_action("صدور صورتحساب تفکیکی", f"شرکت={company}, دوره={period}, شماره={invoice_no}, مبلغ={t_prem}")
            return base + ".docx" + (" + PDF" if pdf_path else "")

    # ================= داشبورد و جداول =================
    def show_dashboard(self):
        self.set_active_nav("dash")
        self.clear_main_frame()
        cursor = self.conn.cursor()

        tot_pols = cursor.execute("SELECT COUNT(*) FROM policies").fetchone()[0]
        cur_period = self.get_period(jdatetime.date.today().strftime("%Y/%m/%d"))
        cur_pols = cursor.execute("SELECT COUNT(*) FROM policies WHERE period=?", (cur_period,)).fetchone()[0]
        tot_comps = cursor.execute("SELECT COUNT(DISTINCT company_name) FROM policies").fetchone()[0]

        ctk.CTkLabel(self.main_frame, text="📊 داشبورد و پرداخت‌ها", font=self.title_font).pack(pady=(15, 5), anchor="e", padx=15)

        tf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        tf.pack(fill="x", pady=10, padx=10)
        self.create_glass_card(tf, "📄", "کل بیمه‌نامه‌ها", to_persian_num(str(tot_pols)), "#00D9C0")
        self.create_glass_card(tf, "🗓️", "بیمه‌های دوره جاری", to_persian_num(str(cur_pols)), "#00E676")
        self.create_glass_card(tf, "🏢", "شرکت‌های فعال", to_persian_num(str(tot_comps)), "#FFC107")

        bf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bf.pack(fill="x", pady=5, padx=10)
        ctk.CTkButton(bf, text="🧾 ایجاد صورت‌حساب", fg_color="#FF5F1F", font=self.main_font, command=self.open_invoice_dialog).pack(side="right", padx=5)
        ctk.CTkButton(bf, text="💳 تسویه اقساط انتخابی", fg_color="#17B978", font=self.main_font, command=self.open_payment_dialog).pack(side="right", padx=5)
        ctk.CTkButton(bf, text="🏦 واریز پاسارگاد", fg_color="#9C6BFF", font=self.main_font, command=self.mark_pasargad).pack(side="right", padx=5)
        ctk.CTkButton(bf, text="📑 اکسل پرداخت‌ها", fg_color="#29B6F6", font=self.main_font, command=lambda: os.startfile(self.payments_excel_path) if os.path.exists(self.payments_excel_path) else None).pack(side="left", padx=5)

        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=5)
        self.tabview._segmented_button.configure(font=self.main_font)

        self.tab_comp = self.tabview.add("گروهی (شرکتی)")
        self.tab_det = self.tabview.add("جزئی (تفکیکی)")

        # فیلترهای بالا
        ff = ctk.CTkFrame(self.main_frame, fg_color="#141936", corner_radius=10)
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
        self.due_start_var.trace_add("write", lambda *args: self.debounce("_deb_dash", 250, self.load_tables))
        ctk.CTkEntry(ff, textvariable=self.due_start_var, placeholder_text="1405/01/01", font=self.main_font, width=100).pack(side="right", padx=5)

        ctk.CTkLabel(ff, text="تا:", font=self.main_font).pack(side="right", padx=(5,0))
        self.due_end_var = ctk.StringVar()
        self.due_end_var.trace_add("write", lambda *args: self.debounce("_deb_dash", 250, self.load_tables))
        ctk.CTkEntry(ff, textvariable=self.due_end_var, placeholder_text="1405/12/29", font=self.main_font, width=100).pack(side="right", padx=5)

        # ---------------- پنل جستجوی مجزا بالای جدول گروهی ----------------
        sf_comp = ctk.CTkFrame(self.tab_comp, fg_color="transparent")
        sf_comp.pack(fill="x", pady=2)
        self.comp_vars = {}
        for col in ["شرکت", "دوره", "شماره قسط"]:
            v = ctk.StringVar()
            v.trace_add("write", lambda *args: self.debounce("_deb_comp", 250, self.load_comp_table))
            self.comp_vars[col] = v
            # Placeholder با نام ستون
            ctk.CTkEntry(sf_comp, textvariable=v, placeholder_text=f"جستجو {col}", font=self.main_font, width=150, text_color="white", placeholder_text_color="gray").pack(side="right", padx=5)

        self.cols_comp = ("انتخاب", "شرکت", "دوره", "شماره قسط", "سررسید", "وضعیت", "مبلغ کل (ریال)")
        tree_f_comp = ctk.CTkFrame(self.tab_comp)
        tree_f_comp.pack(fill="both", expand=True)
        self.tree_comp = ttk.Treeview(tree_f_comp, columns=self.cols_comp, show="headings", selectmode="none")
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
            v.trace_add("write", lambda *args: self.debounce("_deb_det", 250, self.load_det_table))
            self.det_vars[col] = v
            # Placeholder با نام ستون
            ctk.CTkEntry(sf_det, textvariable=v, placeholder_text=f"جستجو {col}", font=self.main_font, width=180, text_color="white", placeholder_text_color="gray").pack(side="right", padx=5)

        self.cols_det = ("انتخاب", "شناسه", "شرکت", "دوره", "بیمه‌گذار", "شماره بیمه", "قسط", "سررسید", "مبلغ (ریال)", "وضعیت")
        tree_f_det = ctk.CTkFrame(self.tab_det)
        tree_f_det.pack(fill="both", expand=True)
        self.tree_det = ttk.Treeview(tree_f_det, columns=self.cols_det, show="headings", selectmode="none")
        vs_det = ttk.Scrollbar(tree_f_det, orient="vertical", command=self.tree_det.yview)
        hs_det = ttk.Scrollbar(tree_f_det, orient="horizontal", command=self.tree_det.xview)
        self.tree_det.configure(yscrollcommand=vs_det.set, xscrollcommand=hs_det.set)
        hs_det.pack(side="bottom", fill="x")
        vs_det.pack(side="right", fill="y")
        self.tree_det.pack(side="left", fill="both", expand=True)

        for t in [self.tree_comp, self.tree_det]:
            # رنگ‌های جدید و روشن‌تر برای خوانایی بهتر روی پس زمینه تیره
            t.tag_configure("unpaid", foreground="#FF5C77") # قرمز روشن
            t.tag_configure("paid", foreground="#00E676") # سبز روشن
            t.tag_configure("pasargad", foreground="#C792FF") # بنفش
            self.apply_row_stripes(t)
            t.bind("<ButtonRelease-1>", lambda e, tree=t: self.handle_tree_click(e, tree))
            t.bind("<Button-3>", lambda e, tree=t: self.handle_right_click_checkbox(e, tree))

        self.tree_comp.bind("<Double-1>", self.show_selected_comp_details)

        for c in self.cols_comp: self.tree_comp.heading(c, text=c)
        for c in self.cols_det: self.tree_det.heading(c, text=c)
        self.make_sortable(self.tree_comp, self.cols_comp)
        self.make_sortable(self.tree_det, self.cols_det)

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

            vals = ("☐", comp, per, to_persian_num(str(inum)), to_persian_num(due_f), self.status_badge(stat), to_persian_num(f"{amt:,}"))
            tag = "paid" if stat == "پرداخت شده" else "pasargad" if stat == "پرداخت به پاسارگاد" else "unpaid"
            stripe = "stripe_even" if len(final_rows) % 2 == 0 else "stripe_odd"
            iid = f"{comp}|{per}|{inum}|{stat}"
            self.tree_comp.insert("", "end", iid=iid, values=vals, tags=(tag, stripe))
            final_rows.append(vals)

        self.auto_resize_columns(self.tree_comp, self.cols_comp, final_rows)

    def show_selected_comp_details(self, event=None):
        iid = self.tree_comp.identify_row(event.y) if event is not None else None
        if not iid:
            selected = self.tree_comp.selection()
            iid = selected[0] if selected else None
        if not iid: return messagebox.showwarning("هشدار", "یک ردیف را دابل کلیک کنید.")
        comp, period, i_num, status = iid.split('|')

        diag = Toplevel(self)
        diag.title(f"ریز اقساط: {comp} - قسط {to_persian_num(str(i_num))}")
        diag.geometry("900x450")
        diag.configure(bg="#1E2344")
        diag.transient(self)

        cols = ("بیمه‌گذار", "شماره بیمه", "سررسید", "مبلغ (ریال)", "وضعیت")
        ts = ttk.Treeview(diag, columns=cols, show="headings")
        for c in cols: ts.heading(c, text=c)
        ts.tag_configure("unpaid", foreground="#FF5C77")
        ts.tag_configure("paid", foreground="#00E676")
        ts.tag_configure("pasargad", foreground="#C792FF")
        self.apply_row_stripes(ts)

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
            v = (r[0], r[1], due, to_persian_num(f"{r[2]:,}"), self.status_badge(r[3]))
            tag = "paid" if r[3] == "پرداخت شده" else "pasargad" if r[3] == "پرداخت به پاسارگاد" else "unpaid"
            stripe = "stripe_even" if len(data_rows) % 2 == 0 else "stripe_odd"
            ts.insert("", "end", values=v, tags=(tag, stripe))
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

            vals = ("☐", _id, comp, per, name, pol, to_persian_num(str(inum)), to_persian_num(due_f), to_persian_num(f"{amt:,}"), self.status_badge(stat))
            tag = "paid" if stat == "پرداخت شده" else "pasargad" if stat == "پرداخت به پاسارگاد" else "unpaid"
            stripe = "stripe_even" if len(final_rows) % 2 == 0 else "stripe_odd"
            self.tree_det.insert("", "end", iid=str(_id), values=vals, tags=(tag, stripe))
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
        diag.geometry("560x680")
        diag.configure(bg="#1E2344")
        diag.transient(self)
        diag.grab_set()

        amt_card = ctk.CTkFrame(diag, fg_color="#2A2F5C", corner_radius=15, border_width=1, border_color="#17B978")
        amt_card.pack(fill="x", padx=20, pady=(20, 10))
        lbl_amt_title = ctk.CTkLabel(amt_card, text="💰 مبلغ کل انتخابی", font=self.main_font, text_color="#CBD5E1")
        lbl_amt_title.pack(pady=(12, 2))
        lbl_amt_value = ctk.CTkLabel(amt_card, text=f"{to_persian_num(f'{tot_amt:,}')} ریال", font=self.big_font, text_color="#00E676")
        lbl_amt_value.pack(pady=(0, 12))

        mode_card = ctk.CTkFrame(diag, fg_color="#141936", corner_radius=12)
        mode_card.pack(fill="x", padx=20, pady=8)
        pay_mode = ctk.StringVar(value="full")

        def refresh_amount_display():
            if pay_mode.get() == "partial":
                entered = clean_number(ent_amt.get())
                remaining = tot_amt - entered
                lbl_amt_title.configure(text="🧮 مبلغ باقیمانده پس از این پرداخت")
                if entered <= 0:
                    lbl_amt_value.configure(text=f"{to_persian_num(f'{tot_amt:,}')} ریال", text_color="#FFC107")
                elif remaining <= 0:
                    lbl_amt_value.configure(text="مبلغ از کل بیشتر است!", text_color="#FF5C77")
                else:
                    lbl_amt_value.configure(text=f"{to_persian_num(f'{remaining:,}')} ریال", text_color="#FFC107")
            else:
                lbl_amt_title.configure(text="💰 مبلغ کل انتخابی")
                lbl_amt_value.configure(text=f"{to_persian_num(f'{tot_amt:,}')} ریال", text_color="#00E676")

        def toggle_ent():
            ent_amt.configure(state="normal" if pay_mode.get() == "partial" else "disabled")
            refresh_amount_display()

        frb = ctk.CTkFrame(mode_card, fg_color="transparent")
        frb.pack(fill="x", padx=10, pady=(10, 5))
        rb_partial = ctk.CTkRadioButton(frb, text="مبلغ دلخواه (ثبت کسری)", variable=pay_mode, value="partial", font=self.main_font, command=toggle_ent)
        ctk.CTkRadioButton(frb, text="✅ تسویه کامل", variable=pay_mode, value="full", font=self.main_font, command=toggle_ent).pack(side="right", padx=10)
        rb_partial.pack(side="right", padx=10)

        ent_amt = ctk.CTkEntry(mode_card, font=self.main_font, placeholder_text="مبلغ پرداختی (ریال)", state="disabled")
        ent_amt.pack(pady=(0, 12), fill="x", padx=20)

        def on_amt_key(event=None):
            raw = ent_amt.get()
            digits = clean_number(raw)
            formatted = f"{digits:,}" if digits else ""
            if formatted != raw:
                ent_amt.delete(0, "end")
                ent_amt.insert(0, formatted)
                try: ent_amt.icursor("end")
                except Exception: pass
            refresh_amount_display()
        ent_amt.bind("<KeyRelease>", on_amt_key)

        if active == "گروهی (شرکتی)":
            ctk.CTkLabel(mode_card, text="در پرداخت ناقصِ گروهی، مبلغ به‌ترتیب شناسه روی اقساط زیرمجموعه اعمال می‌شود؛ باقیمانده به‌صورت قسط جدید و پرداخت‌نشده باقی می‌ماند.",
                         font=self.main_font, text_color="#A8B7CE", wraplength=470, justify="right").pack(padx=20, pady=(0, 12))

        details_card = ctk.CTkFrame(diag, fg_color="#141936", corner_radius=12)
        details_card.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(details_card, text="نوع پرداخت:", font=self.main_font).pack(pady=(12, 0))
        p_type = ctk.StringVar(value="فیش بانکی")
        cb_pt = ctk.CTkOptionMenu(details_card, variable=p_type, values=["فیش بانکی", "نقدی", "چک", "کسر از حقوق"], font=self.main_font)
        cb_pt.pack(pady=5)

        ctk.CTkLabel(details_card, text="توضیحات:", font=self.main_font).pack(pady=(10, 0))
        desc = ctk.CTkEntry(details_card, font=self.main_font, width=300)
        desc.pack(pady=(5, 12))

        self.cur_receipts = [] # پشتیبانی از چند فیش
        def sel_file():
            paths = filedialog.askopenfilenames()
            if paths:
                self.cur_receipts.extend(paths)
                lbl_file.configure(text=f"📎 {len(self.cur_receipts)} فایل انتخاب شد", text_color="#00E676")
        receipt_card = ctk.CTkFrame(diag, fg_color="#141936", corner_radius=12)
        receipt_card.pack(fill="x", padx=20, pady=8)
        ctk.CTkButton(receipt_card, text="📎 انتخاب تصویر فیش/رسید", font=self.main_font, fg_color="#29B6F6", command=sel_file).pack(pady=(12, 5))
        lbl_file = ctk.CTkLabel(receipt_card, text="سندی انتخاب نشده", font=self.main_font, text_color="gray")
        lbl_file.pack(pady=(0, 12))

        def commit():
            amt = tot_amt
            if pay_mode.get() == "partial":
                amt = clean_number(ent_amt.get())
                if amt <= 0 or amt >= tot_amt: return messagebox.showerror("خطا", "مبلغ نامعتبر است.", parent=diag)

            saved_receipt_paths = []
            if self.cur_receipts:
                sample_id = selected_ids_to_pay[0]
                i_num_str = cursor.execute("SELECT inst_num FROM installments WHERE id=?", (sample_id,)).fetchone()[0]
                comp = cursor.execute("SELECT p.company_name FROM installments i JOIN policies p ON i.policy_id=p.id WHERE i.id=?", (sample_id,)).fetchone()[0]
                per = cursor.execute("SELECT p.period FROM installments i JOIN policies p ON i.policy_id=p.id WHERE i.id=?", (sample_id,)).fetchone()[0]

                r_dir = os.path.join(self.period_dir(per), comp, "Receipts")
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
                cursor.execute(f"UPDATE installments SET status='پرداخت شده', payment_type=?, description=?, receipt_path=?, payment_date=?, updated_by=? WHERE id IN ({placeholders})",
                               (p_type.get(), desc.get(), final_receipt_str, now_d, self.current_user, *selected_ids_to_pay))
            else:
                cursor.execute(f"SELECT id, policy_id, inst_num, amount FROM installments WHERE id IN ({placeholders}) ORDER BY id", selected_ids_to_pay)
                for r_id, p_id, i_n, i_a in cursor.fetchall():
                    if amt >= i_a:
                        cursor.execute("UPDATE installments SET status='پرداخت شده', payment_type=?, description=?, receipt_path=?, payment_date=?, updated_by=? WHERE id=?", (p_type.get(), desc.get(), final_receipt_str, now_d, self.current_user, r_id))
                        amt -= i_a
                    elif amt > 0:
                        cursor.execute("UPDATE installments SET amount=?, status='پرداخت شده', payment_type=?, description=?, receipt_path=?, payment_date=?, updated_by=? WHERE id=?", (amt, p_type.get(), desc.get(), final_receipt_str, now_d, self.current_user, r_id))
                        rem = i_a - amt
                        cursor.execute("INSERT INTO installments (policy_id, inst_num, amount, status) VALUES (?, ?, ?, 'پرداخت نشده')", (p_id, i_n, rem))
                        amt = 0
                    else: break
            self.conn.commit()
            self.log_action("ثبت پرداخت", f"{len(selected_ids_to_pay)} قسط، حالت={pay_mode.get()}، نوع پرداخت={p_type.get()}")
            self.generate_excel_reports()
            self.load_tables()
            diag.destroy()
            messagebox.showinfo("موفق", "پرداختی ثبت شد.")

        ctk.CTkButton(diag, text="✅ تایید نهایی", fg_color="#17B978", font=self.title_font, command=commit).pack(pady=20)

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
        cursor.execute(f"UPDATE installments SET status='پرداخت به پاسارگاد', updated_by=? WHERE id IN ({placeholders})", [self.current_user] + real_ids)
        self.conn.commit()
        self.generate_excel_reports()
        self.load_tables()
        self.log_action("انتقال به پاسارگاد", f"{len(real_ids)} قسط")
        messagebox.showinfo("موفق", "تغییر وضعیت انجام شد.")

    # ================= خروجی اکسل =================
    def generate_excel_reports(self):
        cursor = self.conn.cursor()
        data = cursor.execute('''SELECT p.company_name, p.period, p.person_name, p.policy_no,
                          i.inst_num, i.amount, i.due_date, i.status, i.payment_type, i.description, i.payment_date, i.receipt_path,
                          p.personnel_name, p.personnel_no, p.national_id, i.updated_by, p.type
                          FROM installments i JOIN policies p ON i.policy_id = p.id ORDER BY p.company_name, p.id, i.inst_num''').fetchall()
        if not data: return
        cols = ["شرکت", "دوره", "بیمه‌گذار", "شماره بیمه", "قسط", "مبلغ قسط", "سررسید", "وضعیت", "نوع پرداخت", "توضیحات",
                "تاریخ پرداخت", "فیش", "نام پرسنل", "شماره پرسنلی", "شماره ملی", "ثبت‌کننده تسویه", "نوع بیمه"]
        df = pd.DataFrame(data, columns=cols)
        df['کاربر ثبت‌کننده'] = self.current_user
        df.to_excel(self.payments_excel_path, index=False)
        self.apply_b_nazanin(self.payments_excel_path)

        comp_periods = df[['شرکت', 'دوره']].drop_duplicates().values.tolist()
        for comp, period in comp_periods:
            c_dir = os.path.join(self.period_dir(period), comp)
            os.makedirs(c_dir, exist_ok=True)
            self.build_company_workbook(cursor, comp, period, os.path.join(c_dir, f"گزارش_مالی_{comp}.xlsx"))

        for period in df['دوره'].unique():
            p_dir = self.period_dir(period)
            os.makedirs(p_dir, exist_ok=True)
            self.build_period_summary_workbook(cursor, period, os.path.join(p_dir, "گزارش_کلی_دوره.xlsx"))

    def build_company_workbook(self, cursor, comp, period, path):
        """یک فایل اکسل با سه شیت برای هر شرکت/دوره می‌سازد: بیمه‌نامه‌ها، مالی جزئی، مالی کلی شرکت."""
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "بیمه‌نامه‌ها"
        ws1.append(["ردیف", "نوع", "بیمه‌گذار", "نام پرسنل", "شماره پرسنلی", "شماره ملی", "شماره بیمه‌نامه", "تاریخ صدور", "حق بیمه", "وضعیت کلی"])

        ws2 = wb.create_sheet("مالی جزئی")
        ws2.append(["ردیف", "بیمه‌گذار", "شماره بیمه", "قسط", "مبلغ قسط", "سررسید", "وضعیت", "مبلغ پرداخت‌شده", "نوع پرداخت", "فیش", "ثبت‌کننده"])

        pols = cursor.execute('''SELECT id, policy_no, person_name, personnel_name, personnel_no, national_id,
                                  issue_date, premium, type FROM policies WHERE company_name=? AND period=?''',
                               (comp, period)).fetchall()

        tot_premium = tot_paid = tot_unpaid = tot_pasargad = 0
        row_i = 0
        for idx, pol in enumerate(pols, 1):
            pid, pno, pname, pers_name, pers_no, nid, issue_date, premium, ptype = pol
            insts = cursor.execute("SELECT inst_num, amount, due_date, status, receipt_path, updated_by, payment_type FROM installments WHERE policy_id=? ORDER BY inst_num", (pid,)).fetchall()
            overall = "تسویه کامل" if insts and all(i[3] in ("پرداخت شده", "پرداخت به پاسارگاد") for i in insts) else "دارای معوق"
            ws1.append([idx, ptype, pname, pers_name or "", pers_no or "", nid or "", pno, issue_date, premium, overall])
            for inst in insts:
                row_i += 1
                i_num, amt, due, status, receipt, upd_by, p_type = inst
                paid_amt = amt if status in ("پرداخت شده", "پرداخت به پاسارگاد") else 0
                ws2.append([row_i, pname, pno, i_num, amt, due or "", status, paid_amt, p_type or "", receipt or "", upd_by or ""])
                tot_premium += amt
                if status == "پرداخت شده": tot_paid += amt
                elif status == "پرداخت به پاسارگاد": tot_pasargad += amt
                else: tot_unpaid += amt

        ws3 = wb.create_sheet("مالی کلی شرکت")
        ws3.append(["شرح", "تعداد بیمه‌نامه", "مبلغ (ریال)"])
        ws3.append(["جمع حق بیمه", len(pols), tot_premium])
        ws3.append(["جمع پرداخت‌شده", "", tot_paid])
        ws3.append(["جمع پرداخت‌نشده (معوق)", "", tot_unpaid])
        ws3.append(["جمع واریزی به پاسارگاد", "", tot_pasargad])

        inv = cursor.execute('''SELECT word_path, pdf_path, invoice_no FROM invoices
                                 WHERE company_name=? AND period=? AND inv_type='تفکیکی' ORDER BY id DESC LIMIT 1''', (comp, period)).fetchone()
        if inv:
            word_path, pdf_path, invoice_no = inv
            ws3.append([f"لینک صورت‌حساب ({invoice_no})", "", ""])
            link_cell = ws3.cell(row=ws3.max_row, column=1)
            target = pdf_path or word_path
            if target and os.path.exists(target):
                link_cell.hyperlink = "file:///" + target.replace("\\", "/")
                link_cell.font = Font(name='B Nazanin', color="0000FF", underline="single")

        for ws in (ws1, ws2, ws3):
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and not (cell.hyperlink):
                        cell.font = Font(name='B Nazanin', bold=(cell.row == 1))
        wb.save(path)

    def build_period_summary_workbook(self, cursor, period, path):
        """در پوشهٔ هر دوره، یک اکسل خلاصه از همهٔ شرکت‌های آن دوره (تعداد بیمه‌نامه و وضعیت پرداخت) با لینک به
        پروندهٔ هر شرکت و صورت‌حسابش می‌سازد."""
        wb = Workbook()
        ws = wb.active
        ws.title = "خلاصه شرکت‌ها"
        ws.append(["ردیف", "نام شرکت", "تعداد بیمه‌نامه", "جمع حق بیمه", "جمع پرداخت‌شده", "جمع پرداخت‌نشده",
                    "جمع پاسارگاد", "پروندهٔ شرکت", "صورت‌حساب"])

        comps = cursor.execute("SELECT DISTINCT company_name FROM policies WHERE period=? ORDER BY company_name", (period,)).fetchall()
        p_dir = self.period_dir(period)
        for idx, (comp,) in enumerate(comps, 1):
            pol_count = cursor.execute("SELECT COUNT(*) FROM policies WHERE company_name=? AND period=?", (comp, period)).fetchone()[0]
            rows = cursor.execute('''SELECT i.amount, i.status FROM installments i JOIN policies p ON i.policy_id=p.id
                                      WHERE p.company_name=? AND p.period=?''', (comp, period)).fetchall()
            tot = sum(a for a, s in rows)
            paid = sum(a for a, s in rows if s == "پرداخت شده")
            unpaid = sum(a for a, s in rows if s == "پرداخت نشده")
            pasargad = sum(a for a, s in rows if s == "پرداخت به پاسارگاد")
            ws.append([idx, comp, pol_count, tot, paid, unpaid, pasargad, "پرونده شرکت", ""])

            comp_file = os.path.join(p_dir, comp, f"گزارش_مالی_{comp}.xlsx")
            if os.path.exists(comp_file):
                cell = ws.cell(row=ws.max_row, column=8)
                cell.hyperlink = "file:///" + comp_file.replace("\\", "/")
                cell.font = Font(name='B Nazanin', color="0000FF", underline="single")

            inv = cursor.execute('''SELECT word_path, pdf_path, invoice_no FROM invoices
                                     WHERE company_name=? AND period=? AND inv_type='تفکیکی' ORDER BY id DESC LIMIT 1''', (comp, period)).fetchone()
            if inv:
                word_path, pdf_path, invoice_no = inv
                target = pdf_path or word_path
                if target and os.path.exists(target):
                    cell2 = ws.cell(row=ws.max_row, column=9, value=invoice_no)
                    cell2.hyperlink = "file:///" + target.replace("\\", "/")
                    cell2.font = Font(name='B Nazanin', color="0000FF", underline="single")

        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and not cell.hyperlink:
                    cell.font = Font(name='B Nazanin', bold=(cell.row == 1))
        wb.save(path)

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

    # ================= وارد کردن اطلاعات (پردازش هوشمند) =================
    def show_processing(self):
        self.set_active_nav("proc")
        self.clear_main_frame()
        self.b_path = self.s_path = None

        ctk.CTkLabel(self.main_frame, text="📥 وارد کردن اطلاعات و پردازش هوشمند", font=self.title_font).pack(pady=20)

        for ftype, txt in [('b', "📄 فایل بدنه"), ('s', "📄 فایل ثالث")]:
            f = ctk.CTkFrame(self.main_frame, fg_color="transparent")
            f.pack(pady=5)
            ctk.CTkButton(f, text=txt, font=self.main_font, command=lambda t=ftype: self.sel_file(t)).pack(side="right", padx=10)
            lbl = ctk.CTkLabel(f, text="انتخاب نشده", font=self.main_font)
            lbl.pack(side="right")
            if ftype == 'b': self.lbl_b_file = lbl
            else: self.lbl_s_file = lbl

        opt_f = ctk.CTkFrame(self.main_frame, fg_color="#141936", corner_radius=10)
        opt_f.pack(pady=10)
        ctk.CTkLabel(opt_f, text="تعداد اقساط برای هر بیمه‌نامه:", font=self.main_font).pack(side="right", padx=10, pady=10)
        self.ent_inst_count = ctk.CTkEntry(opt_f, font=self.main_font, width=80)
        self.ent_inst_count.insert(0, "9")
        self.ent_inst_count.pack(side="right", padx=10, pady=10)
        ctk.CTkLabel(opt_f, text=f"(روش محاسبه فعلی: {self.config.get('installment_calc_method','راحت')} — قابل تغییر از تنظیمات)",
                     font=self.main_font, text_color="#A8B7CE").pack(side="right", padx=10)

        bf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bf.pack(pady=15)
        ctk.CTkButton(bf, text="⚡ پیش‌پردازش فایل‌ها", font=self.title_font, fg_color="#FF5F1F", command=self.run_preview).pack(side="right", padx=10)
        ctk.CTkButton(bf, text="➕ افزودن دستی", font=self.title_font, fg_color="#29B6F6", command=self.add_manual_row).pack(side="left", padx=10)

        self.preview_scroll = ctk.CTkScrollableFrame(self.main_frame, fg_color="#141936")
        self.preview_scroll.pack(fill="both", expand=True, padx=20, pady=10)
        self.btn_commit = ctk.CTkButton(self.main_frame, text="✅ پردازش نهایی", font=self.title_font, fg_color="#17B978", command=self.commit_processing, state="disabled")
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
        diag.geometry("450x650")
        diag.configure(bg="#1E2344")
        diag.transient(self)

        fields = [("نام شرکت:", "company"), ("بیمه‌گذار:", "name"), ("شماره بیمه:", "policy_no"), ("حق بیمه:", "premium"), ("تاریخ صدور:", "date"),
                  ("نام پرسنل:", "personnel_name"), ("شماره پرسنلی:", "personnel_no"), ("شماره ملی:", "national_id")]
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
                "premium": clean_number(ents["premium"].get()), "date": clean_text_advanced(ents["date"].get()), "type": t_var.get(),
                "personnel_name": clean_text_advanced(ents["personnel_name"].get()),
                "personnel_no": clean_text_advanced(ents["personnel_no"].get()),
                "national_id": clean_text_advanced(ents["national_id"].get()),
                "extra": {}
            })
            self.render_preview_ui()
            diag.destroy()
        ctk.CTkButton(diag, text="ثبت", fg_color="#17B978", font=self.main_font, command=save).pack(pady=20)

    def edit_preview_company(self, comp):
        diag = Toplevel(self)
        diag.title(f"ویرایش: {comp}")
        diag.geometry("1100x600")
        diag.configure(bg="#1E2344")
        diag.transient(self)

        top_bar = ctk.CTkFrame(diag, fg_color="transparent")
        top_bar.pack(fill="x", padx=10, pady=(10, 0))

        def add_field():
            fname = simpledialog.askstring("فیلد جدید", "نام فیلد جدید را وارد کنید:", parent=diag)
            if not fname: return
            for item in self.pending_preview_data[comp]:
                item.setdefault('extra', {})
                item['extra'].setdefault(fname, "")
            diag.destroy()
            self.edit_preview_company(comp)
        ctk.CTkButton(top_bar, text="➕ افزودن فیلد جدید", fg_color="#29B6F6", font=self.main_font, command=add_field).pack(side="left", padx=5)

        sf = ctk.CTkScrollableFrame(diag, fg_color="#141936")
        sf.pack(fill="both", expand=True, padx=10, pady=10)

        extra_keys = sorted({k for item in self.pending_preview_data[comp] for k in item.get('extra', {}).keys()})
        header = ctk.CTkFrame(sf, fg_color="transparent")
        header.pack(fill="x", pady=(0, 5))
        for lbl in reversed(["نوع", "بیمه‌گذار", "شماره بیمه", "حق بیمه", "تاریخ", "نام پرسنل", "شماره پرسنلی", "شماره ملی"] + extra_keys):
            ctk.CTkLabel(header, text=lbl, font=self.main_font, width=110, text_color="#00D9C0").pack(side="right", padx=2)

        elist = []
        for item in self.pending_preview_data[comp]:
            item.setdefault('extra', {})
            for k in extra_keys: item['extra'].setdefault(k, "")
            f = ctk.CTkFrame(sf, fg_color="transparent")
            f.pack(fill="x", pady=3)

            et = ctk.CTkOptionMenu(f, values=["ثالث", "بدنه"], width=100, font=self.main_font)
            et.set(item.get('type', 'ثالث')); et.pack(side="right", padx=2)
            en = ctk.CTkEntry(f, font=self.main_font, width=110); en.insert(0, item['name']); en.pack(side="right", padx=2)
            ep = ctk.CTkEntry(f, font=self.main_font, width=110); ep.insert(0, item['policy_no']); ep.pack(side="right", padx=2)
            epr = ctk.CTkEntry(f, font=self.main_font, width=110); epr.insert(0, str(item['premium'])); epr.pack(side="right", padx=2)
            ed = ctk.CTkEntry(f, font=self.main_font, width=100); ed.insert(0, item['date']); ed.pack(side="right", padx=2)
            epn = ctk.CTkEntry(f, font=self.main_font, width=110); epn.insert(0, item.get('personnel_name', '')); epn.pack(side="right", padx=2)
            epno = ctk.CTkEntry(f, font=self.main_font, width=110); epno.insert(0, item.get('personnel_no', '')); epno.pack(side="right", padx=2)
            enid = ctk.CTkEntry(f, font=self.main_font, width=110); enid.insert(0, item.get('national_id', '')); enid.pack(side="right", padx=2)
            extra_entries = {}
            for k in extra_keys:
                ex = ctk.CTkEntry(f, font=self.main_font, width=110); ex.insert(0, item['extra'].get(k, "")); ex.pack(side="right", padx=2)
                extra_entries[k] = ex
            elist.append((item, et, en, ep, epr, ed, epn, epno, enid, extra_entries))

        def save_e():
            for item, et, en, ep, epr, ed, epn, epno, enid, extra_entries in elist:
                item['type'] = et.get()
                item['name'] = clean_text_advanced(en.get())
                item['policy_no'] = clean_text_advanced(ep.get())
                item['premium'] = clean_number(epr.get())
                item['date'] = clean_text_advanced(ed.get())
                item['personnel_name'] = clean_text_advanced(epn.get())
                item['personnel_no'] = clean_text_advanced(epno.get())
                item['national_id'] = clean_text_advanced(enid.get())
                for k, ex in extra_entries.items():
                    item['extra'][k] = ex.get().strip()
            self.render_preview_ui()
            diag.destroy()
        ctk.CTkButton(diag, text="ذخیره", fg_color="#17B978", font=self.main_font, command=save_e).pack(pady=10)

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
            f = ctk.CTkFrame(self.preview_scroll, fg_color="#2A2F5C")
            f.pack(fill="x", pady=5, padx=5)
            t_prem = sum(i['premium'] for i in items)
            ctk.CTkLabel(f, text=f"{to_persian_num(comp)} | {to_persian_num(str(len(items)))} مورد | {to_persian_num(f'{t_prem:,}')} ریال", font=self.main_font, text_color="#00E676").pack(side="right", padx=10, pady=10)
            ctk.CTkButton(f, text="حذف", fg_color="#FF3B5C", width=60, font=self.main_font, command=lambda c=comp: [self.pending_preview_data.pop(c), self.render_preview_ui()]).pack(side="left", padx=5)
            ctk.CTkButton(f, text="ویرایش", fg_color="#29B6F6", width=60, font=self.main_font, command=lambda c=comp: self.edit_preview_company(c)).pack(side="left", padx=5)

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
                cpn = resolve_column(df, self.config.get(f"{pfx}_personnel_name"), "personnel_name")
                cpno = resolve_column(df, self.config.get(f"{pfx}_personnel_no"), "personnel_no")
                cnid = resolve_column(df, self.config.get(f"{pfx}_national_id"), "national_id")
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
                        "premium": clean_number(row.get(cpr, 0)), "date": clean_text_advanced(row.get(cd, "")), "type": t_name,
                        "personnel_name": clean_text_advanced(row.get(cpn, "")) if cpn else "",
                        "personnel_no": clean_text_advanced(row.get(cpno, "")) if cpno else "",
                        "national_id": clean_text_advanced(row.get(cnid, "")) if cnid else "",
                        "extra": {}
                    })
            except Exception as e:
                logging.exception(f"run_preview: خطا در پردازش فایل {path}")
                messagebox.showerror("خطا", str(e))
        self.render_preview_ui()

    def commit_processing(self):
        cursor = self.conn.cursor()
        try:
            n = int(convert_persian_to_english_number(self.ent_inst_count.get()).strip() or "9")
            if n < 1: n = 9
        except Exception:
            n = 9
        now_d = jdatetime.date.today().strftime('%Y/%m/%d')
        inserted = 0
        for comp, items in self.pending_preview_data.items():
            for it in items:
                per = self.get_period(it['date'])
                extra_json = json.dumps(it.get('extra', {}), ensure_ascii=False)
                cursor.execute('''INSERT OR IGNORE INTO policies
                    (policy_no, person_name, company_name, premium, issue_date, period, type,
                     personnel_name, personnel_no, national_id, extra_fields, created_by, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                    (it['policy_no'], it['name'], comp, it['premium'], it['date'], per, it['type'],
                     it.get('personnel_name', ''), it.get('personnel_no', ''), it.get('national_id', ''),
                     extra_json, self.current_user, now_d))
                if cursor.rowcount > 0:
                    inserted += 1
                    pid = cursor.lastrowid
                    shares = self.compute_installments(it['premium'], n)
                    for i, amt in enumerate(shares, 1):
                        due = self.get_due_date(per, i)
                        cursor.execute("INSERT INTO installments (policy_id, inst_num, amount, due_date) VALUES (?, ?, ?, ?)", (pid, i, amt, due))
        self.conn.commit()
        method = self.config.get('installment_calc_method', 'راحت')
        self.log_action("پردازش هوشمند", f"{inserted} بیمه‌نامه جدید، {n} قسط، روش={method}")
        self.generate_excel_reports()
        messagebox.showinfo("موفقیت", f"پردازش انجام شد. {inserted} بیمه‌نامه جدید ثبت شد.")
        self.pending_preview_data = {}
        self.show_dashboard()

    # ================= تاریخچه فعالیت‌ها (Audit Log) =================
    def show_activity_log(self):
        self.set_active_nav("log")
        self.clear_main_frame()
        ctk.CTkLabel(self.main_frame, text="🕒 تاریخچه فعالیت‌ها", font=self.title_font).pack(pady=(15, 5), anchor="e", padx=15)
        ctk.CTkLabel(self.main_frame, text="برای رهگیری اینکه چه کسی در برنامه چه عملیاتی انجام داده است.",
                     font=self.main_font, text_color="#A8B7CE").pack(anchor="e", padx=15, pady=(0, 10))

        ff = ctk.CTkFrame(self.main_frame, fg_color="#141936", corner_radius=10)
        ff.pack(fill="x", padx=10, pady=5)
        search_var = ctk.StringVar()
        ctk.CTkLabel(ff, text="جستجوی کاربر/عملیات:", font=self.main_font).pack(side="right", padx=5, pady=5)
        ctk.CTkEntry(ff, textvariable=search_var, font=self.main_font, width=250).pack(side="right", padx=5)

        cols = ("تاریخ و زمان", "کاربر", "عملیات", "جزئیات")
        tree_f = ctk.CTkFrame(self.main_frame)
        tree_f.pack(fill="both", expand=True, padx=10, pady=5)
        tree = ttk.Treeview(tree_f, columns=cols, show="headings", selectmode="browse")
        for c in cols: tree.heading(c, text=c)
        self.make_sortable(tree, cols)
        tree.tag_configure("stripe_even", background="#141936")
        tree.tag_configure("stripe_odd", background="#1C2142")
        vs = ttk.Scrollbar(tree_f, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        tree.bind("<Button-3>", lambda e: self.handle_right_click(e, tree))

        def reload(*_):
            for item in tree.get_children(): tree.delete(item)
            rows = self.conn.cursor().execute("SELECT ts, username, action, details FROM audit_log ORDER BY id DESC LIMIT 500").fetchall()
            q = search_var.get().strip().lower()
            data_rows = []
            for r in rows:
                if q and q not in (r[1] or "").lower() and q not in (r[2] or "").lower(): continue
                stripe = "stripe_even" if len(data_rows) % 2 == 0 else "stripe_odd"
                tree.insert("", "end", values=r, tags=(stripe,))
                data_rows.append(r)
            self.auto_resize_columns(tree, cols, data_rows)

        search_var.trace_add("write", lambda *args: self.debounce("_deb_log", 250, reload))
        reload()

    # ================= مدیریت بیمه‌نامه‌ها (ویرایش/حذف پس از ثبت) =================
    def show_policy_manager(self):
        self.set_active_nav("manage")
        self.clear_main_frame()
        ctk.CTkLabel(self.main_frame, text="🗂️ مدیریت بیمه‌نامه‌ها", font=self.title_font).pack(pady=(15, 5), anchor="e", padx=15)
        ctk.CTkLabel(self.main_frame, text="ویرایش مشخصات یا حذف یک بیمه‌نامهٔ ثبت‌شده (برای اصلاح اشتباهات ورود اطلاعات).",
                     font=self.main_font, text_color="#A8B7CE").pack(anchor="e", padx=15, pady=(0, 10))

        sf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        sf.pack(fill="x", pady=2, padx=10)
        pm_vars = {}
        for col in ["شرکت", "بیمه‌گذار", "شماره بیمه", "دوره"]:
            v = ctk.StringVar()
            v.trace_add("write", lambda *args: self.debounce("_deb_pm", 250, reload))
            pm_vars[col] = v
            ctk.CTkEntry(sf, textvariable=v, placeholder_text=f"جستجو {col}", font=self.main_font, width=170).pack(side="right", padx=5)

        cols = ("شناسه", "نوع", "شرکت", "دوره", "بیمه‌گذار", "نام پرسنل", "شماره بیمه", "حق بیمه", "وضعیت اقساط")
        tree_f = ctk.CTkFrame(self.main_frame)
        tree_f.pack(fill="both", expand=True, padx=10, pady=5)
        tree = ttk.Treeview(tree_f, columns=cols, show="headings", selectmode="browse")
        for c in cols: tree.heading(c, text=c)
        self.make_sortable(tree, cols)
        tree.tag_configure("stripe_even", background="#141936")
        tree.tag_configure("stripe_odd", background="#1C2142")
        tree.tag_configure("has_paid", foreground="#00E676")
        tree.tag_configure("clean", foreground="#F1F5F9")
        vs = ttk.Scrollbar(tree_f, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        tree.bind("<Button-3>", lambda e: self.handle_right_click(e, tree))

        def reload(*_):
            for item in tree.get_children(): tree.delete(item)
            rows = self.conn.cursor().execute('''SELECT id, type, company_name, period, person_name, personnel_name, policy_no, premium
                                                  FROM policies ORDER BY company_name, period, id DESC LIMIT 1000''').fetchall()
            s_comp = pm_vars["شرکت"].get().strip().lower()
            s_name = pm_vars["بیمه‌گذار"].get().strip().lower()
            s_pol = pm_vars["شماره بیمه"].get().strip().lower()
            s_per = pm_vars["دوره"].get().strip().lower()
            data_rows = []
            for r in rows:
                pid, ptype, comp, per, name, pers_name, pno, premium = r
                if s_comp and s_comp not in str(comp).lower(): continue
                if s_name and s_name not in str(name).lower(): continue
                if s_pol and s_pol not in str(pno).lower(): continue
                if s_per and s_per not in str(per).lower(): continue
                statuses = [x[0] for x in self.conn.cursor().execute("SELECT status FROM installments WHERE policy_id=?", (pid,)).fetchall()]
                if not statuses: st_label = "بدون قسط"
                elif all(s == "پرداخت نشده" for s in statuses): st_label = "بدون پرداخت"
                elif all(s in ("پرداخت شده", "پرداخت به پاسارگاد") for s in statuses): st_label = "تسویه کامل"
                else: st_label = "دارای پرداخت جزئی"
                tag2 = "clean" if st_label in ("بدون پرداخت", "بدون قسط") else "has_paid"
                st_display = {"بدون قسط": "⬜ بدون قسط", "بدون پرداخت": "🔴 بدون پرداخت",
                              "تسویه کامل": "🟢 تسویه کامل", "دارای پرداخت جزئی": "🟡 دارای پرداخت جزئی"}[st_label]
                stripe = "stripe_even" if len(data_rows) % 2 == 0 else "stripe_odd"
                vals = (pid, ptype, comp, per, name, pers_name or "", pno, to_persian_num(f"{premium:,}"), st_display)
                tree.insert("", "end", iid=str(pid), values=vals, tags=(tag2, stripe))
                data_rows.append(vals)
            self.auto_resize_columns(tree, cols, data_rows)

        def open_edit(event=None):
            sel = tree.selection()
            if not sel: return messagebox.showwarning("هشدار", "یک ردیف را انتخاب کنید.")
            self.edit_policy_dialog(int(sel[0]))
        tree.bind("<Double-1>", open_edit)

        bf = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        bf.pack(pady=10)
        ctk.CTkButton(bf, text="✏️ ویرایش بیمه‌نامه انتخابی", fg_color="#29B6F6", font=self.main_font, command=open_edit).pack(side="right", padx=5)

        def do_delete():
            sel = tree.selection()
            if not sel: return messagebox.showwarning("هشدار", "یک ردیف را انتخاب کنید.")
            self.delete_policy(int(sel[0]))
            reload()
        ctk.CTkButton(bf, text="🗑️ حذف بیمه‌نامه انتخابی", fg_color="#FF3B5C", hover_color="#E0294F", font=self.main_font, command=do_delete).pack(side="right", padx=5)

        reload()

    def edit_policy_dialog(self, policy_id):
        cursor = self.conn.cursor()
        row = cursor.execute('''SELECT policy_no, person_name, company_name, premium, issue_date, period, type,
                                 personnel_name, personnel_no, national_id FROM policies WHERE id=?''', (policy_id,)).fetchone()
        if not row:
            return messagebox.showerror("خطا", "بیمه‌نامه یافت نشد.")
        (policy_no, person_name, company_name, premium, issue_date, period, ptype,
         personnel_name, personnel_no, national_id) = row

        statuses = [x[0] for x in cursor.execute("SELECT status FROM installments WHERE policy_id=?", (policy_id,)).fetchall()]
        has_payment = any(s != "پرداخت نشده" for s in statuses)

        diag = Toplevel(self)
        diag.title(f"ویرایش بیمه‌نامه {policy_no}")
        diag.geometry("500x700")
        diag.configure(bg="#1E2344")
        diag.transient(self)
        diag.grab_set()

        fields = [("شماره بیمه‌نامه:", "policy_no", policy_no), ("نام شرکت:", "company_name", company_name),
                  ("بیمه‌گذار:", "person_name", person_name), ("نام پرسنل:", "personnel_name", personnel_name or ""),
                  ("شماره پرسنلی:", "personnel_no", personnel_no or ""), ("شماره ملی:", "national_id", national_id or ""),
                  ("تاریخ صدور:", "issue_date", issue_date), ("حق بیمه:", "premium", str(premium))]
        ents = {}
        for lbl, key, val in fields:
            f = ctk.CTkFrame(diag, fg_color="transparent")
            f.pack(fill="x", padx=20, pady=4)
            ctk.CTkLabel(f, text=lbl, font=self.main_font, width=130, anchor="e").pack(side="right", padx=10)
            e = ctk.CTkEntry(f, font=self.main_font)
            e.insert(0, val)
            e.pack(side="right", fill="x", expand=True)
            ents[key] = e

        ft = ctk.CTkFrame(diag, fg_color="transparent")
        ft.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(ft, text="نوع بیمه:", font=self.main_font, width=130, anchor="e").pack(side="right", padx=10)
        t_var = ctk.StringVar(value=ptype)
        ctk.CTkOptionMenu(ft, variable=t_var, values=["ثالث", "بدنه"], font=self.main_font).pack(side="right", fill="x", expand=True)

        if has_payment:
            ctk.CTkLabel(diag, text="⚠️ این بیمه‌نامه دارای قسط پرداخت‌شده/واریزی است؛ تغییر حق بیمه به‌تنهایی اقساط را بازسازی نمی‌کند.",
                         font=self.main_font, text_color="#FFC107", wraplength=440, justify="right").pack(padx=20, pady=(10, 0))

        def save():
            new_pno = clean_text_advanced(ents["policy_no"].get())
            new_company = clean_text_advanced(ents["company_name"].get())
            if not new_pno or not new_company:
                return messagebox.showerror("خطا", "شماره بیمه و نام شرکت الزامی است.", parent=diag)
            dup = cursor.execute("SELECT id FROM policies WHERE policy_no=? AND id != ?", (new_pno, policy_id)).fetchone()
            if dup:
                return messagebox.showerror("خطا", "این شماره بیمه‌نامه قبلاً برای رکورد دیگری ثبت شده است.", parent=diag)
            new_premium = clean_number(ents["premium"].get())
            cursor.execute('''UPDATE policies SET policy_no=?, company_name=?, person_name=?, personnel_name=?,
                               personnel_no=?, national_id=?, issue_date=?, premium=?, type=? WHERE id=?''',
                           (new_pno, new_company, clean_text_advanced(ents["person_name"].get()),
                            clean_text_advanced(ents["personnel_name"].get()), clean_text_advanced(ents["personnel_no"].get()),
                            clean_text_advanced(ents["national_id"].get()), clean_text_advanced(ents["issue_date"].get()),
                            new_premium, t_var.get(), policy_id))
            self.conn.commit()
            self.log_action("ویرایش بیمه‌نامه", f"شناسه={policy_id}, شماره={new_pno}, حق‌بیمه جدید={new_premium}")
            self.generate_excel_reports()
            diag.destroy()
            self.show_policy_manager()
            messagebox.showinfo("موفق", "بیمه‌نامه ویرایش شد.")

        def rebuild():
            if has_payment:
                return messagebox.showerror("خطا", "چون این بیمه‌نامه دارای قسط پرداخت‌شده است، اقساط قابل بازسازی خودکار نیستند.", parent=diag)
            n_str = simpledialog.askstring("بازسازی اقساط", "تعداد اقساط جدید را وارد کنید:", initialvalue="9", parent=diag)
            if not n_str: return
            try:
                n = int(convert_persian_to_english_number(n_str).strip())
                if n < 1: raise ValueError
            except Exception:
                return messagebox.showerror("خطا", "تعداد قسط نامعتبر است.", parent=diag)
            new_premium = clean_number(ents["premium"].get())
            per_new = self.get_period(clean_text_advanced(ents["issue_date"].get()))
            cursor.execute("DELETE FROM installments WHERE policy_id=?", (policy_id,))
            shares = self.compute_installments(new_premium, n)
            for i, amt in enumerate(shares, 1):
                due = self.get_due_date(per_new, i)
                cursor.execute("INSERT INTO installments (policy_id, inst_num, amount, due_date) VALUES (?, ?, ?, ?)", (policy_id, i, amt, due))
            cursor.execute("UPDATE policies SET premium=?, period=? WHERE id=?", (new_premium, per_new, policy_id))
            self.conn.commit()
            self.log_action("بازسازی اقساط", f"شناسه={policy_id}, تعداد قسط جدید={n}, حق‌بیمه={new_premium}")
            self.generate_excel_reports()
            messagebox.showinfo("موفق", "اقساط بازسازی شد.", parent=diag)
            diag.destroy()
            self.show_policy_manager()

        bf = ctk.CTkFrame(diag, fg_color="transparent")
        bf.pack(pady=20)
        ctk.CTkButton(bf, text="💾 ذخیره تغییرات", fg_color="#17B978", font=self.main_font, command=save).pack(side="right", padx=5)
        rebuild_btn = ctk.CTkButton(bf, text="🔁 بازسازی اقساط", fg_color="#9C6BFF", font=self.main_font, command=rebuild)
        rebuild_btn.pack(side="right", padx=5)
        if has_payment:
            rebuild_btn.configure(state="disabled")

    def delete_policy(self, policy_id):
        cursor = self.conn.cursor()
        row = cursor.execute("SELECT policy_no, company_name, period FROM policies WHERE id=?", (policy_id,)).fetchone()
        if not row:
            return messagebox.showerror("خطا", "بیمه‌نامه یافت نشد.")
        policy_no, company_name, period = row
        statuses = [x[0] for x in cursor.execute("SELECT status FROM installments WHERE policy_id=?", (policy_id,)).fetchall()]
        has_payment = any(s != "پرداخت نشده" for s in statuses)

        if has_payment:
            if not self.verify_admin(f"بیمه‌نامه «{policy_no}» دارای قسط پرداخت‌شده/واریزی است.\nبرای حذف، رمز مدیریت را وارد کنید:"):
                return
            if not messagebox.askyesno("اخطار", f"بیمه‌نامه «{policy_no}» و همهٔ سوابق پرداخت آن حذف می‌شود.\nاین عمل غیرقابل بازگشت است. ادامه می‌دهید؟"):
                return
        else:
            if not messagebox.askyesno("تایید حذف", f"بیمه‌نامه «{policy_no}» حذف شود؟"):
                return

        cursor.execute("DELETE FROM installments WHERE policy_id=?", (policy_id,))
        cursor.execute("DELETE FROM policies WHERE id=?", (policy_id,))
        self.conn.commit()
        self.log_action("حذف بیمه‌نامه", f"شماره={policy_no}, شرکت={company_name}, دوره={period}, دارای‌پرداخت={has_payment}")
        self.generate_excel_reports()
        messagebox.showinfo("موفق", "بیمه‌نامه حذف شد.")

    # ================= تنظیمات =================
    def show_settings(self):
        self.set_active_nav("set")
        self.clear_main_frame()
        s = ctk.CTkScrollableFrame(self.main_frame, fg_color="transparent")
        s.pack(fill="both", expand=True)
        ctk.CTkLabel(s, text="⚙️ تنظیمات سیستم", font=self.title_font).pack(pady=10)

        # بخش الزام تسویه پیوسته
        f_sq = ctk.CTkFrame(s, fg_color="#2A2F5C", corner_radius=10)
        f_sq.pack(fill="x", pady=10, padx=20)
        ctk.CTkLabel(f_sq, text="قانون تسویه پیوسته اقساط:", font=self.main_font).pack(side="right", padx=10, pady=10)
        self.var_sq = ctk.StringVar(value=self.config.get("enforce_sequential_payment", "فعال"))
        cb_sq = ctk.CTkOptionMenu(f_sq, variable=self.var_sq, values=["فعال", "غیرفعال"], font=self.main_font)
        cb_sq.pack(side="right", padx=10, pady=10)

        # بخش روش محاسبه اقساط
        f_calc = ctk.CTkFrame(s, fg_color="#2A2F5C", corner_radius=10)
        f_calc.pack(fill="x", pady=10, padx=20)
        ctk.CTkLabel(f_calc, text="روش محاسبه اقساط:", font=self.main_font).pack(side="right", padx=10, pady=10)
        self.var_calc = ctk.StringVar(value=self.config.get("installment_calc_method", "راحت"))
        ctk.CTkOptionMenu(f_calc, variable=self.var_calc, values=["راحت", "ماموت"], font=self.main_font).pack(side="right", padx=10, pady=10)
        ctk.CTkLabel(s, text="راحت: باقیمانده تقسیم به قسط اول اضافه می‌شود.  |  ماموت: هر قسط به هزار تومان گرد و باقیمانده‌ها به قسط اول اضافه می‌شود.",
                     font=self.main_font, text_color="#A8B7CE", wraplength=900, justify="right").pack(padx=20, pady=(0, 10), anchor="e")

        self.entries = {}
        fields = [
            ("بیمه‌گذار ثالث:", "s_type"), ("طرف قرارداد ثالث:", "s_company"), ("قرارداد ثالث:", "s_contract"),
            ("تاریخ صدور ثالث:", "s_date"), ("حق بیمه ثالث:", "s_premium"), ("نام ثالث:", "s_name"), ("شماره ثالث:", "s_policy"),
            ("نام پرسنل ثالث:", "s_personnel_name"), ("شماره پرسنلی ثالث:", "s_personnel_no"), ("شماره ملی ثالث:", "s_national_id"),
            ("بیمه‌گذار بدنه:", "b_type"), ("طرف قرارداد بدنه:", "b_company"), ("قرارداد بدنه:", "b_contract"),
            ("تاریخ صدور بدنه:", "b_date"), ("حق بیمه بدنه:", "b_premium"), ("نام بدنه:", "b_name"), ("شماره بدنه:", "b_policy"),
            ("نام پرسنل بدنه:", "b_personnel_name"), ("شماره پرسنلی بدنه:", "b_personnel_no"), ("شماره ملی بدنه:", "b_national_id"),
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
        ctk.CTkButton(s, text="💾 ذخیره تنظیمات", font=self.main_font, fg_color="#17B978", command=self.save_config).pack(pady=20)

        f_pw = ctk.CTkFrame(s, fg_color="#2A2F5C", corner_radius=10)
        f_pw.pack(fill="x", pady=10, padx=20)
        ctk.CTkLabel(f_pw, text="رمز مدیریت (برای تغییر کاربر و حذف کلی اطلاعات):", font=self.main_font).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(f_pw, text="🔑 تغییر رمز مدیریت", font=self.main_font, fg_color="#FF5F1F", command=self.change_admin_password).pack(side="right", padx=10, pady=10)

        tmpl_card = ctk.CTkFrame(s, fg_color="#2A2F5C", corner_radius=12)
        tmpl_card.pack(fill="x", padx=20, pady=(20, 20))
        ctk.CTkLabel(tmpl_card, text="📄 راهنمای کدهای قالب Word صورت‌حساب", font=self.title_font, text_color="#00D9C0").pack(pady=(15, 2))
        ctk.CTkLabel(tmpl_card,
                     text="این کدها را هرجای متن فایل Word خودتان (پوشهٔ Data/Template) بنویسید؛ برنامه هنگام صدور صورت‌حساب آن‌ها را با مقدار واقعی جایگزین می‌کند:",
                     font=self.main_font, text_color="#A8B7CE", wraplength=850, justify="right").pack(padx=15, pady=(0, 12))

        tmpl_codes = [
            ("[نام_شرکت]", "نام شرکت (در صورت‌حساب گروهی: «گروه صنعتی ماموت»)"),
            ("[دوره]", "دورهٔ انتخاب‌شده، مثل «شهریور 1404»"),
            ("[شماره_صورتحساب]", "شمارهٔ ترتیبی صورت‌حساب، مثل «22562-05-0001»"),
            ("[تاریخ_صدور]", "تاریخ امروز به شمسی"),
            ("[مبلغ_کل]", "جمع مبلغ صورت‌حساب (ریال)"),
            ("[تعداد_ردیف]", "تعداد ردیف‌های جدول (تعداد بیمه‌نامه در نوع تفکیکی، یا تعداد شرکت‌ها در نوع گروهی)"),
            ("[جدول]", "محل دقیق درج جدول ردیف‌ها. اگر این کد را ننویسید، جدول به‌صورت خودکار به انتهای سند اضافه می‌شود."),
        ]
        for code, desc in tmpl_codes:
            row = ctk.CTkFrame(tmpl_card, fg_color="#1E2344", corner_radius=8)
            row.pack(fill="x", padx=20, pady=3)
            ctk.CTkLabel(row, text=code, font=ctk.CTkFont(family="Consolas", size=13, weight="bold"),
                         text_color="#00D9C0", width=180, anchor="e").pack(side="right", padx=(5, 10), pady=8)
            ctk.CTkLabel(row, text=desc, font=self.main_font, text_color="#F1F5F9", anchor="e",
                         wraplength=580, justify="right").pack(side="right", fill="x", expand=True, padx=(10, 5), pady=8)

        ctk.CTkLabel(tmpl_card,
                     text="نیازی به ساختن جدول در قالب نیست؛ برنامه خودش جدول را با ستون‌های مناسبِ همان نوع صورت‌حساب می‌سازد.",
                     font=self.main_font, text_color="#A8B7CE", wraplength=850, justify="right").pack(padx=15, pady=(8, 15))

    def save_config(self):
        for k, e in self.entries.items(): self.config[k] = e.get().strip()
        self.config["target_contracts"] = [c.strip() for c in self.txt_cont.get("0.0", "end").split("\n") if c.strip()]
        self.config["enforce_sequential_payment"] = self.var_sq.get()
        self.config["installment_calc_method"] = self.var_calc.get()
        self.persist_config()
        self.log_action("ذخیره تنظیمات")
        messagebox.showinfo("", "ذخیره شد.")

if __name__ == "__main__":
    app = MammutInsuranceApp()
    app.mainloop()
