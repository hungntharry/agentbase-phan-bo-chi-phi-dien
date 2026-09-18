# -*- coding: utf-8 -*-
"""
Desktop App: Update Excel with Payment Data
Input sources: PDF, Excel, Image (JPG/PNG/GIF), API
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pdfplumber
import openpyxl
import re
import os
import sys
import io
import shutil
import subprocess
import threading
import json
from datetime import datetime
import calendar

# Optional imports
try:
    import pytesseract
    from PIL import Image
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# ==========================================
# CONFIG
# ==========================================
DEFAULT_DIR = r"D:\File storage\AI\AI Agent\Điện"
DATA_START_ROW = 5
COL_CODE = 8
COL_KY_TT = 13
COL_KY_TT_TRUOC = 14
COL_TONG_TIEN_DIEM = 15
EXCEL_PREFIX = "Bảng phân bổ thanh toán chi phí điện"

# Tesseract path (common install locations)
TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]

# ==========================================
# DATA EXTRACTION FUNCTIONS
# ==========================================

def add_months(date_str, n):
    """Add n months to a dd/mm/yyyy date string (n can be negative)"""
    try:
        d = datetime.strptime(date_str.strip(), '%d/%m/%Y')
        total_months = d.year * 12 + d.month - 1 + n
        year = total_months // 12
        month = total_months % 12 + 1
        last_day = calendar.monthrange(year, month)[1]
        day = min(d.day, last_day)
        return d.replace(year=year, month=month, day=day).strftime('%d/%m/%Y')
    except:
        return None

def shift_date_range(date_range_str, n):
    """Shift a date range by n months. Input: '01/07/2026 - 31/07/2026'"""
    if not date_range_str:
        return None
    match = re.match(r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', str(date_range_str).strip())
    if not match:
        return None
    start_new = add_months(match.group(1), n)
    end_new = add_months(match.group(2), n)
    if start_new and end_new:
        return f"{start_new} - {end_new}"
    return None

def extract_pdf_data(filepath):
    """Extract from PDF using pdfplumber"""
    data = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    stt = str(row[0]).strip() if row[0] else ""
                    if not stt.isdigit():
                        continue
                    desc = str(row[5]).strip() if row[5] else ""
                    debit = str(row[6]).strip() if row[6] else ""
                    credit = str(row[7]).strip() if row[7] else ""

                    code = _extract_code(desc)
                    amount = _extract_amount(debit, credit)

                    if code and code != 'MSB' and amount > 0:
                        data.append({'code': code, 'amount': amount})
    return data

def extract_excel_data(filepath):
    """Extract from Excel input file (col 1=code, col 3=amount)"""
    data = []
    wb = openpyxl.load_workbook(filepath, keep_vba=False, data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=5, values_only=True):
            if not row or not row[0]:
                continue
            code = str(row[0]).strip()
            amount = 0
            if len(row) > 2 and row[2]:
                try:
                    amount = float(str(row[2]).replace(',', '').replace('.00', '').strip())
                except:
                    pass
            if code and amount > 0:
                data.append({'code': code, 'amount': amount})
    wb.close()
    return data

def extract_image_data(filepath):
    """Extract from image using OCR (Tesseract)"""
    if not HAS_OCR:
        raise Exception("pytesseract chưa cài. Chạy: pip install pytesseract Pillow")

    tess_path = None
    for p in TESSERACT_PATHS:
        if os.path.exists(p):
            tess_path = p
            break
    if not tess_path:
        raise Exception("Tesseract OCR chưa cài.\nDownload: https://github.com/UB-Mannheim/tesseract/releases")

    pytesseract.pytesseract.tesseract_cmd = tess_path

    img = Image.open(filepath)
    text = pytesseract.image_to_string(img, lang='eng')

    data = []
    for line in text.split('\n'):
        code = _extract_code(line)
        amount = _extract_amount_from_text(line)
        if code and code != 'MSB' and amount > 0:
            data.append({'code': code, 'amount': amount})
    return data

def extract_api_data(api_url, api_headers=None):
    """Fetch data from API and extract code + amount"""
    if not HAS_REQUESTS:
        raise Exception("requests chưa cài. Chạy: pip install requests")

    resp = requests.get(api_url, headers=api_headers, timeout=30)
    resp.raise_for_status()
    result = resp.json()

    data = []
    items = result if isinstance(result, list) else result.get('data', result.get('items', []))
    for item in items:
        code = item.get('code') or item.get('ma_khach_hang') or item.get('ma') or ''
        amount = item.get('amount') or item.get('so_tien') or item.get('tien') or 0
        if code and amount:
            data.append({'code': str(code).strip(), 'amount': float(amount)})
    return data

def _extract_code(text):
    """Extract payment code from text"""
    if not text:
        return ""
    text = str(text).strip()
    m = re.match(r'^([A-Z]{2}[0-9A-Z]+)', text)
    if not m:
        m = re.search(r'\b([A-Z]{2}\d{8,})\b', text)
    return m.group(1) if m else ""

def _extract_amount(debit, credit):
    """Extract amount from debit/credit strings"""
    amount_str = ""
    if debit and debit != '0':
        amount_str = debit.replace(',', '').replace('.00', '').strip()
    elif credit and credit != '0':
        amount_str = credit.replace(',', '').replace('.00', '').strip()
    return float(amount_str) if amount_str else 0

def _extract_amount_from_text(text):
    """Extract last number from text line as amount"""
    nums = re.findall(r'[\d,]+\.?\d*', str(text))
    if nums:
        try:
            return float(nums[-1].replace(',', ''))
        except:
            return 0
    return 0

# ==========================================
# CORE FUNCTIONS (unchanged)
# ==========================================

def find_base_dir():
    if not os.path.exists(DEFAULT_DIR):
        return DEFAULT_DIR
    for item in os.listdir(DEFAULT_DIR):
        if item.startswith('Đ') and 'i' in item.lower():
            return os.path.join(DEFAULT_DIR, item)
    return DEFAULT_DIR

def find_excel_file(base_dir):
    for f in os.listdir(base_dir):
        if f.startswith('~$') or 'backup' in f.lower():
            continue
        if f.endswith('.xlsx') and 'gốc' in f.lower():
            return os.path.join(base_dir, f)
    for f in os.listdir(base_dir):
        if f.startswith('~$') or 'backup' in f.lower():
            continue
        if f.endswith('.xlsx'):
            return os.path.join(base_dir, f)
    return None

def build_lookup(data_list):
    lookup = {}
    for item in data_list:
        if item['code'] not in lookup:
            lookup[item['code']] = item
    return lookup

def process_excel(excel_path, lookup, month_shift=0):
    wb = openpyxl.load_workbook(excel_path, keep_vba=False)
    ws = wb.worksheets[1]
    results = []
    updated = skipped_di_doi = skipped_no_match = 0

    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        excel_code = ws.cell(row=row_idx, column=COL_CODE).value
        if not excel_code:
            continue
        excel_code = str(excel_code).strip()
        current_date = ws.cell(row=row_idx, column=COL_KY_TT).value or ""
        ky_truoc = ws.cell(row=row_idx, column=COL_KY_TT_TRUOC).value or ""

        if any(kw in str(current_date).lower() for kw in ['di d', 'dời', 'ngừng']) or \
           any(kw in str(ky_truoc).lower() for kw in ['di d', 'dời', 'ngừng']):
            ws.cell(row=row_idx, column=COL_TONG_TIEN_DIEM).value = None
            skipped_di_doi += 1
            results.append({'row': row_idx, 'code': excel_code, 'action': 'SKIP (di dời)', 'date': '---', 'amount': 'blank'})
            continue

        # Shift col 13 and col 14 by month_shift
        if month_shift != 0:
            new_13 = shift_date_range(current_date, month_shift)
            if new_13:
                ws.cell(row=row_idx, column=COL_KY_TT).value = new_13
            new_14 = shift_date_range(ky_truoc, month_shift)
            if new_14:
                ws.cell(row=row_idx, column=COL_KY_TT_TRUOC).value = new_14

        if excel_code in lookup:
            ws.cell(row=row_idx, column=COL_TONG_TIEN_DIEM).value = lookup[excel_code]['amount']
            updated += 1
            results.append({'row': row_idx, 'code': excel_code, 'action': 'UPDATED', 'date': '(giữ nguyên)', 'amount': f"{lookup[excel_code]['amount']:,.0f}"})
        else:
            skipped_no_match += 1
            results.append({'row': row_idx, 'code': excel_code, 'action': 'NO MATCH', 'date': '---', 'amount': '---'})

    return wb, results, updated, skipped_di_doi, skipped_no_match

def get_output_filename(input_paths, base_dir):
    basename = os.path.basename(input_paths[0])
    month_match = re.search(r'(?:tháng|thang)\s*(\d+\.?\d*)', basename, re.IGNORECASE)
    if month_match:
        month_str = month_match.group(1)
        return os.path.join(base_dir, f"{EXCEL_PREFIX} tháng {month_str}.xlsx"), month_str
    return None, None

# ==========================================
# GUI APP
# ==========================================

class ExcelUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Agent - Phân bổ chi phí điện")
        self.root.geometry("950x700")
        self.root.configure(padx=15, pady=10)

        self.input_paths = []
        self.api_url = ""
        self.excel_path = None
        self.wb = None
        self.results = []
        self.output_path = None

        self._build_ui()

    def _build_ui(self):
        ttk.Label(self.root, text="AI Agent - Phân bổ chi phí điện",
                  font=('Segoe UI', 16, 'bold')).pack(pady=(0, 10))

        # === 1. INPUT FILES ===
        input_frame = ttk.LabelFrame(self.root, text="1. Chọn file đầu vào (PDF / Excel / Image)", padding=10)
        input_frame.pack(fill='x', pady=(0, 8))

        btn_frame = ttk.Frame(input_frame)
        btn_frame.pack(fill='x')

        ttk.Button(btn_frame, text="Chọn file", command=self._select_files).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text="Chọn nhiều file", command=self._select_multiple_files).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text="Xóa danh sách", command=self._clear_inputs).pack(side='left')

        self.input_listbox = tk.Listbox(input_frame, height=3, selectmode=tk.EXTENDED)
        self.input_listbox.pack(fill='x', pady=(8, 0))

        # === 1b. API ===
        api_frame = ttk.LabelFrame(self.root, text="Hoặc nhập API URL", padding=10)
        api_frame.pack(fill='x', pady=(0, 8))

        api_row = ttk.Frame(api_frame)
        api_row.pack(fill='x')
        self.api_entry = ttk.Entry(api_row, width=70)
        self.api_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(api_row, text="Test API", command=self._test_api).pack(side='left', padx=(5, 0))

        # === 2. EXCEL TEMPLATE ===
        excel_frame = ttk.LabelFrame(self.root, text="2. File Excel template (gốc)", padding=10)
        excel_frame.pack(fill='x', pady=(0, 8))

        excel_btn = ttk.Frame(excel_frame)
        excel_btn.pack(fill='x')
        ttk.Button(excel_btn, text="Tự động tìm", command=self._auto_find_excel).pack(side='left', padx=(0, 5))
        ttk.Button(excel_btn, text="Chọn file", command=self._select_excel).pack(side='left')
        self.excel_label = ttk.Label(excel_frame, text="Chưa chọn", foreground='gray')
        self.excel_label.pack(anchor='w', pady=(5, 0))

        # === 3. PROCESS ===
        proc_frame = ttk.Frame(self.root)
        proc_frame.pack(fill='x', pady=(0, 8))
        self.process_btn = ttk.Button(proc_frame, text="3. Xử lý dữ liệu", command=self._process, state='disabled')
        self.process_btn.pack(side='left')
        self.auto_open_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(proc_frame, text="Tự mở Excel sau khi lưu", variable=self.auto_open_var).pack(side='left', padx=20)

        # === 4. RESULTS ===
        res_frame = ttk.LabelFrame(self.root, text="4. Kết quả", padding=10)
        res_frame.pack(fill='both', expand=True, pady=(0, 8))

        self.summary_label = ttk.Label(res_frame, text="Chưa xử lý", foreground='gray')
        self.summary_label.pack(anchor='w', pady=(0, 5))

        cols = ('row', 'code', 'action', 'date', 'amount')
        self.tree = ttk.Treeview(res_frame, columns=cols, show='headings', height=12)
        for c, t, w, a in [('row','Dòng',60,'center'), ('code','Mã khách hàng',160,'w'),
                            ('action','Hành động',120,'center'), ('date','Ngày',100,'center'),
                            ('amount','Số tiền',140,'e')]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a)
        self.tree.tag_configure('updated', foreground='green')
        self.tree.tag_configure('skip', foreground='orange')
        self.tree.tag_configure('nomatch', foreground='red')

        sb = ttk.Scrollbar(res_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # === 5. SAVE ===
        bottom = ttk.Frame(self.root)
        bottom.pack(fill='x')
        self.save_btn = ttk.Button(bottom, text="5. Lưu file Excel", command=self._save, state='disabled')
        self.save_btn.pack(side='left')
        self.status_label = ttk.Label(bottom, text="", foreground='gray')
        self.status_label.pack(side='left', padx=20)

    def _select_files(self):
        path = filedialog.askopenfilename(
            title="Chọn file đầu vào",
            initialdir=find_base_dir(),
            filetypes=[("All supported", "*.pdf *.xlsx *.xls *.jpg *.jpeg *.png *.gif"),
                       ("PDF", "*.pdf"), ("Excel", "*.xlsx *.xls"),
                       ("Image", "*.jpg *.jpeg *.png *.gif")]
        )
        if path:
            self.input_paths = [path]
            self._refresh_input_list()
            self._check_ready()

    def _select_multiple_files(self):
        paths = filedialog.askopenfilenames(
            title="Chọn nhiều file",
            initialdir=find_base_dir(),
            filetypes=[("All supported", "*.pdf *.xlsx *.xls *.jpg *.jpeg *.png *.gif"),
                       ("PDF", "*.pdf"), ("Excel", "*.xlsx *.xls"),
                       ("Image", "*.jpg *.jpeg *.png *.gif")]
        )
        if paths:
            self.input_paths = list(paths)
            self._refresh_input_list()
            self._check_ready()

    def _clear_inputs(self):
        self.input_paths = []
        self._refresh_input_list()
        self._check_ready()

    def _refresh_input_list(self):
        self.input_listbox.delete(0, tk.END)
        for p in self.input_paths:
            ext = os.path.splitext(p)[1].lower()
            self.input_listbox.insert(tk.END, f"[{ext[1:].upper()}] {os.path.basename(p)}")

    def _test_api(self):
        url = self.api_entry.get().strip()
        if not url:
            messagebox.showwarning("Cảnh báo", "Nhập API URL trước!")
            return
        try:
            data = extract_api_data(url)
            messagebox.showinfo("API OK", f"Tải được {len(data)} bản ghi từ API")
        except Exception as e:
            messagebox.showerror("Lỗi API", str(e))

    def _auto_find_excel(self):
        excel = find_excel_file(find_base_dir())
        if excel:
            self.excel_path = excel
            self.excel_label.config(text=f"Đã chọn: {os.path.basename(excel)}", foreground='green')
            self._check_ready()
        else:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file Excel gốc!")

    def _select_excel(self):
        path = filedialog.askopenfilename(
            title="Chọn file Excel gốc", initialdir=find_base_dir(),
            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.excel_path = path
            self.excel_label.config(text=f"Đã chọn: {os.path.basename(path)}", foreground='green')
            self._check_ready()

    def _check_ready(self):
        has_input = len(self.input_paths) > 0 or self.api_entry.get().strip()
        ready = has_input and self.excel_path is not None
        self.process_btn.config(state='normal' if ready else 'disabled')

    def _process(self):
        self.process_btn.config(state='disabled')
        self.save_btn.config(state='disabled')
        self.status_label.config(text="Đang xử lý...", foreground='blue')
        self.root.update()

        try:
            all_data = []

            # Process files
            for path in self.input_paths:
                ext = os.path.splitext(path)[1].lower()
                if ext == '.pdf':
                    all_data.extend(extract_pdf_data(path))
                elif ext in ('.xlsx', '.xls'):
                    all_data.extend(extract_excel_data(path))
                elif ext in ('.jpg', '.jpeg', '.png', '.gif'):
                    all_data.extend(extract_image_data(path))

            # Process API
            api_url = self.api_entry.get().strip()
            if api_url:
                all_data.extend(extract_api_data(api_url))

            lookup = build_lookup(all_data)

            # Calculate month shift: file gốc = month 7
            month_shift = 0
            if self.input_paths:
                basename = os.path.basename(self.input_paths[0])
                m = re.search(r'(?:tháng|thang)\s*(\d+)', basename, re.IGNORECASE)
                if m:
                    month_shift = int(m.group(1)) - 7

            self.wb, self.results, updated, skipped_dd, skipped_nm = process_excel(self.excel_path, lookup, month_shift)

            # Output filename + title
            self.output_path, month_str = get_output_filename(self.input_paths, os.path.dirname(self.excel_path))
            if not self.output_path:
                self.output_path = self.excel_path
            if month_str:
                month_num = month_str.split('.')[0].zfill(2)
                ws = self.wb.worksheets[1]
                ws.cell(row=3, column=1).value = f"BẢNG PHÂN BỔ CHI PHÍ THÁNG {month_num}/2026"
                ws.title = f"Phân bổ_{month_num}.26"

            # Display results
            self.tree.delete(*self.tree.get_children())
            for r in self.results:
                tag = 'updated' if r['action'] == 'UPDATED' else ('skip' if 'di dời' in r['action'] else 'nomatch')
                self.tree.insert('', 'end', values=(r['row'], r['code'], r['action'], r['date'], r['amount']), tags=(tag,))

            summary = f"Input: {len(all_data)} bản ghi | Updated: {updated} | Skip: {skipped_dd} | No match: {skipped_nm} | Tổng: {updated+skipped_dd+skipped_nm}"
            self.summary_label.config(text=summary, foreground='black')
            self.status_label.config(text="Xử lý xong! Bấm 'Lưu file Excel'.", foreground='green')
            self.save_btn.config(state='normal')

        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
            self.status_label.config(text="Lỗi!", foreground='red')
        finally:
            self.process_btn.config(state='normal')

    def _save(self):
        if not self.wb or not self.output_path:
            return
        try:
            backup_path = self.output_path.replace('.xlsx', '_backup.xlsx')
            if os.path.exists(self.output_path):
                shutil.copy2(self.output_path, backup_path)
            self.wb.save(self.output_path)
            self.wb.close()
            filename = os.path.basename(self.output_path)
            self.status_label.config(text=f"Đã lưu: {filename}", foreground='green')
            messagebox.showinfo("Thành công", f"File đã lưu:\n{filename}")
            if self.auto_open_var.get():
                os.startfile(self.output_path)
            self.save_btn.config(state='disabled')
        except PermissionError:
            messagebox.showerror("Lỗi", "File đang mở! Đóng file rồi thử lại.")
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))

# ==========================================
# RUN
# ==========================================
if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelUpdaterApp(root)
    root.mainloop()
