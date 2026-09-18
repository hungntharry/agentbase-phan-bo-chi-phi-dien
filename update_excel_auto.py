# -*- coding: utf-8 -*-
"""
Script: Update Excel with PDF payment data
Usage:
  python update_excel_auto.py                    -> list all PDFs, pick one
  python update_excel_auto.py "filename.pdf"     -> run on specific PDF
  python update_excel_auto.py --all              -> run on all PDFs sequentially

- Automatically finds Excel file in the directory
- Reads payment data from PDF (code, date, amount)
- Matches by "Mã khách hàng" in Excel sheet "Phân bổ"
- Updates "Kỳ thanh toán này" (date) and "Tổng tiền của một điểm" (amount)
- Rows with "di dời" text: amount set to blank, no update
- Creates backup before saving
"""
import pdfplumber
import openpyxl
import re
import sys
import io
import os
import shutil
from datetime import datetime, timedelta

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# ==========================================
# CONFIG
# ==========================================
DEFAULT_DIR = r"D:\File storage\AI\AI Agent\Điện"
DATA_START_ROW = 5
COL_CODE = 8              # Mã khách hàng
COL_KY_TT = 13            # Kỳ thanh toán này
COL_KY_TT_TRUOC = 14      # Kỳ thanh toán trước
COL_TONG_TIEN_DIEM = 15   # Tổng tiền của một điểm

def add_months(date_str, n):
    """Add n months to a dd/mm/yyyy date string (n can be negative)"""
    try:
        d = datetime.strptime(date_str.strip(), '%d/%m/%Y')
        total_months = d.year * 12 + d.month - 1 + n
        year = total_months // 12
        month = total_months % 12 + 1
        import calendar
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

def calc_ky_tt_from_truoc(ky_truoc_str):
    """Calculate 'Kỳ thanh toán này' from 'Kỳ thanh toán trước' by adding 1 month"""
    """Input: '01/06/2026 - 30/06/2026' -> Output: '01/07/2026 - 30/07/2026'"""
    if not ky_truoc_str:
        return None
    match = re.match(r'(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', str(ky_truoc_str).strip())
    if not match:
        return None
    start_new = add_one_month(match.group(1))
    end_new = add_one_month(match.group(2))
    if start_new and end_new:
        return f"{start_new} - {end_new}"
    return None

# ==========================================
# STEP 0: Find files
# ==========================================
print("=" * 80)
print("STEP 0: FINDING FILES")
print("=" * 80)

# Find subfolder starting with "Đ" (Điện)
subfolder = None
for item in os.listdir(DEFAULT_DIR):
    if item.startswith('Đ') and 'i' in item.lower():
        subfolder = item
        break

if subfolder:
    base_path = os.path.join(DEFAULT_DIR, subfolder)
else:
    base_path = DEFAULT_DIR

print(f"Directory: {base_path}")

# Find all PDF files
all_pdfs = []
excel_path = None

for f in os.listdir(base_path):
    full = os.path.join(base_path, f)
    if f.startswith('~$') or 'backup' in f.lower():
        continue
    if f.endswith('.pdf'):
        all_pdfs.append(full)
    elif f.endswith('.xlsx') and 'gốc' in f.lower():
        excel_path = full

# Fallback: any xlsx if no "gốc" file found
if not excel_path:
    for f in os.listdir(base_path):
        full = os.path.join(base_path, f)
        if f.startswith('~$') or 'backup' in f.lower():
            continue
        if f.endswith('.xlsx'):
            excel_path = full

if not all_pdfs:
    print("ERROR: No PDF file found!")
    sys.exit(1)
if not excel_path:
    print("ERROR: No Excel file found!")
    sys.exit(1)

# Determine which PDF(s) to process
arg = sys.argv[1] if len(sys.argv) > 1 else None

if arg == '--all':
    pdfs_to_process = all_pdfs
elif arg and arg.endswith('.pdf'):
    # Find matching PDF by filename
    matched_pdf = None
    for p in all_pdfs:
        if arg.lower() in os.path.basename(p).lower():
            matched_pdf = p
            break
    if not matched_pdf:
        print(f"ERROR: PDF file '{arg}' not found!")
        print(f"Available PDFs:")
        for i, p in enumerate(all_pdfs):
            print(f"  [{i+1}] {os.path.basename(p)}")
        sys.exit(1)
    pdfs_to_process = [matched_pdf]
else:
    # No argument or invalid: list all PDFs and pick
    if len(all_pdfs) == 1:
        pdfs_to_process = [all_pdfs[0]]
    else:
        print(f"\nFound {len(all_pdfs)} PDF files:")
        for i, p in enumerate(all_pdfs):
            print(f"  [{i+1}] {os.path.basename(p)}")
        print(f"\nUsage:")
        print(f"  python update_excel_auto.py \"<filename>\"  -> run on specific PDF")
        print(f"  python update_excel_auto.py --all          -> run on all PDFs")
        sys.exit(0)

backup_path = excel_path.replace('.xlsx', '_backup.xlsx')

print(f"  Excel: {os.path.basename(excel_path)}")
print(f"  Backup: {os.path.basename(backup_path)}")
print(f"  PDF(s) to process: {len(pdfs_to_process)}")
for p in pdfs_to_process:
    print(f"    - {os.path.basename(p)}")

# ==========================================
# STEP 1: Extract data from ALL PDFs
# ==========================================
print("\n" + "=" * 80)
print("STEP 1: EXTRACTING DATA FROM PDF")
print("=" * 80)

pdf_data = []

for pdf_path in pdfs_to_process:
    print(f"\n  Reading: {os.path.basename(pdf_path)}")
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    stt = str(row[0]).strip() if row[0] else ""
                    if not stt.isdigit():
                        continue

                    date_raw = str(row[1]).strip() if row[1] else ""
                    desc = str(row[5]).strip() if row[5] else ""
                    debit = str(row[6]).strip() if row[6] else ""
                    credit = str(row[7]).strip() if row[7] else ""

                    # Extract payment code
                    code_match = re.match(r'^([A-Z]{2}[0-9A-Z]+)', desc)
                    if not code_match:
                        code_match = re.search(r'\b([A-Z]{2}\d{8,})\b', desc)
                    payment_code = code_match.group(1) if code_match else ""

                    # Parse date (dd/mm/yy -> dd/mm/yyyy)
                    date_match = re.match(r'(\d{2}/\d{2}/\d{2})', date_raw)
                    trans_date = date_match.group(1) if date_match else ""
                    if trans_date:
                        parts = trans_date.split('/')
                        trans_date = f"{parts[0]}/{parts[1]}/20{parts[2]}"

                    # Parse amount (prefer debit, fallback to credit)
                    amount_str = ""
                    if debit and debit != '0':
                        amount_str = debit.replace(',', '').replace('.00', '').strip()
                    elif credit and credit != '0':
                        amount_str = credit.replace(',', '').replace('.00', '').strip()
                    amount = float(amount_str) if amount_str else 0

                    # Skip MSB internal transactions
                    if payment_code and payment_code != 'MSB':
                        pdf_data.append({
                            'code': payment_code,
                            'date': trans_date,
                            'amount': amount,
                        })

# Build lookup (first occurrence wins)
pdf_lookup = {}
for item in pdf_data:
    if item['code'] not in pdf_lookup:
        pdf_lookup[item['code']] = item

print(f"Extracted {len(pdf_data)} transactions ({len(pdf_lookup)} unique codes)")

# ==========================================
# STEP 2: Backup Excel
# ==========================================
print("\n" + "=" * 80)
print("STEP 2: BACKING UP EXCEL")
print("=" * 80)
shutil.copy2(excel_path, backup_path)
print(f"Backup created: {os.path.basename(backup_path)}")

# ==========================================
# STEP 3: Update Excel
# ==========================================
print("\n" + "=" * 80)
print("STEP 3: UPDATING EXCEL")
print("=" * 80)

wb = openpyxl.load_workbook(excel_path, keep_vba=False)
ws = wb.worksheets[1]  # Second sheet = "Phân bổ"

updated = 0
skipped_di_doi = 0
skipped_no_match = 0

# Calculate month shift: file gốc = month 7, shift = input_month - 7
month_shift = 0
pdf_basename = os.path.basename(pdfs_to_process[0])
month_match_shift = re.search(r'(?:tháng|thang)\s*(\d+)', pdf_basename, re.IGNORECASE)
if month_match_shift:
    input_month = int(month_match_shift.group(1))
    month_shift = input_month - 7
    print(f"File gốc = tháng 7 | Input = tháng {input_month} | Shift = {month_shift} tháng")

print(f"\n{'Row':>5} | {'Code':<20} | {'Action':<16} | {'Amount':>15}")
print("-" * 65)

for row_idx in range(DATA_START_ROW, ws.max_row + 1):
    excel_code = ws.cell(row=row_idx, column=COL_CODE).value
    if not excel_code:
        continue
    excel_code = str(excel_code).strip()

    current_date = ws.cell(row=row_idx, column=COL_KY_TT).value or ""
    current_date_str = str(current_date).lower()

    # Check for "di dời" / "ngừng" text in col 13 or col 14
    ky_truoc = ws.cell(row=row_idx, column=COL_KY_TT_TRUOC).value or ""
    if any(kw in current_date_str for kw in ['di d', 'dời', 'ngừng']) or \
       any(kw in str(ky_truoc).lower() for kw in ['di d', 'dời', 'ngừng']):
        ws.cell(row=row_idx, column=COL_TONG_TIEN_DIEM).value = None
        skipped_di_doi += 1
        print(f"{row_idx:>5} | {excel_code:<20} | {'SKIP (di dời)':<16} | {'blank':>15}")
        continue

    # Shift col 13 and col 14 by month_shift
    if month_shift != 0:
        new_13 = shift_date_range(current_date, month_shift)
        if new_13:
            ws.cell(row=row_idx, column=COL_KY_TT).value = new_13
        new_14 = shift_date_range(ky_truoc, month_shift)
        if new_14:
            ws.cell(row=row_idx, column=COL_KY_TT_TRUOC).value = new_14

    # Only update col 15 (amount)
    if excel_code in pdf_lookup:
        pdf_item = pdf_lookup[excel_code]
        ws.cell(row=row_idx, column=COL_TONG_TIEN_DIEM).value = pdf_item['amount']
        updated += 1
        print(f"{row_idx:>5} | {excel_code:<20} | {'UPDATED':<16} | {pdf_item['amount']:>15,.0f}")
    else:
        skipped_no_match += 1
        print(f"{row_idx:>5} | {excel_code:<20} | {'NO MATCH':<16} | {'---':>15}")

# ==========================================
# STEP 4: Save with PDF-matching filename
# ==========================================
print("\n" + "=" * 80)
print("STEP 4: SAVING")
print("=" * 80)

# Extract month pattern from PDF filename (e.g., "tháng 6.2026", "thang 7.2026")
pdf_basename = os.path.basename(pdfs_to_process[0])
month_match = re.search(r'(?:tháng|thang)\s*(\d+\.?\d*)', pdf_basename, re.IGNORECASE)
if month_match:
    month_str = month_match.group(1)
    # Construct new Excel filename matching the PDF month
    new_excel_name = f"Bảng phân bổ thanh toán chi phí điện tháng {month_str}.xlsx"
    new_excel_path = os.path.join(base_path, new_excel_name)
    # Update title in sheet (Row 3, Col 1): BẢNG PHÂN BỔ CHI PHÍ THÁNG XX/2026
    month_num = month_str.split('.')[0].zfill(2)
    new_title = f"BẢNG PHÂN BỔ CHI PHÍ THÁNG {month_num}/2026"
    ws.cell(row=3, column=1).value = new_title
    print(f"Updated title: {new_title}")
    # Rename sheet: Phân bổ_XX.26
    ws.title = f"Phân bổ_{month_num}.26"
else:
    # Fallback: use original Excel name
    new_excel_path = excel_path

# Save to the new filename
wb.save(new_excel_path)
wb.close()

print(f"\nFile saved: {os.path.basename(new_excel_path)}")
if new_excel_path != excel_path:
    print(f"  (Renamed from: {os.path.basename(excel_path)})")
print(f"\n{'=' * 40}")
print(f"  UPDATED:          {updated} rows")
print(f"  SKIPPED (di dời): {skipped_di_doi} rows (amount = blank)")
print(f"  NO MATCH:         {skipped_no_match} rows")
print(f"  TOTAL:            {updated + skipped_di_doi + skipped_no_match} rows")
print(f"{'=' * 40}")
