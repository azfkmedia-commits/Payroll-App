import streamlit as st
import pandas as pd
import numpy as np
import easyocr
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta
from io import BytesIO
import re

# --- APP CONFIG ---
st.set_page_config(page_title="Shop Payroll Scanner", layout="centered")

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'])

reader = load_ocr()

# --- STYLING ---
DARK_BLUE = PatternFill(start_color="333F4F", end_color="333F4F", fill_type="solid")
GRAY_BG = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")
YELLOW_OFF = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF", bold=True)
BLACK_BOLD = Font(color="000000", bold=True)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# --- LOGIC ---
def process_images(files):
    attendance = {}
    for f in files:
        img = np.array(Image.open(f))
        results = reader.readtext(img)
        
        current_date = None
        for (_, text, _) in results:
            # Detect Date (YYYY/MM/DD)
            date_match = re.search(r'(\d{4}/\d{2}/\d{2})', text)
            if date_match:
                current_date = datetime.strptime(date_match.group(1), "%Y/%m/%d").strftime("%d-%m-%Y")
            
            # Detect Name (Alpha words > 2 letters)
            clean_text = text.strip().upper()
            if len(clean_text) > 2 and clean_text.isalpha() and current_date:
                if clean_text not in attendance:
                    attendance[clean_text] = set()
                attendance[clean_text].add(current_date)
    return attendance

def generate_payroll(data):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # Generate for April 2026
    start = datetime(2026, 4, 1)
    month_dates = [(start + timedelta(days=i)).strftime("%d-%m-%Y") for i in range(30)]
    
    curr_row = 1
    for name, dates_present in data.items():
        # Header
        ws.merge_cells(start_row=curr_row, column=1, end_row=curr_row, end_column=7)
        c = ws.cell(row=curr_row, column=1, value=name)
        c.fill, c.font, c.alignment = DARK_BLUE, WHITE_FONT, Alignment(horizontal="center")
        curr_row += 1
        
        # Payroll Rows
        present_count = 0
        for d_str in month_dates:
            is_present = d_str in dates_present
            day_name = datetime.strptime(d_str, "%d-%m-%Y").strftime("%A").upper()
            
            row_data = [d_str, day_name, "09:00" if is_present else "", "", "" if is_present else "OFF DAY", "18:00" if is_present else "", ""]
            for i, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=i, value=val)
                cell.border = BORDER
                if not is_present and i == 5: cell.fill = YELLOW_OFF
            
            if is_present: present_count += 1
            curr_row += 1
            
        # Calculation
        gross = present_count * 69.00
        net = gross * 0.89 # 11% EPF deduction
        ws.cell(row=curr_row, column=1, value=f"DAYS: {present_count} | NET: RM {net:.2f}").font = BLACK_BOLD
        curr_row += 3
        
    wb.save(output)
    return output.getvalue()

# --- UI ---
st.title("📸 Shop Payroll Pro")
st.info("Snap photos of logs. I will calculate RM 69/day + 11% EPF automatically.")

files = st.file_uploader("Upload or Take Photo", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if files:
    with st.spinner("Reading logs..."):
        data = process_images(files)
        if data:
            st.success(f"Found {len(data)} staff members.")
            for s in data.keys(): st.write(f"✅ {s}")
            
            report = generate_payroll(data)
            st.download_button("📥 Download April Report", report, "April_Payroll.xlsx")
        else:
            st.warning("No names or dates found. Try a clearer photo.")
