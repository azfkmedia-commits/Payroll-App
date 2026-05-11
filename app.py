import streamlit as st
import pandas as pd
import numpy as np
import easyocr
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, timedelta
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Payroll Pro", layout="wide")

# Constants
DAILY_RATE = 69.00
EPF_EMP_RATE = 0.11

# Excel Styling
DARK_BLUE = PatternFill(start_color="333F4F", end_color="333F4F", fill_type="solid")
GRAY_HEADER = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")
YELLOW_OFF = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF", bold=True)
BLACK_BOLD = Font(color="000000", bold=True)
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'])

reader = load_ocr()

def get_days_in_month(year=2026, month=4):
    start = datetime(year, month, 1)
    days = []
    while start.month == month:
        days.append(start.strftime("%d-%m-%Y"))
        start += timedelta(days=1)
    return days

def create_excel(attendance_data):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    dates = get_days_in_month()
    
    current_row = 1
    # attendance_data is now a dict: { "NAME": [list of dates present] }
    for name, present_dates in attendance_data.items():
        # Header Banner
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell = ws.cell(row=current_row, column=1, value=name.upper())
        cell.fill = DARK_BLUE
        cell.font = WHITE_FONT
        cell.alignment = Alignment(horizontal="center")
        current_row += 1

        # Column Headers
        cols = ["DATE", "DAY", "IN", "LATE IN", "KOMISYEN", "OUT", "EXTRA HOUR"]
        for i, text in enumerate(cols, 1):
            c = ws.cell(row=current_row, column=i, value=text)
            c.fill = GRAY_HEADER
            c.font = BLACK_BOLD
            c.border = THIN_BORDER
        current_row += 1

        # Monthly Data Rows
        present_count = 0
        for d_str in dates:
            day_name = datetime.strptime(d_str, "%d-%m-%Y").strftime("%A").upper()
            is_present = d_str in present_dates
            
            row_data = [d_str, day_name, "09:00" if is_present else "", "", "" if is_present else "OFF DAY", "18:00" if is_present else "", ""]
            
            for i, val in enumerate(row_data, 1):
                c = ws.cell(row=current_row, column=i, value=val)
                c.border = THIN_BORDER
                if not is_present and i == 5: # Highlight OFF DAY in yellow
                    c.fill = YELLOW_OFF
                    c.font = BLACK_BOLD
            
            if is_present: present_count += 1
            current_row += 1

        # Total Row
        gross = present_count * DAILY_RATE
        net = gross * (1 - EPF_EMP_RATE)
        ws.cell(row=current_row, column=1, value=f"TOTAL DAYS: {present_count} | NET SALARY: RM {net:.2f}").font = BLACK_BOLD
        current_row += 3

    wb.save(output)
    return output.getvalue()

# --- APP UI ---
st.title("📸 Dynamic Payroll Scanner")
st.write("Upload logs; names are detected automatically.")

files = st.file_uploader("Upload Photos", type=['jpg', 'png', 'jpeg'], accept_multiple_files=True)

if files:
    # This dictionary will store whatever names the AI finds
    detected_attendance = {} 
    
    with st.spinner("Analyzing text..."):
        for f in files:
            img = np.array(Image.open(f))
            results = reader.readtext(img)
            
            # Logic: If the AI finds a word that looks like a name, add it to the list
            for (_, text, _) in results:
                clean_text = text.strip().upper()
                # Simple filter: ignore numbers/dates, focus on words (Names)
                if len(clean_text) > 2 and clean_text.isalpha():
                    if clean_text not in detected_attendance:
                        detected_attendance[clean_text] = []
                    # For this demo, we assume the found name is present on the current date
                    # In your real setup, you'd link the name to the date nearby in the photo
                    today_str = "01-04-2026" 
                    if today_str not in detected_attendance[clean_text]:
                        detected_attendance[clean_text].append(today_str)

    st.success(f"Detected Employees: {', '.join(detected_attendance.keys())}")
    
    if st.button("Generate Report"):
        data = create_excel(detected_attendance)
        st.download_button("Download Excel", data, "Payroll_Report.xlsx")
