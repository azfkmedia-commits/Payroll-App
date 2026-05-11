import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- SETTINGS ---
st.set_page_config(page_title="Excel Payroll Processor", layout="wide")
DAILY_RATE = 69.00
EPF_RATE = 0.11

# --- STYLING ---
DARK_BLUE = PatternFill(start_color="333F4F", end_color="333F4F", fill_type="solid")
YELLOW_OFF = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF", bold=True)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def generate_payroll_from_excel(df):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # Get all unique employee names from the 'Name' column
    # This ensures it works even if employees change
    employees = df['Name'].unique()
    
    # [span_2](start_span)Setup for April 2026[span_2](end_span)
    start_date = datetime(2026, 4, 1)
    month_days = [(start_date + timedelta(days=i)).strftime("%d/%m/%Y") for i in range(30)]
    
    curr_row = 1
    for name in employees:
        # Get dates this specific person worked
        worked_dates = df[df['Name'] == name]['Date'].astype(str).tolist()
        
        # Staff Header
        ws.merge_cells(start_row=curr_row, column=1, end_row=curr_row, end_column=7)
        c = ws.cell(row=curr_row, column=1, value=name.upper())
        c.fill, c.font, c.alignment = DARK_BLUE, WHITE_FONT, Alignment(horizontal="center")
        curr_row += 1
        
        # Payroll Rows
        present_count = 0
        for d_str in month_days:
            is_present = any(d_str in wd for wd in worked_dates)
            day_name = datetime.strptime(d_str, "%d/%m/%Y").strftime("%A").upper()
            
            row_data = [d_str, day_name, "09:00" if is_present else "", "", "" if is_present else "OFF DAY", "18:00" if is_present else "", ""]
            for i, val in enumerate(row_data, 1):
                cell = ws.cell(row=curr_row, column=i, value=val)
                cell.border = BORDER
                if not is_present and i == 5: cell.fill = YELLOW_OFF
            
            if is_present: present_count += 1
            curr_row += 1
            
        # Calculation: RM 69/day - 11% EPF
        gross = present_count * DAILY_RATE
        net = gross * (1 - EPF_RATE)
        ws.cell(row=curr_row, column=1, value=f"DAYS: {present_count} | TOTAL: RM {net:.2f}").font = Font(bold=True)
        curr_row += 3
        
    wb.save(output)
    return output.getvalue()

# --- UI ---
st.title("📊 Excel-to-Payroll Converter")
st.write("Upload your raw attendance Excel to generate the formatted Payroll report.")

uploaded_file = st.file_uploader("Upload Raw Attendance (Excel)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    if 'Name' in df.columns and 'Date' in df.columns:
        st.success(f"File loaded! Found {len(df['Name'].unique())} employees.")
        
        if st.button("Generate Formatted Report"):
            result = generate_payroll_from_excel(df)
            st.download_button("📥 Download Payroll Excel", result, "April_Payroll_Final.xlsx")
    else:
        st.error("Your Excel must have 'Name' and 'Date' columns.")
