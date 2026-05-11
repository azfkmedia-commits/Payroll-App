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

# --- STABLE OCR LOADING ---
@st.cache_resource
def load_ocr():
    # gpu=False is critical for free hosting stability
    return easyocr.Reader(['en'], gpu=False)

try:
    reader = load_ocr()
    st.sidebar.success("✅ Scanner Engine Ready")
except Exception as e:
    st.sidebar.error("Loading scanner... please wait 2 mins.")

# --- STYLING & CONSTANTS ---
DAILY_RATE = 69.00
EPF_RATE = 0.11
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def process_images(files):
    attendance = {}
    for f in files:
        # Resize to prevent memory crash
        img_raw = Image.open(f).convert('RGB')
        img_raw.thumbnail((800, 800)) 
        img = np.array(img_raw)
        
        results = reader.readtext(img)
        current_date = None
        
        for (_, text, _) in results:
            date_match = re.search(r'(\d{4}/\d{2}/\d{2})', text)
            if date_match:
                current_date = datetime.strptime(date_match.group(1), "%Y/%m/%d").strftime("%d-%m-%Y")
            
            clean_text = text.strip().upper()
            if len(clean_text) > 2 and clean_text.isalpha() and current_date:
                if clean_text not in attendance:
                    attendance[clean_text] = set()
                attendance[clean_text].add(current_date)
    return attendance

# --- UI ---
st.title("📸 Shop Payroll Pro")
st.info("Daily Rate: RM 69.00 | EPF: 11%")

files = st.file_uploader("Upload Attendance Photo", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

if files:
    with st.spinner("Analyzing text... (May take 1 min)"):
        data = process_images(files)
        if data:
            st.success(f"Detected: {', '.join(data.keys())}")
            # ... (Rest of your Excel generation logic goes here)
