import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, date
from openpyxl import load_workbook

st.set_page_config(page_title="Summary Billing Transport", layout="wide")

st.title("Summary Billing Transport")
st.caption("Converted from VBA → Python (Streamlit Cloud)")

uploaded_file = st.file_uploader(
    "Upload Excel file (Summary Billing To Customer...)", 
    type=["xlsx", "xlsm"]
)

if not uploaded_file:
    st.stop()

# --- Helper function to parse Thai Buddhist Era date ---
def parse_thai_date(date_str):
    if not date_str:
        return None
    cleaned = str(date_str).replace("วันที่", "").strip()
    match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", cleaned)
    if match:
        day, month, year = map(int, match.groups())
        year -= 543  # Convert BE to Gregorian
        return datetime(year, month, day).date()
    return None

# --- Detect available sheet names ---
xls = pd.ExcelFile(uploaded_file)
available_sheets = xls.sheet_names
target_sheets = [s for s in available_sheets if s.isdigit() and 1 <= int(s) <= 31]

# --- Use openpyxl to read cell A7 from each sheet ---
wb = load_workbook(uploaded_file, data_only=True)

merged_parts = []
for sheet_name in target_sheets:
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    
    # Read cell A7 directly
    ws = wb[sheet_name]
    raw_date = ws["A7"].value
    order_date = parse_thai_date(raw_date)
    
    # Assign Order_Date and sheet_name to all rows
    df = df.assign(sheet_name=sheet_name, Order_Date=order_date)
    merged_parts.append(df)

# --- Merge all sheets ---
merged_df = pd.concat(merged_parts, ignore_index=True)

# --- Display results ---
st.subheader("Merged DataFrame (Sheets 1–31)")
st.dataframe(merged_df)

st.write("Number of rows:", len(merged_df))
st.write("Sheets merged:", target_sheets)

# --- Download buttons ---
# CSV
csv_data = merged_df.to_csv(index=False).encode("utf-8")
st.download_button(
    label="Download merged data as CSV",
    data=csv_data,
    file_name=f"merged_sheets_{date.today()}.csv",
    mime="text/csv"
)

# Excel
excel_buffer = BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    merged_df.to_excel(writer, index=False, sheet_name="Merged")
excel_data = excel_buffer.getvalue()

st.download_button(
    label="Download merged data as Excel",
    data=excel_data,
    file_name=f"merged_sheets_{date.today()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
