import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Summary Billing Transport", layout="wide")

st.title("Summary Billing Transport")
st.caption("Converted from VBA → Python (Streamlit Cloud)")

uploaded_file = st.file_uploader(
    "Upload Excel file (Summary Billing To Customer...)", 
    type=["xlsx", "xlsm"]
)

if not uploaded_file:
    st.stop()

# --- Input fields for Month and Year ---
st.subheader("Order Date Settings")

month = st.text_input("Enter Month (1–12):")
year = st.text_input("Enter Year (e.g. 2025):")

# Validate inputs
if not month or not year:
    st.error("⚠️ Please enter both Month and Year.")
    st.stop()

try:
    month = int(month)
    year = int(year)
    if not (1 <= month <= 12):
        st.error("⚠️ Month must be between 1 and 12.")
        st.stop()
except ValueError:
    st.error("⚠️ Month and Year must be numbers.")
    st.stop()

# --- Detect available sheet names ---
xls = pd.ExcelFile(uploaded_file)
available_sheets = xls.sheet_names
target_sheets = [s for s in available_sheets if s.isdigit() and 1 <= int(s) <= 31]

# --- Read each sheet using row 8 as header and add Order Date ---
merged_parts = []
for sheet_name in target_sheets:
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=7)
    day = int(sheet_name)
    try:
        order_date = datetime(year, month, day).date()
    except ValueError:
        order_date = None
    df = df.assign(sheet_name=sheet_name, Order_Date=order_date)
    merged_parts.append(df)

# --- Merge all sheets ---
if merged_parts:
    merged_df = pd.concat(merged_parts, ignore_index=True)
else:
    merged_df = pd.DataFrame()
    st.warning("No sheets named 1–31 were found in the uploaded file.")

# --- Rename specific columns once ---
rename_map = {
    "ADDRESS": "ADDRESS 1",
    "Unnamed: 8": "ADDRESS 2",
    "Unnamed: 9": "PROVINCE",
    "DU/ORDER": "Du_Order"
}
merged_df = merged_df.rename(columns=rename_map)

# --- Convert all column names to Proper Case ---
merged_df.columns = [str(col).strip().title() for col in merged_df.columns]

# --- Show actual columns for debugging ---
st.write("Columns after renaming:", merged_df.columns.tolist())

# --- Sort by Order Date and Du_Order if they exist ---
sort_cols = [c for c in ["Order Date", "Du_Order"] if c in merged_df.columns]
if sort_cols:
    merged_df = merged_df.sort_values(by=sort_cols, ascending=True)

# --- Display results ---
st.subheader("Merged DataFrame (Sheets 1–31)")
st.dataframe(merged_df, use_container_width=True)

st.write("Number of rows:", len(merged_df))
st.write("Sheets merged:", target_sheets)

# --- Download Excel with grouping style ---
if not merged_df.empty:
    wb = Workbook()
    ws = wb.active
    ws.title = "Merged"

    # Write DataFrame to sheet
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws.append(r)

    # Style header
    header_fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = header_fill

    # Group rows visually by Order Date (shade alternating groups)
    if "Order Date" in merged_df.columns:
        order_date_col = merged_df.columns.get_loc("Order Date") + 1
        prev_date = None
        fill1 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        fill2 = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        current_fill = fill1

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            date_value = row[order_date_col - 1].value
            if date_value != prev_date:
                # Switch fill when new Order Date starts
                current_fill = fill2 if current_fill == fill1 else fill1
                prev_date = date_value
            for cell in row:
                cell.fill = current_fill

    # Save to buffer
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_data = excel_buffer.getvalue()

    st.download_button(
        label="Download merged data as Excel",
        data=excel_data,
        file_name=f"merged_sheets_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
