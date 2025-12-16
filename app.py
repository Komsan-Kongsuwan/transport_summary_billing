import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime

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
    # Build Order Date from Year, Month, and sheet_name (day)
    day = int(sheet_name)
    try:
        order_date = datetime(year, month, day).date()
    except ValueError:
        order_date = None  # invalid day for that month/year
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
    "DU/ORDER": "DU_Order"
}
merged_df = merged_df.rename(columns=rename_map)

# --- Convert all column names to Proper Case (Title Case) ---
merged_df.columns = [str(col).strip().title() for col in merged_df.columns]

# --- Display results ---
st.subheader("Merged DataFrame (Sheets 1–31)")
st.dataframe(merged_df, use_container_width=True)

st.write("Number of rows:", len(merged_df))
st.write("Sheets merged:", target_sheets)

# --- Download buttons ---
if not merged_df.empty:
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
