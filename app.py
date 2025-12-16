import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
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

# --- Read and merge sheets 1..31 ---
sheet_names = [str(i) for i in range(1, 32)]
dfs = pd.read_excel(uploaded_file, sheet_name=sheet_names)

merged_df = pd.concat(
    [df.assign(sheet_name=name) for name, df in dfs.items()],
    ignore_index=True
)

# --- Display results ---
st.subheader("Merged DataFrame (Sheets 1–31)")
st.dataframe(merged_df)

st.write("Number of rows:", len(merged_df))
st.write("Sheets merged:", sheet_names)

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
