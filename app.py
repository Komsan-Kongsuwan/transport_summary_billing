import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

st.set_page_config(page_title="Summary Billing Transport", layout="wide")

st.title("Summary Billing Transport")
st.caption("Converted from VBA → Python (Streamlit Cloud)")

uploaded_file = st.file_uploader(
    "Upload Excel file (Summary Billing To Customer...)",
    type=["xlsx", "xlsm"]
)

if not uploaded_file:
    st.stop()

# --- Detect available sheet names and filter to 1..31 ---
xls = pd.ExcelFile(uploaded_file)
available_sheets = xls.sheet_names
target_sheets = [s for s in available_sheets if s.isdigit() and 1 <= int(s) <= 31]

# --- Read each sheet using row 8 (header=7) and add sheet_name ---
merged_parts = []
for sheet_name in target_sheets:
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=7)
    # Add sheet name for provenance
    df = df.assign(sheet_name=sheet_name)
    merged_parts.append(df)

# --- Merge all sheets ---
if merged_parts:
    merged_df = pd.concat(merged_parts, ignore_index=True)
else:
    merged_df = pd.DataFrame()
    st.warning("No sheets named 1–31 were found in the uploaded file.")

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
