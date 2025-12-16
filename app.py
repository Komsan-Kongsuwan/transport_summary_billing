from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Sort by Order Date and Du_Order ---
sort_cols = [c for c in ["Order Date", "Du_Order"] if c in merged_df.columns]
if sort_cols:
    merged_df = merged_df.sort_values(by=sort_cols, ascending=True)

# --- Create workbook ---
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

# --- Group by Order Date and Du_Order ---
if "Order Date" in merged_df.columns and "Du_Order" in merged_df.columns:
    order_date_col = merged_df.columns.get_loc("Order Date") + 1
    du_order_col = merged_df.columns.get_loc("Du_Order") + 1

    prev_date = None
    prev_du = None
    start_date_row = 2
    start_du_row = 2

    fill1 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill2 = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    current_fill = fill1

    for row_idx in range(2, ws.max_row + 1):
        date_value = ws.cell(row=row_idx, column=order_date_col).value
        du_value = ws.cell(row=row_idx, column=du_order_col).value

        # Handle Order Date grouping
        if date_value != prev_date:
            if prev_date is not None and row_idx - 1 > start_date_row:
                ws.merge_cells(start_row=start_date_row, start_column=order_date_col,
                               end_row=row_idx - 1, end_column=order_date_col)
                ws.cell(row=start_date_row, column=order_date_col).alignment = Alignment(vertical="center")
            current_fill = fill2 if current_fill == fill1 else fill1
            start_date_row = row_idx
            prev_date = date_value

        # Handle Du_Order grouping
        if du_value != prev_du:
            if prev_du is not None and row_idx - 1 > start_du_row:
                ws.merge_cells(start_row=start_du_row, start_column=du_order_col,
                               end_row=row_idx - 1, end_column=du_order_col)
                ws.cell(row=start_du_row, column=du_order_col).alignment = Alignment(vertical="center")
            start_du_row = row_idx
            prev_du = du_value

        # Apply alternating fill
        for cell in ws[row_idx]:
            cell.fill = current_fill

    # Merge last groups
    if prev_date is not None and ws.max_row >= start_date_row:
        ws.merge_cells(start_row=start_date_row, start_column=order_date_col,
                       end_row=ws.max_row, end_column=order_date_col)
        ws.cell(row=start_date_row, column=order_date_col).alignment = Alignment(vertical="center")

    if prev_du is not None and ws.max_row >= start_du_row:
        ws.merge_cells(start_row=start_du_row, start_column=du_order_col,
                       end_row=ws.max_row, end_column=du_order_col)
        ws.cell(row=start_du_row, column=du_order_col).alignment = Alignment(vertical="center")

# --- Save to buffer for Streamlit download ---
excel_buffer = BytesIO()
wb.save(excel_buffer)
excel_data = excel_buffer.getvalue()

st.download_button(
    label="Download grouped Excel",
    data=excel_data,
    file_name=f"merged_sheets_{date.today()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
