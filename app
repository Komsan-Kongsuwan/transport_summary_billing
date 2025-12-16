import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, date
from io import BytesIO
import calendar

st.set_page_config(page_title="Billing Summary Processor", layout="wide")

st.title("📊 Transport Billing Summary Generator")
st.markdown("Upload your billing data file and generate a consolidated summary report")

# Sidebar for inputs
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # Month and Year selection
    current_year = datetime.now().year
    current_month = datetime.now().month
    
    selected_month = st.selectbox(
        "Select Month",
        options=list(range(1, 13)),
        format_func=lambda x: calendar.month_name[x],
        index=current_month - 1
    )
    
    selected_year = st.number_input(
        "Select Year",
        min_value=2020,
        max_value=2030,
        value=current_year,
        step=1
    )
    
    st.divider()
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xlsm'],
        help="Upload the billing data file with daily sheets (1, 2, 3...31)"
    )

def process_billing_data(file, month, year):
    """Main processing function that mimics the VBA logic"""
    
    # Load workbook
    wb = openpyxl.load_workbook(file, keep_vba=True, data_only=False)
    
    # Column headers definition
    column_headers = [
        "Order Date", "DU", "DU-Order", "CM Code.", "Sold To",
        "Destination Code", "Ship To", "Address 1", "Address 2",
        "Province", "Post Code", "Area", "Total Pick Q'TY",
        "Ship Total WT", "Up 10KG/Chg", "Rate/KG", "Min/Charge",
        "All Charge", "Pick Date", "Transport", "Premium",
        "Remark", "Part No.", "Pick Q'TY", "Total Weight"
    ]
    
    # Create temporary dataframe to collect data
    temp_data = []
    
    # Get days in month
    days_in_month = calendar.monthrange(year, month)[1]
    
    # Process each numeric sheet (representing days)
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for day in range(1, days_in_month + 1):
        sheet_name = str(day)
        if sheet_name in wb.sheetnames:
            status_text.text(f"Processing day {day}...")
            sheet = wb[sheet_name]
            
            # Create delivery date
            delivery_date = date(year, month, day)
            
            # Read data starting from row 9
            row = 9
            while sheet.cell(row, 1).value is not None:
                temp_data.append({
                    'Order Date': delivery_date,
                    'DU': sheet.cell(row, 1).value,
                    'Order': sheet.cell(row, 2).value,
                    'DU-Order': sheet.cell(row, 3).value,
                    'CM Code': sheet.cell(row, 4).value,
                    'Sold To': sheet.cell(row, 5).value,
                    'CN Code': sheet.cell(row, 6).value,
                    'Ship To': sheet.cell(row, 7).value,
                    'Address1': sheet.cell(row, 8).value,
                    'Address2': sheet.cell(row, 9).value,
                    'Province': sheet.cell(row, 10).value,
                    'Post Code': sheet.cell(row, 11).value,
                    'Tel': sheet.cell(row, 12).value,
                    'Part Number': sheet.cell(row, 13).value,
                    'Pick QTY': sheet.cell(row, 14).value,
                    'Free Gift': sheet.cell(row, 15).value,
                    'Remark': sheet.cell(row, 16).value
                })
                row += 1
        
        progress_bar.progress((day) / days_in_month)
    
    status_text.text("Sorting data...")
    
    # Convert to DataFrame and sort
    temp_df = pd.DataFrame(temp_data)
    if len(temp_df) > 0:
        temp_df = temp_df.sort_values(by=['Order Date', 'DU-Order'])
    
    # Load reference sheets
    cargo_sheet = wb['Cargo and Weight']
    sell_sheet = wb['Sell Price']
    
    # Build lookup dictionaries
    status_text.text("Building lookup tables...")
    
    # Cargo and Weight lookup (Part Number -> Weight)
    cargo_lookup = {}
    row = 3
    while cargo_sheet.cell(row, 2).value is not None:
        part_num = cargo_sheet.cell(row, 2).value
        weight = cargo_sheet.cell(row, 5).value
        if part_num and weight:
            cargo_lookup[str(part_num)] = float(weight)
        row += 1
    
    # Sell Price lookup (Post Code -> Area, Min Charge, Rate/KG)
    sell_lookup = {}
    row = 2
    while sell_sheet.cell(row, 1).value is not None:
        post_code = sell_sheet.cell(row, 1).value
        area = sell_sheet.cell(row, 3).value
        min_charge = sell_sheet.cell(row, 4).value
        rate_kg = sell_sheet.cell(row, 5).value
        if post_code:
            sell_lookup[str(int(post_code))] = {
                'area': area,
                'min_charge': float(min_charge) if min_charge else 0,
                'rate_kg': float(rate_kg) if rate_kg else 0
            }
        row += 1
    
    # Process summary data
    status_text.text("Generating summary...")
    summary_data = []
    
    if len(temp_df) > 0:
        # Group by DU-Order
        grouped = temp_df.groupby('DU-Order')
        
        for du_order, group in grouped:
            # Get first row for basic info
            first_row = group.iloc[0]
            
            # Calculate totals
            total_qty = group['Pick QTY'].sum()
            
            # Calculate total weight
            total_weight = 0
            for _, row_data in group.iterrows():
                part_num = str(row_data['Part Number']) if pd.notna(row_data['Part Number']) else ''
                qty = row_data['Pick QTY'] if pd.notna(row_data['Pick QTY']) else 0
                
                if part_num in cargo_lookup:
                    total_weight += qty * cargo_lookup[part_num]
            
            # Get pricing info
            post_code = str(int(first_row['Post Code'])) if pd.notna(first_row['Post Code']) else ''
            
            if post_code in sell_lookup:
                area = sell_lookup[post_code]['area']
                min_charge = sell_lookup[post_code]['min_charge']
                rate_kg = sell_lookup[post_code]['rate_kg']
            else:
                area = "Post Code Not Found"
                min_charge = 0
                rate_kg = 0
            
            # Calculate charges
            up_10kg = max(total_weight - 10, 0) if total_weight > 10 else 0
            all_charge = (up_10kg * rate_kg) + min_charge
            transport = "STL" if area == "BKK" else "DASH"
            
            # Create summary row
            summary_row = {
                'Order Date': first_row['Order Date'],
                'DU': '',
                'DU-Order': du_order,
                'CM Code.': first_row['CM Code'],
                'Sold To': first_row['Sold To'],
                'Destination Code': first_row['CN Code'],
                'Ship To': first_row['Ship To'],
                'Address 1': first_row['Address1'],
                'Address 2': first_row['Address2'],
                'Province': first_row['Province'],
                'Post Code': first_row['Post Code'],
                'Area': area,
                'Total Pick Q\'TY': total_qty,
                'Ship Total WT': total_weight,
                'Up 10KG/Chg': up_10kg,
                'Rate/KG': rate_kg,
                'Min/Charge': min_charge,
                'All Charge': all_charge,
                'Pick Date': '',
                'Transport': transport,
                'Premium': '',
                'Remark': first_row['Remark']
            }
            
            # Add detail rows
            for idx, row_data in group.iterrows():
                part_num = str(row_data['Part Number']) if pd.notna(row_data['Part Number']) else ''
                qty = row_data['Pick QTY'] if pd.notna(row_data['Pick QTY']) else 0
                
                # Calculate individual weight
                if part_num in cargo_lookup:
                    item_weight = qty * cargo_lookup[part_num]
                else:
                    item_weight = 'Part Not Found'
                
                detail_row = summary_row.copy()
                detail_row['Part No.'] = part_num
                detail_row['Pick Q\'TY'] = qty
                detail_row['Total Weight'] = item_weight
                
                summary_data.append(detail_row)
    
    status_text.text("Creating output file...")
    progress_bar.progress(1.0)
    
    # Create output workbook
    output_wb = openpyxl.Workbook()
    summary_sheet = output_wb.active
    summary_sheet.title = "Summary"
    
    # Write headers
    for col_idx, header in enumerate(column_headers, start=1):
        cell = summary_sheet.cell(1, col_idx, header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    
    # Set column widths
    col_widths = [12, 3, 17, 10, 24, 16, 28, 13, 13, 9, 10, 9, 15, 13, 14, 9, 13, 15, 14, 9, 9, 18, 16, 10, 12]
    for col_idx, width in enumerate(col_widths, start=1):
        summary_sheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width
    
    summary_sheet.row_dimensions[1].height = 35
    
    # Freeze first row
    summary_sheet.freeze_panes = 'A2'
    
    # Write data
    row_num = 2
    for item in summary_data:
        for col_idx, header in enumerate(column_headers, start=1):
            value = item.get(header, '')
            cell = summary_sheet.cell(row_num, col_idx, value)
            
            # Format numbers
            if header in ['Total Pick Q\'TY', 'Ship Total WT', 'Up 10KG/Chg', 'All Charge', 'Pick Q\'TY', 'Total Weight']:
                if isinstance(value, (int, float)) and value != 'Part Not Found':
                    cell.number_format = '#,##0.00'
        
        row_num += 1
    
    # Add summary formulas at bottom
    if row_num > 2:
        last_data_row = row_num - 1
        
        # Summary row
        summary_sheet.cell(row_num + 3, 13, f'=SUM(M2:M{last_data_row})')
        summary_sheet.cell(row_num + 3, 14, f'=SUM(N2:N{last_data_row})')
        summary_sheet.cell(row_num + 3, 18, f'=SUM(R2:R{last_data_row})')
        
        # Labels
        summary_sheet.cell(row_num + 4, 13, 'CTN')
        summary_sheet.cell(row_num + 4, 14, 'KG')
        summary_sheet.cell(row_num + 4, 18, 'BATH')
        
        # FSC calculation
        summary_sheet.cell(row_num + 6, 13, 'Fuel surcharge (FSC)')
        summary_sheet.cell(row_num + 6, 14, f'=R{row_num + 3}*0.1362')
        summary_sheet.cell(row_num + 6, 18, f'=R{row_num + 3}+N{row_num + 6}')
        
        summary_sheet.cell(row_num + 7, 14, 'BATH')
        summary_sheet.cell(row_num + 7, 18, 'BATH')
        
        # Format summary cells
        for col in [13, 14, 18]:
            summary_sheet.cell(row_num + 3, col).number_format = '#,##0.00'
            summary_sheet.cell(row_num + 6, col).number_format = '#,##0.00'
            summary_sheet.cell(row_num + 3, col).font = Font(bold=True)
            summary_sheet.cell(row_num + 6, col).font = Font(bold=True)
        
        summary_sheet.cell(row_num + 6, 18).fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        
        # Alignment
        for row in [row_num + 3, row_num + 4, row_num + 6, row_num + 7]:
            for col in [13, 14, 18]:
                summary_sheet.cell(row, col).alignment = Alignment(horizontal='center')
    
    status_text.text("Complete!")
    progress_bar.empty()
    
    # Save to BytesIO
    output = BytesIO()
    output_wb.save(output)
    output.seek(0)
    
    return output, len(summary_data)

# Main processing
if uploaded_file is not None:
    st.success("✅ File uploaded successfully!")
    
    # Display file info
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Selected Month", calendar.month_name[selected_month])
    with col2:
        st.metric("Selected Year", selected_year)
    with col3:
        st.metric("File Name", uploaded_file.name)
    
    st.divider()
    
    # Process button
    if st.button("🚀 Generate Summary", type="primary", use_container_width=True):
        with st.spinner("Processing billing data..."):
            try:
                output_file, record_count = process_billing_data(
                    uploaded_file,
                    selected_month,
                    selected_year
                )
                
                st.success(f"✅ Processing complete! Generated {record_count} records.")
                
                # Download button
                output_filename = f"Summary_Billing_{calendar.month_name[selected_month]}_{selected_year}.xlsx"
                
                st.download_button(
                    label="⬇️ Download Summary Report",
                    data=output_file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedococument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.balloons()
                
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
                st.exception(e)
else:
    st.info("👆 Please upload an Excel file to begin processing")
    
    # Display instructions
    with st.expander("📖 How to use this app"):
        st.markdown("""
        ### Instructions:
        
        1. **Select Month and Year** in the sidebar
           - Choose the month and year for the billing period
        
        2. **Upload Excel File** 
           - Upload your billing data file (.xlsx or .xlsm)
           - File should contain:
             - Daily sheets named 1, 2, 3... 31 (for each day)
             - "Cargo and Weight" sheet with part numbers and weights
             - "Sell Price" sheet with post codes and pricing info
        
        3. **Generate Summary**
           - Click the "Generate Summary" button
           - Wait for processing to complete
        
        4. **Download Result**
           - Download the generated Summary report
           - The report will contain consolidated billing data
        
        ### Output:
        - Consolidated summary sheet with all orders
        - Calculated weights and charges
        - Transport assignments (STL/DASH)
        - Summary totals with FSC (Fuel Surcharge)
        """)

# Footer
st.divider()
st.caption("Billing Summary Processor v1.0 | Built with Streamlit")
