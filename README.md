# Transport Billing Summary Generator

A Streamlit web application that processes transport billing data and generates consolidated summary reports. This app converts VBA Excel macros to a cloud-based Python solution.

## Features

- 📤 **Upload Excel Files**: Support for .xlsx and .xlsm formats
- 📅 **Month/Year Selection**: Easily select the billing period
- 🔄 **Automatic Processing**: Reads daily sheets (1-31) and consolidates data
- 💰 **Price Calculations**: Automatic weight and charge calculations
- 📊 **Summary Reports**: Generates formatted Excel summary with totals
- ⬇️ **Easy Download**: One-click download of processed reports

## How It Works

The app processes billing data by:

1. Reading daily sheets (numbered 1, 2, 3...31) representing each day of the month
2. Combining data with delivery dates based on selected month/year
3. Looking up part weights from "Cargo and Weight" sheet
4. Looking up pricing from "Sell Price" sheet based on post codes
5. Calculating total quantities, weights, and charges per order
6. Generating a formatted "Summary" sheet with all calculations

## Input File Requirements

Your Excel file must contain:

### Required Sheets:

1. **Daily Sheets (1, 2, 3...31)**
   - Named with day numbers (1-31)
   - Data starts at row 9
   - Columns: DU, Order, DU-Order, CM Code, Sold To, CN Code, Ship To, Addresses, Province, Post Code, Tel, Part Number, Pick QTY, Free Gift, Remark

2. **Cargo and Weight Sheet**
   - Row 2: Headers
   - Row 3+: Data
   - Columns: No, Product (Part Number), Package, Weight_Actual
   - Used for part number weight lookup

3. **Sell Price Sheet**
   - Row 1: Headers
   - Row 2+: Data
   - Columns: PostCodeMain, ProvinceEng, Area, MinCharge, Sell price/kg
   - Used for pricing based on post codes

## Installation

### Local Installation

```bash
# Clone or download the files
cd billing-summary-app

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run billing_summary_app.py
```

### Cloud Deployment (Streamlit Cloud)

1. **Create a GitHub Repository**
   - Upload `billing_summary_app.py` and `requirements.txt`

2. **Deploy to Streamlit Cloud**
   - Go to https://share.streamlit.io/
   - Sign in with GitHub
   - Click "New app"
   - Select your repository and branch
   - Set main file path: `billing_summary_app.py`
   - Click "Deploy"

3. **Share Your App**
   - You'll get a public URL like: `https://yourapp.streamlit.app`
   - Share this URL with your team

### Alternative Cloud Platforms

#### Heroku
```bash
# Add Procfile
echo "web: streamlit run billing_summary_app.py --server.port=$PORT" > Procfile

# Deploy
git init
git add .
git commit -m "Initial commit"
heroku create your-app-name
git push heroku main
```

#### Google Cloud Run
```bash
# Create Dockerfile
# Deploy using gcloud CLI
```

## Usage

1. **Open the App**
   - Navigate to the app URL (local or cloud)

2. **Configure Settings** (in sidebar)
   - Select the month
   - Select the year

3. **Upload File**
   - Click "Browse files"
   - Select your billing Excel file (.xlsx or .xlsm)

4. **Generate Summary**
   - Click "Generate Summary" button
   - Wait for processing to complete (progress bar will show status)

5. **Download Report**
   - Click "Download Summary Report" button
   - Save the generated Excel file

## Output

The generated Excel file contains a "Summary" sheet with:

### Data Columns:
- Order Date, DU, DU-Order, CM Code, Sold To
- Destination Code, Ship To, Addresses, Province, Post Code
- Area, Total Pick Q'TY, Ship Total WT
- Up 10KG/Chg, Rate/KG, Min/Charge, All Charge
- Pick Date, Transport, Premium, Remark
- Part No., Pick Q'TY, Total Weight

### Summary Calculations:
- Total quantities (CTN)
- Total weight (KG)
- Total charges (BATH)
- Fuel Surcharge (FSC) at 13.62%
- Grand total with FSC

### Formatting:
- Color-coded headers (green)
- Frozen header row
- Properly formatted numbers (#,##0.00)
- Auto-sized columns
- Summary totals highlighted

## Processing Logic

### Weight Calculation:
```
Total Weight = Pick QTY × Part Weight (from Cargo and Weight sheet)
```

### Charge Calculation:
```
Up 10KG/Chg = max(Total Weight - 10, 0)
All Charge = (Up 10KG/Chg × Rate/KG) + Min/Charge
```

### Transport Assignment:
```
Transport = "STL" if Area == "BKK" else "DASH"
```

### Fuel Surcharge:
```
FSC = Total Charges × 13.62%
Grand Total = Total Charges + FSC
```

## Error Handling

The app handles common errors:

- **Part Not Found**: If a part number is not in the Cargo and Weight sheet
- **Post Code Not Found**: If a post code is not in the Sell Price sheet
- **Missing Sheets**: If required sheets are not present in the uploaded file
- **Invalid Data**: If data format is incorrect

## File Size Limits

- **Local**: No practical limit
- **Streamlit Cloud**: 200MB upload limit
- **Heroku**: 300MB slug size limit

For larger files, consider using local installation or upgrading cloud platform plan.

## Troubleshooting

### "Module not found" error
```bash
pip install -r requirements.txt
```

### "File too large" error
- Reduce file size or use local installation
- Remove unnecessary sheets from input file

### "Sheet not found" error
- Ensure all required sheets exist:
  - Daily sheets (1, 2, 3...31)
  - "Cargo and Weight"
  - "Sell Price"

### Processing takes too long
- Normal for files with many orders
- Cloud deployments may timeout for very large files
- Consider local installation for large datasets

## Support

For issues or questions:
1. Check this README
2. Verify input file format matches requirements
3. Test with a smaller sample file first

## Technical Details

- **Framework**: Streamlit
- **Excel Processing**: openpyxl
- **Data Processing**: pandas
- **Python Version**: 3.8+

## License

This application is provided as-is for internal use.

## Version History

- **v1.0** (2024): Initial release
  - Core billing summary functionality
  - Month/year selection
  - Excel upload and download
  - Automatic calculations
  - Cloud deployment ready
