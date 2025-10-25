# Mining Data Analysis Suite - Fixed Version

## ✅ Solution to Date Parsing Error

The error "NaTType does not support strftime" has been fixed by properly handling date parsing for Spanish month abbreviations and invalid dates. The new scripts create charts directly in Excel sheets as requested.

## 📁 Files Provided

### Main Solutions:

1. **`mining_chart_generator.py`** - Command-line tool that creates Excel with embedded charts
2. **`mining_analyzer_gui.py`** - User-friendly GUI version with file browser
3. **`mining_excel_analyzer.py`** - Alternative version with data tables and basic charts

### Supporting Files:

- **`requirements.txt`** - Python package dependencies
- **`sample_mining_data.xlsx`** - Sample data for testing
- **`sample_mining_data_with_charts.xlsx`** - Example output with charts

## 🚀 Quick Start

### Installation:
```bash
pip install pandas matplotlib numpy openpyxl xlrd
```

### Option 1: GUI Version (Easiest)
```bash
python mining_analyzer_gui.py
```
1. Click "Browse Excel File" to select your data
2. Click "Generate Charts in Excel"
3. Charts will be embedded in a new Excel file

### Option 2: Command Line
```bash
python mining_chart_generator.py your_data.xlsx
```
Or interactive mode:
```bash
python mining_chart_generator.py
# Then enter file path when prompted
```

### Option 3: With Sheet Name
```bash
python mining_chart_generator.py your_data.xlsx "Sheet Name"
```

## 📊 Charts Generated

The system creates 4 comprehensive chart sets embedded in Excel:

### 1. Production Overview
- Ore Production Comparison (RGM vs Sar)
- Overburden Movement Trends
- Total Material Movement (Stacked Bar)
- Stripping Ratio Analysis

### 2. Efficiency Analysis
- Fleet Utilization vs Production
- Diesel Consumption Trends
- Productivity per Fleet Unit
- Fuel Efficiency (Liters per kt)

### 3. Comparative Analysis
- Production Share (Pie Charts)
- Monthly Ore Comparison (Bar Charts)
- Overburden Share Distribution
- Performance Metrics Comparison

### 4. Trend Analysis
- Moving Averages (3 & 6 months)
- Year-over-Year Comparisons
- Fleet vs Diesel Correlation
- Monthly Seasonality Patterns

## 📋 Excel Output Structure

The generated Excel file contains:
- **Summary** - Key statistics and metrics overview
- **Processed_Data** - Clean, formatted data table
- **Production Overview** - Production charts (embedded image)
- **Efficiency Analysis** - Efficiency charts (embedded image)
- **Comparative Analysis** - Comparison charts (embedded image)
- **Trend Analysis** - Trend charts (embedded image)

## 🔧 Data Format Requirements

Your Excel file should have this structure:

| Metric | Category | Unit | ene-20 | feb-20 | mar-20 | ... |
|--------|----------|------|--------|--------|--------|-----|
| Ore Mined | RGM | kt | 406.8 | 549.1 | 805.9 | ... |
| Overburden | RGM | kt | 4273.2 | 4272.4 | 3235.0 | ... |
| Ore Mined | Sar | kt | 91.3 | 3.1 | 377.8 | ... |
| Overburden | Sar | kt | 445.5 | 731.9 | 669.0 | ... |
| Active Fleet Count (Aprox) | | | 831 | 848 | 865 | ... |
| Liter of Diesel Consumed | | | 5145492 | 5190595 | 5319575 | ... |

### Date Format:
- Spanish abbreviations: ene, feb, mar, abr, may, jun, jul, ago, sep, oct, nov, dic
- Year format: YY or YYYY (e.g., "20" or "2020")
- Full format: "mon-YY" (e.g., "ene-20", "feb-21")

## ✨ Key Features

### Error Handling:
- ✅ Fixed NaTType date parsing errors
- ✅ Handles missing data gracefully
- ✅ Supports Spanish and English month names
- ✅ Automatic data cleaning and validation

### Flexibility:
- Works with any sheet in the Excel file
- Auto-detects data structure
- Handles varying date formats
- Creates charts even with partial data

### Professional Output:
- High-quality chart images embedded in Excel
- Formatted tables with proper styling
- Summary statistics automatically calculated
- Color-coded visualizations for clarity

## 🐛 Troubleshooting

### "NaTType does not support strftime" - FIXED
This error has been resolved by:
- Proper date parsing with error handling
- Support for Spanish month abbreviations
- Filtering out invalid dates before processing

### Charts not appearing:
- Ensure data columns contain numeric values
- Check date format matches expected pattern
- Verify at least one complete data series exists

### File not found:
- Use full path or place file in same directory
- Check file extension (.xlsx or .xls)
- Ensure file is not open in Excel

### Missing data:
- Script handles missing values by treating them as 0
- Partial data will still generate charts
- Check column headers match expected format

## 📈 Example Output

After processing, you'll get an Excel file with:
1. **Summary sheet** with key statistics
2. **Data sheets** with processed information
3. **Chart sheets** with professional visualizations
4. All charts embedded as images for compatibility

## 💡 Tips

1. **Best Results**: Ensure your data has consistent formatting
2. **Large Files**: Processing may take 10-30 seconds
3. **Custom Names**: Use the GUI to specify custom output filenames
4. **Multiple Sheets**: Specify sheet name if file has multiple sheets

## 🔄 Updates from Original Version

### What's Fixed:
- ✅ Date parsing errors (NaTType issue)
- ✅ Charts now embedded directly in Excel
- ✅ Spanish month abbreviation support
- ✅ Better error handling and validation

### What's New:
- 📊 Charts as embedded images in Excel sheets
- 🎯 Summary statistics sheet
- 🖱️ Simple GUI option available
- 📈 More robust data processing

## 📞 Usage Examples

### Basic:
```bash
python mining_chart_generator.py data.xlsx
```

### With options:
```bash
python mining_chart_generator.py data.xlsx "Mining_Data"
```

### GUI (recommended for ease):
```bash
python mining_analyzer_gui.py
# Then use the interface to select file and generate charts
```

## ✅ Success Indicators

You'll know it worked when you see:
- "✅ SUCCESS! Charts created and embedded in Excel"
- New file created: `yourfile_with_charts.xlsx`
- Multiple sheets in the output Excel file
- Charts visible when opening in Excel

---

**Version**: 2.0.0 (Fixed)  
**Last Updated**: October 2025  
**Status**: ✅ Working - Date parsing error resolved
