"""
Filepath: kts_analyzer/services/report_service.py
----------------------------------------------
KTS Analyzer - Report Service (VCSM)

**Refactored**
- **Chart Consolidation:** Created a new "All Charts Dashboard" sheet.
  The _create_data_sheet method now ONLY writes data.
- A new method, _add_charts_to_dashboard, now creates the monthly/yearly
  charts and places them all on the new dashboard sheet.
- **Date Formatting:** _create_data_sheet now writes tables manually
  (cell by cell) to apply the 'yyyy-mm' format directly to the
  Date column in the data tables, as requested.
- All correlation analysis logic is retained.
--------------------------------------
"""

import pandas as pd
import re
from io import BytesIO
from openpyxl.utils import get_column_letter

class XlsxReportService:
    """
    Handles the creation of the final Excel report using XlsxWriter
    to generate native, data-linked charts.
    """

    def __init__(self):
        """Initialize the report service."""
        pass

    def _sanitize_sheet_name(self, name: str) -> str:
        """Cleans a string to be a valid Excel sheet name."""
        name = re.sub(r'[\\/*?:\[\]]', '', name)
        return name[:31]

    def generate_report(self, 
                        output_file: str, 
                        all_data_grouped: pd.api.typing.DataFrameGroupBy, 
                        summary_data: pd.DataFrame,
                        analysis_data: pd.DataFrame):
        """
        Generates the full Excel report with data and native charts.
        
        Args:
            output_file: Path to save the new report.
            all_data_grouped: A DataFrameGroupBy object with the time-series data.
            summary_data: A DataFrame with summary statistics.
            analysis_data: A wide-format DataFrame for correlation.
        """
        
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- Define Reusable Formats ---
                header_format = workbook.add_format({
                    'bold': True, 'valign': 'top', 'fg_color': '#4F81BD',
                    'font_color': 'white', 'border': 1, 'text_wrap': True
                })
                # *** REQUESTED FORMAT CHANGE ***
                date_format = workbook.add_format({'num_format': 'yyyy-mm', 'border': 1})
                number_format = workbook.add_format({'num_format': '#,##0.0', 'border': 1})
                year_format = workbook.add_format({'num_format': '0000', 'border': 1})

                # --- 1. New Analysis Sheet (First Sheet) ---
                if not analysis_data.empty:
                    self._create_analysis_sheet(writer, workbook, analysis_data, 
                                                header_format, date_format, number_format)

                # --- 2. Summary Sheet ---
                self._create_summary_sheet(writer, summary_data, header_format)
                
                # --- 3. New "All Charts Dashboard" Sheet ---
                charts_dashboard = workbook.add_worksheet("All Charts Dashboard")
                charts_dashboard.hide_gridlines(2)
                current_dashboard_row = 1 # 1-based index for rows

                # --- 4. Data-Only Sheets & Chart Population ---
                
                # Create a list of processed groups to iterate over
                # We do this because the group-by object is a generator
                processed_groups = []
                for group_key, data in all_data_grouped:
                    if data.empty:
                        continue
                    processed_groups.append((group_key, data))

                # Loop 1: Create all data-only sheets
                for group_key, data in processed_groups:
                    group_key_str = " - ".join(map(str, group_key))
                    sheet_name = self._sanitize_sheet_name(group_key_str)
                    
                    self._create_data_sheet(writer, workbook, sheet_name, group_key, data,
                                            header_format, date_format, number_format, year_format)

                # Loop 2: Populate the charts dashboard
                for group_key, data in processed_groups:
                    group_key_str = " - ".join(map(str, group_key))
                    sheet_name = self._sanitize_sheet_name(group_key_str)

                    # Add charts for this group to the dashboard
                    current_dashboard_row = self._add_charts_to_dashboard(
                        workbook,
                        charts_dashboard,
                        sheet_name,
                        group_key,
                        data,
                        current_dashboard_row
                    )

        except Exception as e:
            raise IOError(f"Failed to save Excel report: {e}")

    def _create_analysis_sheet(self, writer, workbook, analysis_data, 
                               header_format, date_format, number_format):
        """Creates the new main dashboard for correlation analysis."""
        
        sheet_name = "Correlation Analysis"
        
        # Write the wide-format data
        analysis_data.to_excel(writer, sheet_name=sheet_name, index=True, startrow=1)
        worksheet = writer.sheets[sheet_name]
        worksheet.hide_gridlines(2)
        
        # Format Header
        for col_num, value in enumerate(analysis_data.columns.values):
             worksheet.write(1, col_num + 1, value, header_format)
        worksheet.write(1, 0, analysis_data.index.name, header_format) # Format index header

        # Format Columns
        worksheet.set_column('A:A', 12, date_format) # Date Index
        worksheet.set_column('B:Z', 18, number_format) # All data columns
        worksheet.set_row(1, 40) # Taller header row for wrapped text
        
        # --- Add Correlation Charts ---
        data_rows = len(analysis_data)
        
        # Find column indices dynamically
        try:
            cols = list(analysis_data.columns)
            # Get index (1-based) for header, +1 for 0-based list
            col_date = 'A'
            col_fleet = get_column_letter(cols.index('Active Fleet Count (Aprox)') + 2)
            col_fuel = get_column_letter(cols.index('Liter of Diesel Consumed') + 2)
            col_material = get_column_letter(cols.index('Total Material (kt)') + 2)
            col_efficiency = get_column_letter(cols.index('Efficiency (kt per Liter)') + 2)
            col_productivity = get_column_letter(cols.index('Productivity (kt per Fleet)') + 2)
            
            chart_start_row = 3 # Start charts below header
            
            # Chart 1: Total Material vs. Active Fleet
            chart1 = self._add_correlation_chart(
                workbook=workbook,
                title='Total Material Mined vs. Active Fleet',
                cat_formula=f"='{sheet_name}'!${col_date}$3:${col_date}${2 + data_rows}",
                val1_formula=f"='{sheet_name}'!${col_material}$3:${col_material}${2 + data_rows}",
                val1_name='Total Material (kt)',
                val1_axis_name='Total Material (kt)',
                val2_formula=f"='{sheet_name}'!${col_fleet}$3:${col_fleet}${2 + data_rows}",
                val2_name='Active Fleet',
                val2_axis_name='Active Fleet Count'
            )
            worksheet.insert_chart(chart_start_row, 0, chart1, {'x_scale': 1.8, 'y_scale': 1.4})

            # Chart 2: Fuel Consumed vs. Active Fleet
            chart2 = self._add_correlation_chart(
                workbook=workbook,
                title='Fuel Consumed vs. Active Fleet',
                cat_formula=f"='{sheet_name}'!${col_date}$3:${col_date}${2 + data_rows}",
                val1_formula=f"='{sheet_name}'!${col_fuel}$3:${col_fuel}${2 + data_rows}",
                val1_name='Liters Consumed',
                val1_axis_name='Liters Consumed',
                val2_formula=f"='{sheet_name}'!${col_fleet}$3:${col_fleet}${2 + data_rows}",
                val2_name='Active Fleet',
                val2_axis_name='Active Fleet Count'
            )
            worksheet.insert_chart(chart_start_row + 25, 0, chart2, {'x_scale': 1.8, 'y_scale': 1.4})

            # Chart 3: Efficiency (kt per Liter)
            chart3 = self._add_correlation_chart(
                workbook=workbook,
                title='Efficiency (Total kt Mined per Liter Fuel)',
                cat_formula=f"='{sheet_name}'!${col_date}$3:${col_date}${2 + data_rows}",
                val1_formula=f"='{sheet_name}'!${col_efficiency}$3:${col_efficiency}${2 + data_rows}",
                val1_name='Efficiency (kt/L)',
                val1_axis_name='kt / Liter',
                val2_formula=None # Single-axis chart
            )
            worksheet.insert_chart(chart_start_row, 8, chart3, {'x_scale': 1.8, 'y_scale': 1.4})
            
            # Chart 4: Productivity (kt per Fleet)
            chart4 = self._add_correlation_chart(
                workbook=workbook,
                title='Productivity (Total kt Mined per Fleet)',
                cat_formula=f"='{sheet_name}'!${col_date}$3:${col_date}${2 + data_rows}",
                val1_formula=f"='{sheet_name}'!${col_productivity}$3:${col_productivity}${2 + data_rows}",
                val1_name='Productivity (kt/Fleet)',
                val1_axis_name='kt / Fleet Unit',
                val2_formula=None # Single-axis chart
            )
            worksheet.insert_chart(chart_start_row + 25, 8, chart4, {'x_scale': 1.8, 'y_scale': 1.4})

        except ValueError as e:
            # Handle if a key column isn't found
            worksheet.write('A1', f"Could not generate charts: Required column missing ({e}).", 
                            workbook.add_format({'bold': True, 'font_color': 'red'}))
        except Exception as e:
            worksheet.write('A1', f"An error occurred generating charts: {e}", 
                            workbook.add_format({'bold': True, 'font_color': 'red'}))

    def _add_correlation_chart(self, workbook, title, cat_formula, val1_formula, 
                               val1_name, val1_axis_name, val2_formula, 
                               val2_name=None, val2_axis_name=None):
        """Helper function to create a 1 or 2 axis line chart."""
        
        chart = workbook.add_chart({'type': 'line'})

        # Series 1 (Primary Y-Axis)
        chart.add_series({
            'categories': cat_formula,
            'values':     val1_formula,
            'name':       val1_name,
            'marker':     {'type': 'automatic'},
            'y_axis':     0,
        })
        
        # Series 2 (Secondary Y-Axis), if provided
        if val2_formula:
            chart.add_series({
                'categories': cat_formula,
                'values':     val2_formula,
                'name':       val2_name,
                'marker':     {'type': 'automatic'},
                'y_axis':     1, # Use secondary axis
            })
        
        chart.set_title({'name': title})
        chart.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        
        # Set Y-Axis (Primary)
        chart.set_y_axis({'name': val1_axis_name, 'num_format': '#,##0'})
        
        # Set Y2-Axis (Secondary), if provided
        if val2_axis_name:
            chart.set_y2_axis({'name': val2_axis_name, 'num_format': '#,##0'})
            chart.set_legend({'position': 'bottom'})
        else:
            chart.set_legend({'position': 'none'})
            
        return chart

    def _create_summary_sheet(self, writer, summary_data, header_format):
        """Writes and formats the summary data to its own sheet."""
        sheet_name = "Summary"
        summary_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        worksheet = writer.sheets[sheet_name]
        worksheet.hide_gridlines(2)
        
        for col_num, value in enumerate(summary_data.columns.values):
            worksheet.write(1, col_num, value, header_format)
            
        for i, col in enumerate(summary_data.columns):
            width = max(len(str(col)), summary_data[col].astype(str).map(len).max())
            worksheet.set_column(i, i, width + 2)

    def _create_data_sheet(self, writer, workbook, sheet_name, group_key, data,
                             header_format, date_format, number_format, year_format):
        """
        Creates a single sheet with time-series data AND formats the date column.
        This method NO LONGER creates charts.
        """
        
        monthly_data = data[['Value']].copy()
        yearly_data = monthly_data['Value'].resample('YE').sum().to_frame(name='Total')
        yearly_data.index = yearly_data.index.year
        yearly_data.index.name = "Year"

        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet 
        
        title = f"Data: {group_key[0]} - {group_key[1]} ({group_key[2]})"
        worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))

        # --- Manually write Monthly Data to apply 'yyyy-mm' format ---
        worksheet.write('A3', "Monthly Data", header_format)
        worksheet.merge_range('A3:B3', "Monthly Data", header_format)
        
        worksheet.write('A4', "Date", header_format)
        worksheet.write('B4', "Value", header_format)
        
        row = 4 # Start data on row 5 (0-indexed row 4)
        for date, value in monthly_data['Value'].items():
            worksheet.write_datetime(row, 0, date, date_format)
            worksheet.write_number(row, 1, value, number_format)
            row += 1
        
        # --- Manually write Yearly Data ---
        worksheet.write('D3', "Yearly Summary", header_format)
        worksheet.merge_range('D3:E3', "Yearly Summary", header_format)

        worksheet.write('D4', "Year", header_format)
        worksheet.write('E4', "Total", header_format)
        
        row = 4 # Start data on row 5
        for year, total in yearly_data['Total'].items():
            worksheet.write_number(row, 3, year, year_format)
            worksheet.write_number(row, 4, total, number_format)
            row += 1

        # --- Set Column Widths ---
        worksheet.set_column('A:A', 12) # yyyy-mm
        worksheet.set_column('B:B', 15) # Number
        worksheet.set_column('D:D', 10) # Year
        worksheet.set_column('E:E', 15) # Number

    def _add_charts_to_dashboard(self, workbook, dashboard_sheet, data_sheet_name,
                                 group_key, data, current_row):
        """
        Creates monthly and yearly charts for a data group and adds them
        to the specified dashboard sheet.
        
        Args:
            workbook: The XlsxWriter workbook object.
            dashboard_sheet: The worksheet object for the dashboard.
            data_sheet_name: The name of the sheet containing the data (e.g., 'Ore Mined - RGM - kt').
            group_key: The tuple key for the data (Category, SubCategory, Unit).
            data: The DataFrame (used to get row counts).
            current_row: The row on the dashboard to start inserting charts.
            
        Returns:
            The new current_row for the next set of charts.
        """
        
        num_monthly = len(data)
        num_yearly = data['Value'].resample('YE').sum().count()
        
        # --- Create Title for this Chart Group ---
        title = f"Analysis: {group_key[0]} - {group_key[1]} ({group_key[2]})"
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'bottom': 1, 'font_color': '#4F81BD'})
        dashboard_sheet.merge_range(current_row, 0, current_row, 15, title, title_format)
        current_row += 1 # Move down past title

        # --- Monthly Chart ---
        quoted_sheet_name = f"'{data_sheet_name}'"
        chart_monthly = workbook.add_chart({'type': 'line'})
        
        # Note: Data starts at row 5 (A5, B5) because of manual write
        cat_monthly = f'={quoted_sheet_name}!$A$5:$A${4 + num_monthly}'
        val_monthly = f'={quoted_sheet_name}!$B$5:$B${4 + num_monthly}'
        
        chart_monthly.add_series({
            'categories': cat_monthly,
            'values':     val_monthly,
            'name':       f'Monthly {group_key[2]}',
            'marker':     {'type': 'automatic'},
        })
        
        chart_monthly.set_title({'name': f'Monthly: {group_key[0]} - {group_key[1]}'})
        chart_monthly.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart_monthly.set_y_axis({'name': group_key[2], 'num_format': '#,##0'})
        chart_monthly.set_legend({'position': 'none'})
        
        dashboard_sheet.insert_chart(current_row, 0, chart_monthly, {'x_scale': 1.8, 'y_scale': 1.4})

        # --- Yearly Chart ---
        chart_yearly = workbook.add_chart({'type': 'column'})
        
        # Note: Data starts at row 5 (D5, E5)
        cat_yearly = f'={quoted_sheet_name}!$D$5:$D${4 + num_yearly}'
        val_yearly = f'={quoted_sheet_name}!$E$5:$E${4 + num_yearly}'

        chart_yearly.add_series({
            'categories': cat_yearly,
            'values':     val_yearly,
            'name':       f'Yearly Total {group_key[2]}',
            'data_labels': {'value': True, 'num_format': '#,##0'},
        })
        
        chart_yearly.set_title({'name': f'Yearly Summary: {group_key[0]} - {group_key[1]}'})
        chart_yearly.set_x_axis({'name': 'Year'})
        chart_yearly.set_y_axis({'name': f'Total {group_key[2]}', 'num_format': '#,##0'})
        chart_yearly.set_legend({'position': 'none'})

        dashboard_sheet.insert_chart(current_row, 8, chart_yearly, {'x_scale': 1.8, 'y_scale': 1.4})

        # Return the next available row, adding 25 rows for chart height + 2 for spacing
        return current_row + 27

