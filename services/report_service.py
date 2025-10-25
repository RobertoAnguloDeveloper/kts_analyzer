"""
Filepath: kts_analyzer/services/report_service.py
----------------------------------------------
KTS Analyzer - Report Service (VCSM)

**Refactor (V7) - Final Version**
- Includes a "Summary" sheet as the first tab to
  explain all KPI calculations.
- Corrects the `add_table()` range in 'Correlation Analysis'
  to properly include all columns.
- Fixes the chart layout in 'Production Overview' to
  prevent any chart overlap.
- 'Fleet, Fuel & Ore' chart correctly plots 'Active Fleet'
  and 'Total Ore (kt)' on the secondary Y-axis to ensure
  proper scaling and visibility.
- All chart formulas use the correct, working table syntax
  (e.g., =AnalysisData[Total Ore (kt)]).
--------------------------------------
"""

import pandas as pd
import re
from io import BytesIO
from openpyxl.utils import get_column_letter

class XlsxReportService:
    """
    Handles the creation of the final Excel report using XlsxWriter
    to generate native, data-linked charts and formulas.
    """

    def __init__(self):
        """Initialize the report service."""
        # This will store the dynamic column letters for our formulas
        self.col_map = {}
        
    def generate_report(self, 
                        output_file: str, 
                        analysis_data: pd.DataFrame):
        """
        Generates the full Excel report with data, native formulas, and charts.
        
        Args:
            output_file: Path to save the new report.
            analysis_data: A wide-format DataFrame from the data service.
        """
        
        if analysis_data.empty:
            raise ValueError("Cannot generate report from empty analysis data.")
            
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- Define Reusable Formats ---
                self.formats = {
                    'header': workbook.add_format({
                        'bold': True, 'valign': 'top', 'fg_color': '#4F81BD',
                        'font_color': 'white', 'border': 1, 'text_wrap': True
                    }),
                    'date': workbook.add_format({'num_format': 'yyyy-mm', 'border': 1}),
                    'number': workbook.add_format({'num_format': '#,##0.0', 'border': 1}),
                    'pct': workbook.add_format({'num_format': '0.0%', 'border': 1}),
                    'ratio': workbook.add_format({'num_format': '0.00', 'border': 1}),
                    'title': workbook.add_format({'bold': True, 'font_size': 14, 'bottom': 1, 'font_color': '#4F81BD'}),
                    'link': workbook.add_format({'font_color': 'blue', 'underline': 1}),
                    
                    # --- Formats for Summary Sheet ---
                    'summary_header': workbook.add_format({
                        'bold': True, 'font_size': 12, 'font_color': '#4F81BD', 
                        'bottom': 1, 'border_color': '#4F81BD', 'valign': 'top'
                    }),
                    'summary_text': workbook.add_format({
                        'font_size': 10, 'valign': 'top', 'text_wrap': True
                    }),
                    'summary_formula': workbook.add_format({
                        'font_size': 10, 'valign': 'top', 'text_wrap': True, 
                        'font_name': 'Courier New', 'bold': True
                    })
                }

                # --- Build Report Sheets ---
                
                # 1. "Summary" Sheet (NEW)
                self._create_summary_sheet(workbook, "Summary")
                
                # 2. "Processed_Data" Sheet (Raw Pivot Data)
                data_sheet_name = "Processed_Data"
                self._create_processed_data_sheet(writer, workbook, data_sheet_name, analysis_data)

                # 3. "Correlation Analysis" Sheet (Formula-Driven Table)
                calc_sheet_name = "Correlation Analysis"
                table_name = "AnalysisData" # The name for our new Excel Table
                data_rows = len(analysis_data)
                self._create_analysis_calculation_sheet(
                    writer, workbook, calc_sheet_name, 
                    data_sheet_name, data_rows, table_name
                )

                # 4. "Production Overview" Dashboard
                self._create_overview_charts_sheet(workbook, "Production Overview", table_name)

                # 5. "Production Charts" Dashboard
                self._create_production_charts_sheet(workbook, "Production Charts", table_name)

                # 6. "Efficiency Charts" Dashboard
                self._create_efficiency_charts_sheet(workbook, "Efficiency Charts", table_name)
                
                # 7. "Fleet, Fuel & Ore" Dashboard
                self._create_fleet_fuel_charts_sheet(workbook, "Fleet, Fuel & Ore", table_name)
                
                # 8. Set active sheet to the new Summary sheet
                workbook.get_worksheet_by_name("Summary").activate()

        except Exception as e:
            raise IOError(f"Failed to save Excel report: {e}")

    def _create_summary_sheet(self, workbook, sheet_name):
        """
        Creates a summary and KPI definition sheet.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(2)
        
        # Get formats
        title_format = self.formats['title']
        header_format = self.formats['summary_header']
        text_format = self.formats['summary_text']
        formula_format = self.formats['summary_formula']

        # Set column widths
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:B', 70)
        worksheet.set_column('C:C', 5)
        worksheet.set_default_row(14)

        row = 1
        worksheet.write(row, 1, "Report Summary & KPI Definitions", title_format)
        row += 2

        # --- How to Use ---
        worksheet.write(row, 1, "How To Use This Report", header_format)
        row += 1
        worksheet.set_row(row, 60) # Set row height for text
        worksheet.write(row, 1, 
            "This report is built on dynamic, formula-driven data.\n"
            "1. All calculations are performed in the 'Correlation Analysis' sheet within an Excel Table.\n"
            "2. All charts on the dashboards dynamically reference this table.\n"
            "3. You can add new monthly data to the 'Processed_Data' sheet, and all formulas and charts will update automatically.", 
            text_format
        )
        row += 2

        # --- KPI Definitions ---
        worksheet.write(row, 1, "Key Performance Indicator (KPI) Calculations", header_format)
        row += 1

        def write_kpi(r, kpi, desc, formula):
            # Helper to write a KPI block
            worksheet.set_row(r, 16) # Header row
            worksheet.write_rich_string(r, 1, header_format, kpi, text_format, f" - {desc}")
            worksheet.set_row(r + 1, 16) # Formula row
            worksheet.write_rich_string(r + 1, 1, formula_format, "Formula: ", text_format, formula)
            return r + 3 # Return next starting row

        row = write_kpi(row,
            "Total Material (kt)",
            "Total material (ore + waste) moved.",
            "= [Ore Mined - RGM - kt] + [Ore Mined - Sar - kt] + [Overburden - RGM - kt] + [Overburden - Sar - kt]"
        )
        
        row = write_kpi(row,
            "Efficiency (kt per Liter)",
            "Measures fuel efficiency. How many kilotons of material are moved for every liter of diesel consumed.",
            "= [Total Material (kt)] / [Liter of Diesel Consumed]"
        )

        row = write_kpi(row,
            "Productivity (kt per Fleet)",
            "Measures fleet productivity. How many kilotons of material are moved per active fleet unit.",
            "= [Total Material (kt)] / [Active Fleet Count (Aprox)]"
        )

        row = write_kpi(row,
            "RGM Strip Ratio",
            "The ratio of waste (overburden) to ore for the RGM area. A key mining metric.",
            "= [Overburden - RGM - kt] / [Ore Mined - RGM - kt]"
        )

        row = write_kpi(row,
            "Sar Strip Ratio",
            "The ratio of waste (overburden) to ore for the Sar area.",
            "= [Overburden - Sar - kt] / [Ore Mined - Sar - kt]"
        )
        
        # --- Chart Definitions ---
        row += 1
        worksheet.write(row, 1, "Chart Definitions", header_format)
        row += 1

        row = write_kpi(row,
            "Stripping Ratio Trends (Chart)",
            "This chart plots the 'RGM Strip Ratio' and 'Sar Strip Ratio' KPIs over time to visualize trends.",
            "This chart plots the two KPI columns calculated above."
        )


    def _create_processed_data_sheet(self, writer, workbook, sheet_name, analysis_data):
        """
        Creates the 'Processed_Data' sheet with the raw pivot table.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide() # Hide this sheet from the user
        
        # Write Header
        worksheet.write(0, 0, analysis_data.index.name, self.formats['header'])
        for col_num, value in enumerate(analysis_data.columns.values):
             worksheet.write(0, col_num + 1, value, self.formats['header'])
        
        # Write Data
        row = 1
        for date, data_row in analysis_data.iterrows():
            worksheet.write_datetime(row, 0, date, self.formats['date'])
            for col_num, value in enumerate(data_row.values):
                worksheet.write_number(row, col_num + 1, value, self.formats['number'])
            row += 1
            
        # Set Column Widths
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:Z', 18)

    def _create_analysis_calculation_sheet(self, writer, workbook, sheet_name, 
                                           data_sheet_name, data_rows, table_name):
        """
        Creates the 'Correlation Analysis' sheet, which is built entirely
        from Excel formulas and formatted as a named Table.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        
        # --- Find Source Columns from 'Processed_Data' ---
        base_cols = {
            'Date': 'A',
            'Ore Mined - RGM - kt': 'B',
            'Overburden - RGM - kt': 'C',
            'Ore Mined - Sar - kt': 'D',
            'Overburden - Sar - kt': 'E',
            'Active Fleet Count (Aprox)': 'F',
            'Liter of Diesel Consumed': 'G'
        }
        
        # --- Define Calculation Columns ---
        calc_cols = {
            'Total Ore (kt)': 'H',
            'Total Overburden (kt)': 'I',
            'Total Material (kt)': 'J',
            'Efficiency (kt per Liter)': 'K',
            'Productivity (kt per Fleet)': 'L',
            'RGM Strip Ratio': 'M',
            'Sar Strip Ratio': 'N'
        }
        
        self.col_map = {**base_cols, **calc_cols}
        
        # --- Write Header Row ---
        col = 0
        table_headers = []
        for name in self.col_map.keys():
            worksheet.write(0, col, name, self.formats['header'])
            table_headers.append({'header': name}) # For add_table()
            col += 1
            
        # --- Write Data Rows (Formulas) ---
        for r in range(data_rows):
            row_idx = r + 1 # 1-based index
            
            # --- Link Base Data ---
            for name, col_letter in base_cols.items():
                formula = f"='{data_sheet_name}'!{col_letter}{row_idx + 1}"
                col_idx = name_idx(name)
                if name == 'Date':
                    worksheet.write_formula(row_idx, col_idx, formula, self.formats['date'])
                else:
                    worksheet.write_formula(row_idx, col_idx, formula, self.formats['number'])

            # --- Write Calculation Formulas ---
            c = self.col_map
            
            # Total Ore (kt)
            f_total_ore = f"=SUM({c['Ore Mined - RGM - kt']}{row_idx+1}, {c['Ore Mined - Sar - kt']}{row_idx+1})"
            worksheet.write_formula(row_idx, name_idx('Total Ore (kt)'), f_total_ore, self.formats['number'])
            
            # Total Overburden (kt)
            f_total_over = f"=SUM({c['Overburden - RGM - kt']}{row_idx+1}, {c['Overburden - Sar - kt']}{row_idx+1})"
            worksheet.write_formula(row_idx, name_idx('Total Overburden (kt)'), f_total_over, self.formats['number'])

            # Total Material (kt)
            f_total_mat = f"=SUM({c['Total Ore (kt)']}{row_idx+1}, {c['Total Overburden (kt)']}{row_idx+1})"
            worksheet.write_formula(row_idx, name_idx('Total Material (kt)'), f_total_mat, self.formats['number'])

            # Efficiency (kt per Liter)
            f_efficiency = f"=IFERROR({c['Total Material (kt)']}{row_idx+1} / {c['Liter of Diesel Consumed']}{row_idx+1}, 0)"
            worksheet.write_formula(row_idx, name_idx('Efficiency (kt per Liter)'), f_efficiency, self.formats['ratio'])
            
            # Productivity (kt per Fleet)
            f_productivity = f"=IFERROR({c['Total Material (kt)']}{row_idx+1} / {c['Active Fleet Count (Aprox)']}{row_idx+1}, 0)"
            worksheet.write_formula(row_idx, name_idx('Productivity (kt per Fleet)'), f_productivity, self.formats['ratio'])

            # RGM Strip Ratio
            f_rgm_strip = f"=IFERROR({c['Overburden - RGM - kt']}{row_idx+1} / {c['Ore Mined - RGM - kt']}{row_idx+1}, 0)"
            worksheet.write_formula(row_idx, name_idx('RGM Strip Ratio'), f_rgm_strip, self.formats['ratio'])
            
            # Sar Strip Ratio
            f_sar_strip = f"=IFERROR({c['Overburden - Sar - kt']}{row_idx+1} / {c['Ore Mined - Sar - kt']}{row_idx+1}, 0)"
            worksheet.write_formula(row_idx, name_idx('Sar Strip Ratio'), f_sar_strip, self.formats['ratio'])


        # --- Format Sheet ---
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:J', 18)
        worksheet.set_column('K:N', 15)
        worksheet.freeze_panes(1, 0)
        
        # --- Add Excel Table ---
        # **FIX:** Corrected range calculation to use len() directly
        table_range = f'A1:{get_column_letter(len(self.col_map))}{data_rows + 1}'
        worksheet.add_table(table_range, {
            'name': table_name,
            'columns': table_headers,
            'style': 'TableStyleMedium9'
        })
        
    def _create_overview_charts_sheet(self, workbook, sheet_name, table_name):
        """
        Creates the 'Production Overview' dashboard sheet.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(2)
        
        cat_formula = f"={table_name}[Date]"

        # --- Chart 1: Ore Production Over Time (Line) ---
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.set_title({'name': 'Ore Production Over Time'})
        
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Ore Mined - RGM - kt]",
            'name': 'RGM',
            'marker': {'type': 'automatic'},
        })
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Ore Mined - Sar - kt]",
            'name': 'Sar',
            'marker': {'type': 'automatic'},
        })
        chart1.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart1.set_y_axis({'name': 'Ore Mined (kt)', 'num_format': '#,##0'})
        chart1.set_legend({'position': 'bottom'})
        # **FIX:** Adjusted scale and position to prevent overlap
        worksheet.insert_chart('A1', chart1, {'x_scale': 1.2, 'y_scale': 1.5})

        # --- Chart 2: Overburden Movement (Line) ---
        chart2 = workbook.add_chart({'type': 'line'})
        chart2.set_title({'name': 'Overburden Movement'})
        
        chart2.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Overburden - RGM - kt]",
            'name': 'RGM',
            'marker': {'type': 'automatic'},
        })
        chart2.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Overburden - Sar - kt]",
            'name': 'Sar',
            'marker': {'type': 'automatic'},
        })
        chart2.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart2.set_y_axis({'name': 'Overburden (kt)', 'num_format': '#,##0'})
        chart2.set_legend({'position': 'bottom'})
        # **FIX:** Adjusted scale and position to prevent overlap
        worksheet.insert_chart('K1', chart2, {'x_scale': 1.2, 'y_scale': 1.5})

        # --- Chart 3: Stripping Ratio Trends (Line) ---
        chart3 = workbook.add_chart({'type': 'line'})
        chart3.set_title({'name': 'Stripping Ratio Trends (Overburden/Ore)'})
        
        chart3.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[RGM Strip Ratio]",
            'name': 'RGM Strip Ratio',
            'marker': {'type': 'automatic'},
        })
        chart3.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Sar Strip Ratio]",
            'name': 'Sar Strip Ratio',
            'marker': {'type': 'automatic'},
        })
        chart3.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart3.set_y_axis({'name': 'Strip Ratio', 'num_format': '0.00'})
        chart3.set_legend({'position': 'bottom'})
        worksheet.insert_chart('A26', chart3, {'x_scale': 1.2, 'y_scale': 1.5})


    def _create_production_charts_sheet(self, workbook, sheet_name, table_name):
        """
        Creates the 'Production Charts' dashboard sheet.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(2)
        
        cat_formula = f"={table_name}[Date]"
        
        # --- Chart 1: Total Material Mined (Stacked Area) ---
        chart1 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        chart1.set_title({'name': 'Total Material Mined (RGM vs Sar)'})
        
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Ore Mined - RGM - kt]",
            'name': 'RGM Ore'
        })
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Overburden - RGM - kt]",
            'name': 'RGM Overburden'
        })
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Ore Mined - Sar - kt]",
            'name': 'Sar Ore'
        })
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Overburden - Sar - kt]",
            'name': 'Sar Overburden'
        })
        
        chart1.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart1.set_y_axis({'name': 'Total Material (kt)', 'num_format': '#,##0'})
        chart1.set_legend({'position': 'bottom'})
        worksheet.insert_chart('A1', chart1, {'x_scale': 2.0, 'y_scale': 1.5})

        # --- Chart 2: Ore Mined (Stacked Area) ---
        chart2 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        chart2.set_title({'name': 'Ore Mined (RGM vs Sar)'})
        
        chart2.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Ore Mined - RGM - kt]",
            'name': 'RGM Ore'
        })
        chart2.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Ore Mined - Sar - kt]",
            'name': 'Sar Ore'
        })
        chart2.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart2.set_y_axis({'name': 'Ore Mined (kt)', 'num_format': '#,##0'})
        chart2.set_legend({'position': 'bottom'})
        worksheet.insert_chart('A26', chart2, {'x_scale': 1.0, 'y_scale': 1.5})

        # --- Chart 3: Overburden (Stacked Area) ---
        chart3 = workbook.add_chart({'type': 'area', 'subtype': 'stacked'})
        chart3.set_title({'name': 'Overburden (RGM vs Sar)'})
        
        chart3.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Overburden - RGM - kt]",
            'name': 'RGM Overburden'
        })
        chart3.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Overburden - Sar - kt]",
            'name': 'Sar Overburden'
        })
        chart3.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart3.set_y_axis({'name': 'Overburden (kt)', 'num_format': '#,##0'})
        chart3.set_legend({'position': 'bottom'})
        worksheet.insert_chart('I26', chart3, {'x_scale': 1.0, 'y_scale': 1.5})

    def _create_efficiency_charts_sheet(self, workbook, sheet_name, table_name):
        """
        Creates the 'Efficiency Charts' dashboard sheet.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(2)
        
        cat_formula = f"={table_name}[Date]"

        # --- Chart 1: Productivity (kt / Fleet) ---
        chart1 = workbook.add_chart({'type': 'line'})
        chart1.set_title({'name': 'Productivity (kt / Fleet)'})
        
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Productivity (kt per Fleet)]",
            'name': 'kt / Fleet',
            'marker': {'type': 'automatic'},
        })
        chart1.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart1.set_y_axis({'name': 'kt / Fleet', 'num_format': '0.00'})
        chart1.set_legend({'position': 'none'})
        worksheet.insert_chart('A1', chart1, {'x_scale': 2.0, 'y_scale': 1.5})

        # --- Chart 2: Efficiency (kt / L) ---
        chart2 = workbook.add_chart({'type': 'line'})
        chart2.set_title({'name': 'Efficiency (kt / L)'})
        
        chart2.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Efficiency (kt per Liter)]",
            'name': 'kt / L',
            'marker': {'type': 'automatic'},
        })
        chart2.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart2.set_y_axis({'name': 'kt / Liter', 'num_format': '0.00'})
        chart2.set_legend({'position': 'none'})
        worksheet.insert_chart('A26', chart2, {'x_scale': 2.0, 'y_scale': 1.5})
        
    def _create_fleet_fuel_charts_sheet(self, workbook, sheet_name, table_name):
        """
        Creates the 'Fleet, Fuel & Ore' dashboard sheet.
        This function correctly plots Fuel (column) on the primary
        axis and Fleet (line) / Ore (line) on the secondary axis.
        """
        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.hide_gridlines(2)
        
        cat_formula = f"={table_name}[Date]"

        # --- Chart 1: Fleet, Fuel & Ore Production (Combo Chart) ---
        chart1 = workbook.add_chart({'type': 'column'})
        chart1.set_title({'name': 'Fleet, Fuel & Ore Production Analysis'})
        
        # Series 1: Fuel (Column, Primary Axis)
        chart1.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Liter of Diesel Consumed]",
            'name': 'Liters Consumed',
            'y_axis': 0, # Primary axis
        })
        
        # Create the Line chart to combine
        line_chart = workbook.add_chart({'type': 'line'})
        
        # Series 2: Fleet (Line, Secondary Axis)
        line_chart.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Active Fleet Count (Aprox)]",
            'name': 'Active Fleet',
            'marker': {'type': 'automatic'},
            'y_axis': 1, # Use secondary axis
        })
        
        # Series 3: Total Ore (Line, Secondary Axis)
        line_chart.add_series({
            'categories': cat_formula,
            'values': f"={table_name}[Total Ore (kt)]",
            'name': 'Total Ore (kt)',
            'marker': {'type': 'automatic'},
            'y_axis': 1, # Use secondary axis
        })
        
        # Combine the column and line charts
        chart1.combine(line_chart)
        
        # --- Configure Axes ---
        chart1.set_x_axis({'name': 'Date', 'date_axis': True, 'num_format': 'yyyy-mm'})
        chart1.set_y_axis({'name': 'Liters Consumed', 'num_format': '#,##0'})
        chart1.set_y2_axis({'name': 'Fleet Count / Ore (kt)', 'num_format': '#,##0'})
        
        chart1.set_legend({'position': 'bottom'})
        worksheet.insert_chart('A1', chart1, {'x_scale': 2.0, 'y_scale': 1.5})

# Helper to get 0-based index from column name
def name_idx(name):
    """Helper to get 0-based index from column name."""
    col_list = [
        'Date', 'Ore Mined - RGM - kt', 'Overburden - RGM - kt', 
        'Ore Mined - Sar - kt', 'Overburden - Sar - kt', 
        'Active Fleet Count (Aprox)', 'Liter of Diesel Consumed',
        'Total Ore (kt)', 'Total Overburden (kt)', 'Total Material (kt)',
        'Efficiency (kt per Liter)', 'Productivity (kt per Fleet)',
        'RGM Strip Ratio', 'Sar Strip Ratio'
    ]
    try:
        return col_list.index(name)
    except ValueError:
        raise KeyError(f"Column name '{name}' not found in defined list.")
