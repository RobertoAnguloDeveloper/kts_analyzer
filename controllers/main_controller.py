"""
Filepath: kts_analyzer/controllers/main_controller.py
-------------------------------------------------
KTS Analyzer - Controller (VCSM)

**Refactored**
- Now calls data_service.get_analysis_dataframe()
- Passes the new `analysis_df` to the report_service.
- Logic remains clean and orchestral.
--------------------------------
"""

import sys
import os
import traceback
import pandas as pd

# --- Import Services ---
from services.data_service import MiningDataService
from services.report_service import XlsxReportService

class MainController:
    """
    The main application controller.
    Orchestrates the entire data analysis and report generation process.
    """

    def __init__(self):
        """
        Initialize the controller and its associated services.
        """
        self.view = None
        
        try:
            self.data_service = MiningDataService()
            self.report_service = XlsxReportService()
        except Exception as e:
            msg = f"Failed to initialize services: {e}\n{traceback.format_exc()}"
            self._update_status(msg, error=True)
            raise

    def register_view(self, view):
        """
        Allows the View to register itself with the Controller.
        
        Args:
            view: The MainView instance.
        """
        self.view = view

    def run_analysis(self, input_file: str, output_file: str, sheet_name: str = None):
        """
        The core analysis and report generation logic.
        
        Args:
            input_file: Path to the source Excel file.
            output_file: Path to save the new report.
            sheet_name: Optional name of the sheet to read.
        """
        
        try:
            # --- 1. Load Data ---
            self._update_status("Loading and preparing data...")
            df = self.data_service.load_and_prepare_data(input_file, sheet_name)
            
            if df.empty:
                self._update_status("No valid data found after cleaning. Aborting.", error=True)
                return

            self._update_status("Data loaded successfully.")

            # --- 2. Process Data ---
            self._update_status("Grouping and summarizing data...")
            
            # Get the GroupBy object for iteration
            grouped_data = self.data_service.get_data_groups(df)
            
            # Create the summary statistics DataFrame
            summary_df = df.groupby(['Category', 'SubCategory', 'Unit'])['Value'] \
                           .agg(['sum', 'mean', 'max', 'min']).reset_index()
                           
            # --- 3. Create Analysis DataFrame ---
            self._update_status("Calculating correlations and derived metrics...")
            analysis_df = self.data_service.get_analysis_dataframe(df)
            
            # --- 4. Generate Report (Data + Charts) ---
            self._update_status("Generating Excel report with native charts...")
            
            self.report_service.generate_report(
                output_file, 
                grouped_data, 
                summary_df,
                analysis_df # Pass the new analysis data
            )
            
            self._update_status(f"Report saved successfully to:\n{os.path.abspath(output_file)}", final=True)

        except (FileNotFoundError, ValueError, IOError) as e:
            self._update_status(f"Error: {e}", error=True)
        except Exception as e:
            msg = f"An unexpected error occurred: {e}\n{traceback.format_exc()}"
            self._update_status(msg, error=True)

    def run_cli(self, input_file: str, output_file: str, sheet_name: str = None):
        """
        Wrapper to run the analysis in Command-Line Interface mode.
        """
        print("Running in Command-Line Mode...")
        # ... (rest of the method is unchanged) ...

    def run_cli(self, input_file: str, output_file: str, sheet_name: str = None):
        """
        Wrapper to run the analysis in Command-Line Interface mode.
        """
        print("Running in Command-Line Mode...")
        print(f"Input:    {input_file}")
        if sheet_name:
            print(f"Sheet:    {sheet_name}")
        print(f"Output:   {output_file}")
        print("-" * 60)
        
        # Call the main analysis logic
        self.run_analysis(input_file, output_file, sheet_name)
        
        print("-" * 60)

    def _update_status(self, message: str, error: bool = False, final: bool = False):
        """
        Internal helper to send status updates.
        """
        
        if error:
            print(f"ERROR: {message}", file=sys.stderr)
        else:
            print(message)
        
        if self.view:
            self.view.root.after(0, self.view.update_status, message, error, final)

