"""
Filepath: kts_analyzer/controllers/main_controller.py
-------------------------------------------------
KTS Analyzer - Controller (VCSM)

**Refactored**
- Simplified `run_analysis` logic to align with the new
  "Excel-first formula" architecture.
- Removed all logic related to `get_data_groups` and
  `summarize_data`, as these are no longer used.
- The controller now only orchestrates the loading of
  the base pivot table and passing it to the report service.
- Removed `plotting_service` as it's no longer used.
--------------------------------
"""

import sys
import os
import traceback
import threading

# --- Import Services ---
from services.data_service import MiningDataService
# PlottingService is no longer needed
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
        
        # Instantiate all necessary services
        try:
            self.data_service = MiningDataService()
            # self.plotting_service = PlottingService() # No longer needed
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
        Runs the full analysis and report generation process.
        This function is designed to be run in a separate thread from the GUI.
        
        Args:
            input_file: Path to the source Excel file.
            output_file: Path to save the new report.
            sheet_name: Optional name of the sheet to read.
        """
        
        try:
            # 1. Load Data
            self._update_status("Loading and preparing data...")
            df_long = self.data_service.load_and_prepare_data(input_file, sheet_name)
            if df_long.empty:
                raise ValueError("No valid data found in the file.")
            self._update_status("Data loaded successfully.")

            # 2. Get Analysis Pivot Table
            self._update_status("Creating base pivot table...")
            analysis_df = self.data_service.get_analysis_dataframe(df_long)
            if analysis_df.empty:
                raise ValueError("Failed to create pivot table.")
            self._update_status("Base table created.")

            # 3. Generate Report
            # The report service now handles ALL calculations and charts.
            self._update_status("Generating Excel report with native formulas and charts...")
            self.report_service.generate_report(
                output_file,
                analysis_df
            )
            
            final_path = os.path.abspath(output_file)
            self._update_status(f"Report saved successfully to:\n{final_path}", final=True)

        except Exception as e:
            # Print the full traceback for debugging
            print(f"ERROR: {traceback.format_exc()}", file=sys.stderr)
            self._update_status(f"An error occurred: {e}", error=True, final=True)

    def run_analysis_threaded(self, input_file: str, output_file: str, sheet_name: str = None):
        """
        Launches the `run_analysis` function in a new thread
        to keep the GUI responsive.
        
        Args:
            input_file: Path to the source Excel file.
            output_file: Path to save the new report.
            sheet_name: Optional name of the sheet to read.
        """
        analysis_thread = threading.Thread(
            target=self.run_analysis,
            args=(input_file, output_file, sheet_name),
            daemon=True # Ensures thread exits when main app exits
        )
        analysis_thread.start()

    def run_cli(self, input_file: str, output_file: str, sheet_name: str = None):
        """
        Wrapper to run the analysis in Command-Line Interface mode.
        
        Args:
            input_file: Path to the source Excel file.
            output_file: Path to save the new report.
            sheet_name: Optional name of the sheet to read.
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
        It prints to console (for CLI) and updates the View (for GUI).
        
        Args:
            message: The status message.
            error: True if this is an error.
            final: True if this is the final status.
        """
        
        # Always print to console for logging
        if error:
            print(f"ERROR: {message}", file=sys.stderr)
        else:
            print(message)
            
        # If a GUI is attached, update it
        if self.view:
            # Ensure GUI updates run on the main thread
            self.view.root.after(0, self.view.update_status, message, error, final)
