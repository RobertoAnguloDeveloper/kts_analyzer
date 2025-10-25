"""
Filepath: kts_analyzer/services/data_service.py
--------------------------------------------
KTS Analyzer - Data Service (VCSM)

**Refactored**
- In alignment with the new "Excel-first formula" architecture,
  `get_analysis_dataframe` has been simplified.
- It NO LONGER calculates any derived metrics (Total Material,
  Efficiency, etc.).
- Its sole responsibility is now to pivot the clean, long-format
  data into a wide-format DataFrame.
- All calculation logic is being moved to the ReportService,
  which will write Excel formulas.
----------------------------------
"""

import pandas as pd
import warnings

class MiningDataService:
    """
    Handles loading and preparation of the mining data from Excel.
    """

    def __init__(self):
        """Initialize the data service."""
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        
        # Spanish to English month mapping
        self.month_map = {
            'ene': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'abr': 'Apr',
            'may': 'May', 'jun': 'Jun', 'jul': 'Jul', 'ago': 'Aug',
            'sep': 'Sep', 'oct': 'Oct', 'nov': 'Nov', 'dic': 'Dec'
        }

    def load_and_prepare_data(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Loads the Excel file, cleans it, melts it, and prepares it for analysis.
        
        Args:
            file_path: Path to the source Excel file.
            sheet_name: Optional name of the sheet to read.
            
        Returns:
            A pandas DataFrame, indexed by Date, with columns:
            ['Category', 'SubCategory', 'Unit', 'Value']
            
        Raises:
            FileNotFoundError: If the file_path is invalid.
            ValueError: If the data format is unexpected.
        """
        
        try:
            sheet_to_load = sheet_name if sheet_name else 0
            df = pd.read_excel(file_path, sheet_name=sheet_to_load, header=0)
        
        except FileNotFoundError:
            raise
        except Exception as e:
            raise ValueError(f"Failed to read Excel file '{file_path}'. Error: {e}")

        if df.empty:
            raise ValueError("The specified sheet is empty.")

        try:
            id_cols = list(df.columns[:3])
            date_cols = list(df.columns[3:])
            
            df_long = pd.melt(df, 
                              id_vars=id_cols, 
                              value_vars=date_cols, 
                              var_name='Date', 
                              value_name='Value')
        except Exception as e:
            raise ValueError(f"Failed to 'melt' data: {e}. Check data structure.")
        
        if not pd.api.types.is_numeric_dtype(df_long['Value']):
            df_long['Value'] = df_long['Value'].astype(str) \
                                              .str.replace(r'\.', '', regex=True) \
                                              .str.replace(',', '.')
        
        df_long['Value'] = pd.to_numeric(df_long['Value'], errors='coerce')
        
        if pd.api.types.is_object_dtype(df_long['Date']):
             df_long['Date'] = df_long['Date'].astype(str).str.lower() \
                                             .replace(self.month_map, regex=True)

        try:
            df_long['Date'] = pd.to_datetime(df_long['Date'])
        except Exception as e:
            raise ValueError(f"Failed to parse dates after cleaning: {e}. Check date columns.")

        df_long.dropna(subset=['Value'], inplace=True)
        
        if df_long.empty:
            return pd.DataFrame()
            
        df_long.rename(columns={
            id_cols[0]: 'Category',
            id_cols[1]: 'SubCategory',
            id_cols[2]: 'Unit'
        }, inplace=True)

        # --- FIX: Clean whitespace from ID columns ---
        for col in ['Category', 'SubCategory', 'Unit']:
            if col in df_long.columns:
                df_long[col] = df_long[col].astype(str).str.strip()
        
        # Replace empty strings or 'nan' strings with proper NA
        df_long.replace({'': pd.NA, 'nan': pd.NA, 'None': pd.NA}, inplace=True)
        
        df_long.set_index('Date', inplace=True)
        df_long.sort_index(inplace=True)

        return df_long

    def get_analysis_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Creates a wide-format DataFrame for correlation analysis.
        This service no longer calculates derived metrics.
        
        Args:
            df: The long-format DataFrame from load_and_prepare_data.
            
        Returns:
            A wide-format DataFrame with metrics as columns.
        """
        if df.empty:
            return pd.DataFrame()
            
        # Create a unique metric name, e.g., "Ore Mined - RGM - kt"
        # or "Active Fleet Count (Aprox)"
        df_copy = df.copy()
        
        df_copy['MetricName'] = df_copy.apply(
            lambda row: " - ".join(
                filter(pd.notna, [row['Category'], row['SubCategory'], row['Unit']])
            ), 
            axis=1
        )
        
        # Pivot the table to wide format: Dates as index, Metrics as columns
        analysis_df = df_copy.pivot_table(
            index='Date', 
            columns='MetricName', 
            values='Value'
        )
        
        # Re-order columns to a logical format
        # This is a bit manual but ensures a predictable layout for formulas
        ordered_cols = [
            'Ore Mined - RGM - kt',
            'Overburden - RGM - kt',
            'Ore Mined - Sar - kt',
            'Overburden - Sar - kt',
            'Active Fleet Count (Aprox)',
            'Liter of Diesel Consumed'
        ]
        
        # Filter for columns that actually exist in the data
        final_cols = [col for col in ordered_cols if col in analysis_df.columns]
        # Add any extra columns that weren't in the preferred list
        extra_cols = [col for col in analysis_df.columns if col not in final_cols]
        
        analysis_df = analysis_df[final_cols + extra_cols]

        return analysis_df

    # -----------------------------------------------------------------
    # --- DEPRECATED METHODS ---
    # These functions are no longer needed, as the new architecture
    # does not create individual metric sheets or Python-based summaries.
    # -----------------------------------------------------------------

    def get_data_groups(self, df: pd.DataFrame):
        """
        DEPRECATED: This method is no longer used by the controller.
        """
        print("Warning: get_data_groups() is deprecated.")
        return pd.DataFrame().groupby(lambda: True) 

    def summarize_data(self, df_groups: pd.api.typing.DataFrameGroupBy) -> pd.DataFrame:
        """
        DEPRECATED: This method is no longer used by the controller.
        """
        print("Warning: summarize_data() is deprecated.")
        return pd.DataFrame()

