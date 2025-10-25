"""
Filepath: kts_analyzer/services/data_service.py
--------------------------------------------
KTS Analyzer - Data Service (VCSM)

**Refactored**
- **CRITICAL FIX:** Added .str.strip() to 'Category', 'SubCategory',
  and 'Unit' columns after loading. This removes leading/trailing
  whitespace that was causing KeyError (e.g., 'Liter of Diesel Consumed ').
- Replaced empty strings ('') with pd.NA to ensure they are
  filtered out when creating metric names.
- Retained all previous logic for analysis.
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
        # This prevents KeyErrors when looking up 'Liter of Diesel Consumed'
        for col in ['Category', 'SubCategory', 'Unit']:
            if col in df_long.columns:
                df_long[col] = df_long[col].astype(str).str.strip()
        
        # Replace empty strings or 'nan' strings with proper NA
        df_long.replace({'': pd.NA, 'nan': pd.NA, 'None': pd.NA}, inplace=True)
        
        df_long.set_index('Date', inplace=True)
        df_long.sort_index(inplace=True)

        return df_long

    def get_data_groups(self, df: pd.DataFrame) -> pd.api.typing.DataFrameGroupBy:
        """
        Groups the prepared DataFrame by the identifier columns.
        
        Args:
            df: The prepared DataFrame from load_and_prepare_data.
            
        Returns:
            A DataFrameGroupBy object.
        """
        if df.empty:
            return pd.DataFrame().groupby(lambda: True) 
            
        return df.groupby(['Category', 'SubCategory', 'Unit'])

    def get_analysis_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Creates a wide-format DataFrame for correlation analysis
        and calculates key derived metrics.
        
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
        
        # Now that data is cleaned, this filter is robust
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
        
        # --- Calculate Derived Metrics for Analysis ---
        
        try:
            # Sum all 'Ore Mined' columns
            ore_cols = [c for c in analysis_df.columns if 'Ore Mined' in c]
            if ore_cols:
                analysis_df['Total Ore Mined (kt)'] = analysis_df[ore_cols].sum(axis=1)
            
            # Sum all 'Overburden' columns
            overburden_cols = [c for c in analysis_df.columns if 'Overburden' in c]
            if overburden_cols:
                analysis_df['Total Overburden (kt)'] = analysis_df[overburden_cols].sum(axis=1)
            
            # Create 'Total Material (kt)'
            total_material_cols = [c for c in ['Total Ore Mined (kt)', 'Total Overburden (kt)'] if c in analysis_df.columns]
            if total_material_cols:
                 analysis_df['Total Material (kt)'] = analysis_df[total_material_cols].sum(axis=1)
            
            # Define key metric columns
            fleet_col = 'Active Fleet Count (Aprox)'
            fuel_col = 'Liter of Diesel Consumed' # This key will now be found
            material_col = 'Total Material (kt)'

            # Calculate Efficiency (kt per Liter)
            if material_col in analysis_df.columns and fuel_col in analysis_df.columns:
                # Use .replace(0, pd.NA) to avoid ZeroDivisionError
                analysis_df['Efficiency (kt per Liter)'] = \
                    analysis_df[material_col] / analysis_df[fuel_col].replace(0, pd.NA)
            
            # Calculate Productivity (kt per Fleet)
            if material_col in analysis_df.columns and fleet_col in analysis_df.columns:
                analysis_df['Productivity (kt per Fleet)'] = \
                    analysis_df[material_col] / analysis_df[fleet_col].replace(0, pd.NA)

        except KeyError as e:
            print(f"Warning: Could not calculate derived metric. Missing column: {e}")
        except Exception as e:
            print(f"Warning: Error calculating derived metrics: {e}")

        return analysis_df

