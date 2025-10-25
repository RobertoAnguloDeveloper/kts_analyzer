import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class MiningDataAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Mining Data Analyzer")
        self.root.geometry("1200x800")
        
        # Variables
        self.file_path = None
        self.excel_data = None
        self.df = None
        self.sheet_names = []
        
        # Create main frame
        self.setup_gui()
        
    def setup_gui(self):
        # Top frame for file selection
        top_frame = tk.Frame(self.root, bg='#2c3e50', height=100)
        top_frame.pack(fill='x', padx=5, pady=5)
        
        # Title
        title_label = tk.Label(top_frame, text="Mining Operations Data Analyzer", 
                              font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white')
        title_label.pack(pady=10)
        
        # File selection frame
        file_frame = tk.Frame(top_frame, bg='#2c3e50')
        file_frame.pack(pady=5)
        
        # Upload button
        self.upload_btn = tk.Button(file_frame, text="📁 Upload Excel File", 
                                   command=self.upload_file, font=('Arial', 10),
                                   bg='#3498db', fg='white', padx=20, pady=5)
        self.upload_btn.pack(side='left', padx=5)
        
        # Sheet selection
        tk.Label(file_frame, text="Select Sheet:", bg='#2c3e50', 
                fg='white', font=('Arial', 10)).pack(side='left', padx=5)
        
        self.sheet_combo = ttk.Combobox(file_frame, state='disabled', width=30)
        self.sheet_combo.pack(side='left', padx=5)
        self.sheet_combo.bind('<<ComboboxSelected>>', self.load_sheet)
        
        # Process button
        self.process_btn = tk.Button(file_frame, text="📊 Generate Charts", 
                                    command=self.generate_charts, font=('Arial', 10),
                                    bg='#27ae60', fg='white', padx=20, pady=5,
                                    state='disabled')
        self.process_btn.pack(side='left', padx=5)
        
        # Status label
        self.status_label = tk.Label(top_frame, text="Please upload an Excel file to begin", 
                                    bg='#2c3e50', fg='#ecf0f1', font=('Arial', 9))
        self.status_label.pack(pady=5)
        
        # Main content frame with scrollbar
        self.canvas_frame = tk.Frame(self.root)
        self.canvas_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create scrollable frame
        self.canvas = tk.Canvas(self.canvas_frame, bg='white')
        self.scrollbar = ttk.Scrollbar(self.canvas_frame, orient='vertical', 
                                      command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='white')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
    def upload_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if self.file_path:
            try:
                # Read Excel file to get sheet names
                self.excel_data = pd.ExcelFile(self.file_path)
                self.sheet_names = self.excel_data.sheet_names
                
                # Update sheet combobox
                self.sheet_combo['values'] = self.sheet_names
                self.sheet_combo['state'] = 'readonly'
                if self.sheet_names:
                    self.sheet_combo.current(0)
                    
                self.status_label.config(text=f"File loaded: {self.file_path.split('/')[-1]}")
                messagebox.showinfo("Success", "Excel file loaded successfully!\nPlease select a sheet.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")
                
    def load_sheet(self, event=None):
        if self.sheet_combo.get():
            try:
                # Read the selected sheet
                self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_combo.get())
                self.process_data()
                self.process_btn['state'] = 'normal'
                self.status_label.config(text=f"Sheet '{self.sheet_combo.get()}' loaded successfully!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sheet: {str(e)}")
                
    def process_data(self):
        """Process and clean the data"""
        try:
            # Find the header row (containing months)
            header_row = None
            for idx, row in self.df.iterrows():
                if any(isinstance(val, str) and 'ene-' in str(val).lower() for val in row.values):
                    header_row = idx
                    break
            
            if header_row is None:
                # Try to identify if first row contains date-like values
                first_row = self.df.iloc[0]
                if any('-20' in str(val) or '-21' in str(val) or '-22' in str(val) 
                      or '-23' in str(val) or '-24' in str(val) for val in first_row.values):
                    header_row = 0
            
            # Set the header row
            if header_row is not None:
                self.df.columns = self.df.iloc[header_row]
                self.df = self.df.iloc[header_row + 1:].reset_index(drop=True)
            
            # Clean column names
            self.df.columns = [str(col).strip() if pd.notna(col) else f'Col_{i}' 
                              for i, col in enumerate(self.df.columns)]
            
            # Identify metric columns (first few columns)
            metric_cols = []
            for col in self.df.columns[:4]:
                if pd.notna(col) and col not in ['nan', '']:
                    metric_cols.append(col)
            
            # Create a cleaner dataframe
            data_dict = {}
            
            # Extract metrics
            for idx, row in self.df.iterrows():
                if pd.notna(row.iloc[0]):
                    metric_name = str(row.iloc[0]).strip()
                    if row.iloc[1:3].notna().any():  # Has subcategory
                        subcategory = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
                        unit = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ''
                        key = f"{metric_name}_{subcategory}".replace(' ', '_')
                    else:
                        key = metric_name.replace(' ', '_')
                    
                    # Get the data values (from column 3 onwards)
                    values = []
                    for col in self.df.columns[3:]:
                        if '-' in str(col):  # This is a date column
                            val = row[col]
                            if pd.notna(val):
                                # Clean the value
                                if isinstance(val, str):
                                    val = val.replace('.', '').replace(',', '.')
                                try:
                                    values.append(float(val))
                                except:
                                    values.append(np.nan)
                            else:
                                values.append(np.nan)
                    
                    if values and not all(pd.isna(values)):
                        data_dict[key] = values
            
            # Get date columns
            date_cols = [col for col in self.df.columns[3:] if '-' in str(col)]
            
            # Create clean dataframe
            self.clean_df = pd.DataFrame(data_dict)
            self.clean_df['Date'] = date_cols[:len(self.clean_df)]
            
            # Convert date strings to datetime
            def parse_date(date_str):
                months = {'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 
                         'may': '05', 'jun': '06', 'jul': '07', 'ago': '08',
                         'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12'}
                parts = str(date_str).split('-')
                if len(parts) == 2:
                    month = months.get(parts[0].lower(), '01')
                    year = '20' + parts[1] if len(parts[1]) == 2 else parts[1]
                    return pd.to_datetime(f"{year}-{month}-01")
                return pd.NaT
            
            self.clean_df['Date'] = self.clean_df['Date'].apply(parse_date)
            self.clean_df = self.clean_df.set_index('Date').sort_index()
            
            # Fill NaN values with 0 for calculations
            self.clean_df = self.clean_df.fillna(0)
            
        except Exception as e:
            print(f"Error processing data: {str(e)}")
            messagebox.showwarning("Warning", f"Data processing had issues: {str(e)}")
            
    def generate_charts(self):
        """Generate all charts"""
        if self.clean_df is None or self.clean_df.empty:
            messagebox.showerror("Error", "No data to visualize!")
            return
        
        # Clear previous charts
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        try:
            # Create charts
            self.create_production_charts()
            self.create_efficiency_charts()
            self.create_comparative_charts()
            self.create_trend_analysis()
            
            self.status_label.config(text="Charts generated successfully!")
            messagebox.showinfo("Success", "All charts have been generated!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate charts: {str(e)}")
            print(f"Error details: {str(e)}")
            
    def create_production_charts(self):
        """Create production-related charts"""
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('Production Analysis', fontsize=16, fontweight='bold')
        
        # Chart 1: Ore Mined Comparison
        ax1 = axes[0, 0]
        if 'Ore_Mined_RGM' in self.clean_df.columns:
            ax1.plot(self.clean_df.index, self.clean_df['Ore_Mined_RGM'], 
                    marker='o', label='RGM', linewidth=2, markersize=4)
        if 'Ore_Mined_Sar' in self.clean_df.columns:
            ax1.plot(self.clean_df.index, self.clean_df['Ore_Mined_Sar'], 
                    marker='s', label='Sar', linewidth=2, markersize=4)
        ax1.set_title('Ore Mined Over Time', fontweight='bold')
        ax1.set_xlabel('Date')
        ax1.set_ylabel('Ore Mined (kt)')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        ax1.tick_params(axis='x', rotation=45)
        
        # Chart 2: Overburden Comparison
        ax2 = axes[0, 1]
        if 'Overburden_RGM' in self.clean_df.columns:
            ax2.plot(self.clean_df.index, self.clean_df['Overburden_RGM'], 
                    marker='o', label='RGM', linewidth=2, markersize=4)
        if 'Overburden_Sar' in self.clean_df.columns:
            ax2.plot(self.clean_df.index, self.clean_df['Overburden_Sar'], 
                    marker='s', label='Sar', linewidth=2, markersize=4)
        ax2.set_title('Overburden Moved Over Time', fontweight='bold')
        ax2.set_xlabel('Date')
        ax2.set_ylabel('Overburden (kt)')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        ax2.tick_params(axis='x', rotation=45)
        
        # Chart 3: Total Material Movement
        ax3 = axes[1, 0]
        total_ore = pd.Series(0, index=self.clean_df.index)
        total_overburden = pd.Series(0, index=self.clean_df.index)
        
        if 'Ore_Mined_RGM' in self.clean_df.columns:
            total_ore += self.clean_df['Ore_Mined_RGM']
        if 'Ore_Mined_Sar' in self.clean_df.columns:
            total_ore += self.clean_df['Ore_Mined_Sar']
        if 'Overburden_RGM' in self.clean_df.columns:
            total_overburden += self.clean_df['Overburden_RGM']
        if 'Overburden_Sar' in self.clean_df.columns:
            total_overburden += self.clean_df['Overburden_Sar']
            
        ax3.bar(self.clean_df.index, total_ore, label='Total Ore', alpha=0.7)
        ax3.bar(self.clean_df.index, total_overburden, bottom=total_ore, 
               label='Total Overburden', alpha=0.7)
        ax3.set_title('Total Material Movement', fontweight='bold')
        ax3.set_xlabel('Date')
        ax3.set_ylabel('Material (kt)')
        ax3.legend()
        ax3.grid(True, alpha=0.3, axis='y')
        ax3.tick_params(axis='x', rotation=45)
        
        # Chart 4: Stripping Ratio
        ax4 = axes[1, 1]
        if 'Overburden_RGM' in self.clean_df.columns and 'Ore_Mined_RGM' in self.clean_df.columns:
            strip_ratio_rgm = self.clean_df['Overburden_RGM'] / (self.clean_df['Ore_Mined_RGM'] + 0.001)
            ax4.plot(self.clean_df.index, strip_ratio_rgm, marker='o', 
                    label='RGM Strip Ratio', linewidth=2, markersize=4)
        if 'Overburden_Sar' in self.clean_df.columns and 'Ore_Mined_Sar' in self.clean_df.columns:
            strip_ratio_sar = self.clean_df['Overburden_Sar'] / (self.clean_df['Ore_Mined_Sar'] + 0.001)
            ax4.plot(self.clean_df.index, strip_ratio_sar, marker='s', 
                    label='Sar Strip Ratio', linewidth=2, markersize=4)
        ax4.set_title('Stripping Ratio Trends', fontweight='bold')
        ax4.set_xlabel('Date')
        ax4.set_ylabel('Strip Ratio')
        ax4.legend()
        ax4.grid(True, alpha=0.3)
        ax4.tick_params(axis='x', rotation=45)
        
        plt.tight_layout()
        self.embed_chart(fig, "Production Analysis")
        
    def create_efficiency_charts(self):
        """Create efficiency and fleet utilization charts"""
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('Operational Efficiency Analysis', fontsize=16, fontweight='bold')
        
        # Chart 1: Fleet Count vs Production
        ax1 = axes[0, 0]
        if 'Active_Fleet_Count_(Aprox)' in self.clean_df.columns:
            ax1_twin = ax1.twinx()
            
            total_production = pd.Series(0, index=self.clean_df.index)
            if 'Ore_Mined_RGM' in self.clean_df.columns:
                total_production += self.clean_df['Ore_Mined_RGM']
            if 'Ore_Mined_Sar' in self.clean_df.columns:
                total_production += self.clean_df['Ore_Mined_Sar']
            
            ax1.bar(self.clean_df.index, total_production, alpha=0.5, 
                   color='skyblue', label='Total Ore Production')
            ax1_twin.plot(self.clean_df.index, self.clean_df['Active_Fleet_Count_(Aprox)'], 
                         color='red', marker='o', linewidth=2, markersize=4, 
                         label='Fleet Count')
            
            ax1.set_xlabel('Date')
            ax1.set_ylabel('Ore Production (kt)', color='blue')
            ax1_twin.set_ylabel('Fleet Count', color='red')
            ax1.set_title('Fleet Utilization vs Production', fontweight='bold')
            ax1.tick_params(axis='x', rotation=45)
            ax1.grid(True, alpha=0.3)
            
            # Add legends
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax1_twin.get_legend_handles_labels()
            ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
        
        # Chart 2: Diesel Consumption Trends
        ax2 = axes[0, 1]
        if 'Liter_of_Diesel_Consumed' in self.clean_df.columns:
            ax2.plot(self.clean_df.index, self.clean_df['Liter_of_Diesel_Consumed']/1000000, 
                    marker='o', color='green', linewidth=2, markersize=4)
            ax2.set_title('Diesel Consumption Over Time', fontweight='bold')
            ax2.set_xlabel('Date')
            ax2.set_ylabel('Diesel Consumed (Million Liters)')
            ax2.grid(True, alpha=0.3)
            ax2.tick_params(axis='x', rotation=45)
        
        # Chart 3: Productivity per Fleet Unit
        ax3 = axes[1, 0]
        if 'Active_Fleet_Count_(Aprox)' in self.clean_df.columns:
            total_material = pd.Series(0, index=self.clean_df.index)
            for col in ['Ore_Mined_RGM', 'Ore_Mined_Sar', 'Overburden_RGM', 'Overburden_Sar']:
                if col in self.clean_df.columns:
                    total_material += self.clean_df[col]
            
            productivity = total_material / (self.clean_df['Active_Fleet_Count_(Aprox)'] + 0.001)
            ax3.plot(self.clean_df.index, productivity, marker='o', 
                    color='purple', linewidth=2, markersize=4)
            ax3.set_title('Productivity per Fleet Unit', fontweight='bold')
            ax3.set_xlabel('Date')
            ax3.set_ylabel('Material per Fleet Unit (kt)')
            ax3.grid(True, alpha=0.3)
            ax3.tick_params(axis='x', rotation=45)
        
        # Chart 4: Fuel Efficiency
        ax4 = axes[1, 1]
        if 'Liter_of_Diesel_Consumed' in self.clean_df.columns:
            total_material = pd.Series(0, index=self.clean_df.index)
            for col in ['Ore_Mined_RGM', 'Ore_Mined_Sar', 'Overburden_RGM', 'Overburden_Sar']:
                if col in self.clean_df.columns:
                    total_material += self.clean_df[col]
            
            fuel_efficiency = self.clean_df['Liter_of_Diesel_Consumed'] / (total_material + 0.001)
            ax4.plot(self.clean_df.index, fuel_efficiency, marker='o', 
                    color='orange', linewidth=2, markersize=4)
            ax4.set_title('Fuel Efficiency (Liters per kt)', fontweight='bold')
            ax4.set_xlabel('Date')
            ax4.set_ylabel('Liters per kt')
            ax4.grid(True, alpha=0.3)
            ax4.tick_params(axis='x', rotation=45)
        
        plt.tight_layout()
        self.embed_chart(fig, "Efficiency Analysis")
        
    def create_comparative_charts(self):
        """Create comparative analysis charts"""
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('Comparative Analysis: RGM vs Sar', fontsize=16, fontweight='bold')
        
        # Chart 1: Production Share
        ax1 = axes[0, 0]
        if 'Ore_Mined_RGM' in self.clean_df.columns and 'Ore_Mined_Sar' in self.clean_df.columns:
            rgm_total = self.clean_df['Ore_Mined_RGM'].sum()
            sar_total = self.clean_df['Ore_Mined_Sar'].sum()
            ax1.pie([rgm_total, sar_total], labels=['RGM', 'Sar'], 
                   autopct='%1.1f%%', startangle=90, colors=['#3498db', '#e74c3c'])
            ax1.set_title('Total Ore Production Share', fontweight='bold')
        
        # Chart 2: Monthly Production Comparison
        ax2 = axes[0, 1]
        if 'Ore_Mined_RGM' in self.clean_df.columns and 'Ore_Mined_Sar' in self.clean_df.columns:
            width = 10
            x = np.arange(len(self.clean_df.index))
            ax2.bar(x - width/2, self.clean_df['Ore_Mined_RGM'], width, 
                   label='RGM', alpha=0.7, color='#3498db')
            ax2.bar(x + width/2, self.clean_df['Ore_Mined_Sar'], width, 
                   label='Sar', alpha=0.7, color='#e74c3c')
            ax2.set_title('Monthly Ore Production Comparison', fontweight='bold')
            ax2.set_xlabel('Date')
            ax2.set_ylabel('Ore Mined (kt)')
            ax2.legend()
            ax2.grid(True, alpha=0.3, axis='y')
            
            # Set x-axis labels (show every 6th label)
            labels = [date.strftime('%b-%y') if i % 6 == 0 else '' 
                     for i, date in enumerate(self.clean_df.index)]
            ax2.set_xticks(x)
            ax2.set_xticklabels(labels, rotation=45, ha='right')
        
        # Chart 3: Overburden Share
        ax3 = axes[1, 0]
        if 'Overburden_RGM' in self.clean_df.columns and 'Overburden_Sar' in self.clean_df.columns:
            rgm_ob = self.clean_df['Overburden_RGM'].sum()
            sar_ob = self.clean_df['Overburden_Sar'].sum()
            ax3.pie([rgm_ob, sar_ob], labels=['RGM', 'Sar'], 
                   autopct='%1.1f%%', startangle=90, colors=['#9b59b6', '#f39c12'])
            ax3.set_title('Total Overburden Share', fontweight='bold')
        
        # Chart 4: Efficiency Comparison
        ax4 = axes[1, 1]
        metrics = []
        rgm_values = []
        sar_values = []
        
        if all(col in self.clean_df.columns for col in ['Ore_Mined_RGM', 'Overburden_RGM', 
                                                        'Ore_Mined_Sar', 'Overburden_Sar']):
            # Average Strip Ratio
            rgm_strip = (self.clean_df['Overburden_RGM'] / 
                        (self.clean_df['Ore_Mined_RGM'] + 0.001)).mean()
            sar_strip = (self.clean_df['Overburden_Sar'] / 
                        (self.clean_df['Ore_Mined_Sar'] + 0.001)).mean()
            
            metrics.append('Avg Strip Ratio')
            rgm_values.append(rgm_strip)
            sar_values.append(sar_strip)
            
            # Average Production
            metrics.append('Avg Ore (kt)')
            rgm_values.append(self.clean_df['Ore_Mined_RGM'].mean())
            sar_values.append(self.clean_df['Ore_Mined_Sar'].mean())
            
            # Total Production
            metrics.append('Total Ore (kt/1000)')
            rgm_values.append(self.clean_df['Ore_Mined_RGM'].sum()/1000)
            sar_values.append(self.clean_df['Ore_Mined_Sar'].sum()/1000)
            
            x = np.arange(len(metrics))
            width = 0.35
            ax4.bar(x - width/2, rgm_values, width, label='RGM', color='#3498db')
            ax4.bar(x + width/2, sar_values, width, label='Sar', color='#e74c3c')
            ax4.set_xlabel('Metrics')
            ax4.set_ylabel('Values')
            ax4.set_title('Performance Metrics Comparison', fontweight='bold')
            ax4.set_xticks(x)
            ax4.set_xticklabels(metrics)
            ax4.legend()
            ax4.grid(True, alpha=0.3, axis='y')
        
        plt.tight_layout()
        self.embed_chart(fig, "Comparative Analysis")
        
    def create_trend_analysis(self):
        """Create trend analysis and forecasting charts"""
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('Trend Analysis & Insights', fontsize=16, fontweight='bold')
        
        # Chart 1: Moving Averages
        ax1 = axes[0, 0]
        if 'Ore_Mined_RGM' in self.clean_df.columns:
            ma_3 = self.clean_df['Ore_Mined_RGM'].rolling(window=3, center=True).mean()
            ma_6 = self.clean_df['Ore_Mined_RGM'].rolling(window=6, center=True).mean()
            
            ax1.plot(self.clean_df.index, self.clean_df['Ore_Mined_RGM'], 
                    alpha=0.3, label='Actual', linewidth=1)
            ax1.plot(self.clean_df.index, ma_3, label='3-Month MA', 
                    linewidth=2, color='red')
            ax1.plot(self.clean_df.index, ma_6, label='6-Month MA', 
                    linewidth=2, color='green')
            ax1.set_title('RGM Ore Production - Moving Averages', fontweight='bold')
            ax1.set_xlabel('Date')
            ax1.set_ylabel('Ore Mined (kt)')
            ax1.legend()
            ax1.grid(True, alpha=0.3)
            ax1.tick_params(axis='x', rotation=45)
        
        # Chart 2: Year-over-Year Comparison
        ax2 = axes[0, 1]
        yearly_data = {}
        for col in ['Ore_Mined_RGM', 'Ore_Mined_Sar']:
            if col in self.clean_df.columns:
                yearly = self.clean_df[col].groupby(self.clean_df.index.year).sum()
                for year, value in yearly.items():
                    if year not in yearly_data:
                        yearly_data[year] = {}
                    yearly_data[year][col] = value
        
        if yearly_data:
            years = sorted(yearly_data.keys())
            rgm_yearly = [yearly_data[y].get('Ore_Mined_RGM', 0) for y in years]
            sar_yearly = [yearly_data[y].get('Ore_Mined_Sar', 0) for y in years]
            
            x = np.arange(len(years))
            width = 0.35
            ax2.bar(x - width/2, rgm_yearly, width, label='RGM', color='#3498db')
            ax2.bar(x + width/2, sar_yearly, width, label='Sar', color='#e74c3c')
            ax2.set_xlabel('Year')
            ax2.set_ylabel('Total Ore Mined (kt)')
            ax2.set_title('Yearly Production Comparison', fontweight='bold')
            ax2.set_xticks(x)
            ax2.set_xticklabels(years)
            ax2.legend()
            ax2.grid(True, alpha=0.3, axis='y')
        
        # Chart 3: Correlation Analysis
        ax3 = axes[1, 0]
        if ('Active_Fleet_Count_(Aprox)' in self.clean_df.columns and 
            'Liter_of_Diesel_Consumed' in self.clean_df.columns):
            ax3.scatter(self.clean_df['Active_Fleet_Count_(Aprox)'], 
                       self.clean_df['Liter_of_Diesel_Consumed']/1000000,
                       alpha=0.6, s=50, c=range(len(self.clean_df)), cmap='viridis')
            ax3.set_xlabel('Fleet Count')
            ax3.set_ylabel('Diesel Consumed (Million L)')
            ax3.set_title('Fleet Count vs Diesel Consumption', fontweight='bold')
            ax3.grid(True, alpha=0.3)
            
            # Add trend line
            z = np.polyfit(self.clean_df['Active_Fleet_Count_(Aprox)'], 
                          self.clean_df['Liter_of_Diesel_Consumed']/1000000, 1)
            p = np.poly1d(z)
            ax3.plot(sorted(self.clean_df['Active_Fleet_Count_(Aprox)']), 
                    p(sorted(self.clean_df['Active_Fleet_Count_(Aprox)'])), 
                    "r--", alpha=0.8, label=f'Trend: y={z[0]:.2f}x+{z[1]:.2f}')
            ax3.legend()
        
        # Chart 4: Monthly Seasonality
        ax4 = axes[1, 1]
        if 'Ore_Mined_RGM' in self.clean_df.columns:
            monthly_avg = self.clean_df['Ore_Mined_RGM'].groupby(
                self.clean_df.index.month).mean()
            
            months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            
            ax4.bar(range(1, 13), monthly_avg.values, color='steelblue', alpha=0.7)
            ax4.set_xlabel('Month')
            ax4.set_ylabel('Average Ore Mined (kt)')
            ax4.set_title('RGM Production - Monthly Seasonality', fontweight='bold')
            ax4.set_xticks(range(1, 13))
            ax4.set_xticklabels(months, rotation=45)
            ax4.grid(True, alpha=0.3, axis='y')
            
            # Add average line
            avg_line = monthly_avg.mean()
            ax4.axhline(y=avg_line, color='red', linestyle='--', 
                       label=f'Overall Avg: {avg_line:.1f}')
            ax4.legend()
        
        plt.tight_layout()
        self.embed_chart(fig, "Trend Analysis")
        
    def embed_chart(self, fig, title):
        """Embed matplotlib chart in tkinter frame"""
        frame = tk.Frame(self.scrollable_frame, bg='white', relief='raised', bd=2)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Add title
        title_label = tk.Label(frame, text=title, font=('Arial', 12, 'bold'), 
                              bg='white')
        title_label.pack(pady=5)
        
        # Embed the chart
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        
        # Add save button
        save_btn = tk.Button(frame, text="💾 Save Chart", 
                           command=lambda: self.save_chart(fig, title),
                           bg='#2ecc71', fg='white', font=('Arial', 9))
        save_btn.pack(pady=5)
        
    def save_chart(self, fig, title):
        """Save chart to file"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), 
                      ("All files", "*.*")],
            initialfile=f"{title.replace(' ', '_')}.png"
        )
        if file_path:
            fig.savefig(file_path, dpi=300, bbox_inches='tight')
            messagebox.showinfo("Success", f"Chart saved to {file_path}")

def main():
    root = tk.Tk()
    app = MiningDataAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()
