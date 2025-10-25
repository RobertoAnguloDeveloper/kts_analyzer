import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
import os
import threading
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

plt.style.use('seaborn-v0_8-darkgrid')

class MiningAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Mining Data Analyzer - Excel Chart Generator")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # Variables
        self.input_file = None
        self.output_file = None
        self.generator = None
        
        # Setup GUI
        self.setup_gui()
        
    def setup_gui(self):
        # Main frame
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Title
        title_label = tk.Label(main_frame, 
                               text="Mining Data Excel Analyzer", 
                               font=('Arial', 18, 'bold'),
                               bg='#f0f0f0')
        title_label.pack(pady=(0, 20))
        
        # Description
        desc_label = tk.Label(main_frame,
                             text="Generate comprehensive charts from your mining data\nand embed them directly in Excel",
                             font=('Arial', 10),
                             bg='#f0f0f0',
                             justify='center')
        desc_label.pack(pady=(0, 20))
        
        # File selection frame
        file_frame = tk.LabelFrame(main_frame, 
                                  text="Step 1: Select Excel File",
                                  font=('Arial', 10, 'bold'),
                                  bg='#f0f0f0',
                                  padx=10, pady=10)
        file_frame.pack(fill='x', pady=(0, 15))
        
        # File path display
        self.file_label = tk.Label(file_frame,
                                   text="No file selected",
                                   font=('Arial', 9),
                                   bg='#f0f0f0',
                                   fg='gray',
                                   anchor='w')
        self.file_label.pack(fill='x', pady=(0, 10))
        
        # Browse button
        self.browse_btn = tk.Button(file_frame,
                                    text="📁 Browse Excel File",
                                    font=('Arial', 10),
                                    command=self.browse_file,
                                    bg='#3498db',
                                    fg='white',
                                    padx=20,
                                    pady=8,
                                    cursor='hand2')
        self.browse_btn.pack()
        
        # Options frame
        options_frame = tk.LabelFrame(main_frame,
                                     text="Step 2: Options (Optional)",
                                     font=('Arial', 10, 'bold'),
                                     bg='#f0f0f0',
                                     padx=10, pady=10)
        options_frame.pack(fill='x', pady=(0, 15))
        
        # Sheet selection
        sheet_label = tk.Label(options_frame,
                               text="Sheet name (leave empty for first sheet):",
                               font=('Arial', 9),
                               bg='#f0f0f0')
        sheet_label.pack(anchor='w')
        
        self.sheet_entry = tk.Entry(options_frame, font=('Arial', 9))
        self.sheet_entry.pack(fill='x', pady=(5, 10))
        
        # Output filename
        output_label = tk.Label(options_frame,
                               text="Output filename (leave empty for auto-naming):",
                               font=('Arial', 9),
                               bg='#f0f0f0')
        output_label.pack(anchor='w')
        
        self.output_entry = tk.Entry(options_frame, font=('Arial', 9))
        self.output_entry.pack(fill='x', pady=(5, 0))
        
        # Process frame
        process_frame = tk.LabelFrame(main_frame,
                                     text="Step 3: Generate Charts",
                                     font=('Arial', 10, 'bold'),
                                     bg='#f0f0f0',
                                     padx=10, pady=10)
        process_frame.pack(fill='x', pady=(0, 15))
        
        # Process button
        self.process_btn = tk.Button(process_frame,
                                     text="📊 Generate Charts in Excel",
                                     font=('Arial', 11, 'bold'),
                                     command=self.process_data,
                                     bg='#27ae60',
                                     fg='white',
                                     padx=30,
                                     pady=12,
                                     cursor='hand2',
                                     state='disabled')
        self.process_btn.pack(pady=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(process_frame, 
                                       mode='indeterminate',
                                       length=400)
        self.progress.pack(pady=(0, 10))
        
        # Status label
        self.status_label = tk.Label(process_frame,
                                     text="Ready to process",
                                     font=('Arial', 9),
                                     bg='#f0f0f0',
                                     fg='gray')
        self.status_label.pack()
        
        # Footer
        footer_label = tk.Label(main_frame,
                               text="Charts will be embedded in new Excel sheets",
                               font=('Arial', 8, 'italic'),
                               bg='#f0f0f0',
                               fg='gray')
        footer_label.pack(side='bottom', pady=(10, 0))
        
    def browse_file(self):
        """Browse for Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Mining Data Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.input_file = filename
            # Display only filename, not full path
            display_name = os.path.basename(filename)
            if len(display_name) > 50:
                display_name = display_name[:47] + "..."
            self.file_label.config(text=display_name, fg='black')
            self.process_btn.config(state='normal')
            self.status_label.config(text="File selected. Ready to generate charts.")
            
    def process_data(self):
        """Process data in a separate thread"""
        if not self.input_file:
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        
        # Disable buttons during processing
        self.browse_btn.config(state='disabled')
        self.process_btn.config(state='disabled')
        self.progress.start(10)
        self.status_label.config(text="Processing... Please wait")
        
        # Run processing in separate thread
        thread = threading.Thread(target=self._process_data_thread)
        thread.daemon = True
        thread.start()
        
    def _process_data_thread(self):
        """Process data in background thread"""
        try:
            # Get options
            sheet_name = self.sheet_entry.get().strip() or None
            output_file = self.output_entry.get().strip() or None
            
            # If output file specified, ensure it has .xlsx extension
            if output_file and not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            
            # Import the generator class
            from mining_chart_generator import MiningDataChartGenerator
            
            # Update status
            self.root.after(0, lambda: self.status_label.config(text="Loading data..."))
            
            # Create generator
            self.generator = MiningDataChartGenerator(self.input_file, sheet_name)
            
            # Load data
            self.generator.load_data()
            
            # Update status
            self.root.after(0, lambda: self.status_label.config(text="Processing data..."))
            
            # Process data
            self.generator.process_data()
            
            # Update status
            self.root.after(0, lambda: self.status_label.config(text="Generating charts..."))
            
            # Create Excel with charts
            self.output_file = self.generator.create_excel_with_charts(output_file)
            
            # Success
            self.root.after(0, self._process_complete)
            
        except Exception as e:
            # Error
            self.root.after(0, lambda: self._process_error(str(e)))
            
    def _process_complete(self):
        """Handle successful completion"""
        self.progress.stop()
        self.browse_btn.config(state='normal')
        self.process_btn.config(state='normal')
        self.status_label.config(text="✅ Charts generated successfully!")
        
        # Show success message with option to open file
        result = messagebox.askyesno(
            "Success!",
            f"Charts have been successfully generated and embedded in Excel!\n\n"
            f"Output file: {os.path.basename(self.output_file)}\n\n"
            f"The file contains:\n"
            f"• Summary statistics\n"
            f"• Production overview charts\n"
            f"• Efficiency analysis charts\n"
            f"• Comparative analysis charts\n"
            f"• Trend analysis charts\n\n"
            f"Would you like to open the folder containing the file?"
        )
        
        if result:
            # Open folder containing the file
            folder = os.path.dirname(os.path.abspath(self.output_file))
            if os.name == 'nt':  # Windows
                os.startfile(folder)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{folder}"')
                
    def _process_error(self, error_msg):
        """Handle processing error"""
        self.progress.stop()
        self.browse_btn.config(state='normal')
        self.process_btn.config(state='normal')
        self.status_label.config(text="❌ Error occurred during processing")
        
        messagebox.showerror(
            "Processing Error",
            f"An error occurred while processing the data:\n\n{error_msg}\n\n"
            f"Please ensure:\n"
            f"• The Excel file contains mining data in the expected format\n"
            f"• Date columns are formatted as 'mon-YY' (e.g., ene-20, feb-20)\n"
            f"• The file is not corrupted or password-protected"
        )


def main():
    """Main function to run the GUI"""
    root = tk.Tk()
    app = MiningAnalyzerGUI(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
