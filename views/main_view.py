"""
Filepath: kts_analyzer/views/main_view.py
--------------------------------------
KTS Analyzer - View (VCSM)

**No changes required to this file.**
The View is perfectly decoupled. It communicates with the
controller and is unaware of the service-layer implementation.
This is a key benefit of the VCSM pattern.
--------------------------
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import threading
import os

# Try to import ttkthemes for a modern look
try:
    from ttkthemes import ThemedTk
    THEMES_AVAILABLE = True
except ImportError:
    THEMES_AVAILABLE = False

class MainView:
    """
    The main GUI view for the KTS Analyzer application.
    Manages all UI components and delegates actions to the controller.
    """
    
    def __init__(self, root, controller):
        """
        Initialize the MainView.
        
        Args:
            root: The root Tkinter window (or ThemedTk instance).
            controller: The main application controller.
        """
        self.controller = controller
        self.controller.register_view(self)  # Link view to controller
        
        # --- Root Window Setup ---\
        self.root = root
        self.root.title("KTS Data Analyzer")
        self.root.geometry("600x450")
        self.root.minsize(500, 400)
        
        # Set a modern theme if available
        if THEMES_AVAILABLE and isinstance(root, ThemedTk):
            self.root.set_theme("arc")

        # --- Define Fonts ---
        self.title_font = font.Font(family="Segoe UI", size=16, weight="bold")
        self.label_font = font.Font(family="Segoe UI", size=11)
        self.button_font = font.Font(family="Segoe UI", size=11, weight="bold")
        self.status_font = font.Font(family="Courier New", size=9)

        # --- UI Component Variables ---
        self.input_file_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        
        # --- Create Main Frame ---
        self.main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Configure grid layout
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(4, weight=1)

        # --- UI Components ---
        
        # Title
        title_label = ttk.Label(self.main_frame, text="KTS Data Analyzer", font=self.title_font)
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky="W")

        # 1. Input File
        in_label = ttk.Label(self.main_frame, text="Input File:", font=self.label_font)
        in_label.grid(row=1, column=0, sticky="W", pady=5, padx=(0, 10))
        
        in_entry = ttk.Entry(self.main_frame, textvariable=self.input_file_var, width=60)
        in_entry.grid(row=1, column=1, sticky="EW", pady=5)
        
        in_button = ttk.Button(self.main_frame, text="Browse...", command=self.browse_input)
        in_button.grid(row=1, column=2, sticky="E", pady=5, padx=(10, 0))

        # 2. Output File
        out_label = ttk.Label(self.main_frame, text="Output File:", font=self.label_font)
        out_label.grid(row=2, column=0, sticky="W", pady=5, padx=(0, 10))
        
        out_entry = ttk.Entry(self.main_frame, textvariable=self.output_file_var, width=60)
        out_entry.grid(row=2, column=1, sticky="EW", pady=5)
        
        out_button = ttk.Button(self.main_frame, text="Browse...", command=self.browse_output)
        out_button.grid(row=2, column=2, sticky="E", pady=5, padx=(10, 0))

        # 3. Sheet Name (Optional)
        sheet_label = ttk.Label(self.main_frame, text="Sheet Name:", font=self.label_font)
        sheet_label.grid(row=3, column=0, sticky="W", pady=5, padx=(0, 10))
        
        sheet_entry = ttk.Entry(self.main_frame, textvariable=self.sheet_name_var, width=25)
        sheet_entry.grid(row=3, column=1, sticky="W", pady=5)
        sheet_info = ttk.Label(self.main_frame, text="(Optional: defaults to first sheet)")
        sheet_info.grid(row=3, column=1, sticky="W", padx=(190, 0))


        # 4. Status Box
        status_frame = ttk.LabelFrame(self.main_frame, text="Status Log", padding=10)
        status_frame.grid(row=4, column=0, columnspan=3, sticky="NSEW", pady=(20, 10))
        
        status_frame.rowconfigure(0, weight=1)
        status_frame.columnconfigure(0, weight=1)

        self.status_text = tk.Text(status_frame, height=10, width=70, 
                                   wrap=tk.WORD, font=self.status_font, 
                                   state=tk.DISABLED, relief=tk.FLAT)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text['yscrollcommand'] = scrollbar.set
        
        scrollbar.grid(row=0, column=1, sticky="NS")
        self.status_text.grid(row=0, column=0, sticky="NSEW")
        
        # Tag for error coloring
        self.status_text.tag_configure("error", foreground="red")

        # 5. Run Button
        self.run_button = ttk.Button(self.main_frame, text="Run Analysis", 
                                     command=self.start_analysis_thread, 
                                     style="Accent.TButton")
        self.run_button.grid(row=5, column=1, columnspan=2, sticky="E", pady=(10, 0))
        
        # Configure button style (if using ttk)
        try:
            style = ttk.Style()
            style.configure("Accent.TButton", font=self.button_font, padding=5)
        except tk.TclError:
            pass # Fails on some systems

    def browse_input(self):
        """Open file dialog to select input Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            self.input_file_var.set(file_path)
            
            # Suggest an output file path
            if not self.output_file_var.get():
                base, ext = os.path.splitext(file_path)
                output_path = f"{base}_Report.xlsx"
                self.output_file_var.set(output_path)

    def browse_output(self):
        """Open file dialog to select output Excel file."""
        file_path = filedialog.asksaveasfilename(
            title="Save Report As",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
            defaultextension=".xlsx"
        )
        if file_path:
            self.output_file_var.set(file_path)

    def start_analysis_thread(self):
        """
Moves the analysis to a separate thread to keep the GUI responsive.
        """
        in_file = self.input_file_var.get()
        out_file = self.output_file_var.get()
        sheet = self.sheet_name_var.get() or None # Use None if empty

        if not in_file or not out_file:
            messagebox.showerror("Missing Information", "Please specify both an input and output file.")
            return

        # Disable button, clear status
        self.run_button.config(state=tk.DISABLED, text="Running...")
        self.update_status("Starting analysis...", clear=True)
        
        # Run controller logic in a new thread
        analysis_thread = threading.Thread(
            target=self.controller.run_analysis,
            args=(in_file, out_file, sheet),
            daemon=True # Ensure thread closes if app is closed
        )
        analysis_thread.start()

    def update_status(self, message: str, error: bool = False, final: bool = False, clear: bool = False):
        """
        Updates the status text box. This method is thread-safe
        because the controller calls it via root.after().
        
        Args:
            message: The message to append.
            error: If True, style message as an error.
            final: If True, re-enable the Run button.
            clear: If True, clear the text box first.
        """
        self.status_text.config(state=tk.NORMAL)
        
        if clear:
            self.status_text.delete(1.0, tk.END)
            
        tag = "error" if error else "normal"
        self.status_text.insert(tk.END, f"{message}\n", tag)
        self.status_text.see(tk.END) # Auto-scroll
        
        self.status_text.config(state=tk.DISABLED)

        if final or error:
            self.run_button.config(state=tk.NORMAL, text="Run Analysis")
