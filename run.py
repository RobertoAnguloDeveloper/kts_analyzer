#!/usr/bin/env python3
"""
Filepath: kts_analyzer/run.py
--------------------------------
KTS Analyzer - Main Application Entry Point

**No changes required to this file.**
The VCSM architecture ensures this entry point is
unaffected by the service-layer refactor.
--------------------------------
"""

import sys
import argparse
import os
import traceback

# --- GUI Availability Check ---
try:
    import tkinter as tk
    from tkinter import ttk  # Import ttk for themed widgets
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False
    print("Note: tkinter module not found. GUI mode is disabled.", file=sys.stderr)

# --- Import Application Components ---
# We only import the view if GUI is possible.
if GUI_AVAILABLE:
    from views.main_view import MainView

from controllers.main_controller import MainController


def main():
    """
    Main application entry point.
    Parses arguments and launches either CLI or GUI mode.
    """
    print("=" * 60)
    print("  KTS Data Analyzer (VCSM Architecture)")
    print("=" * 60)

    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(
        description="Analyze mining data and generate an Excel report with charts.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  GUI Mode (recommended):\n"
            "    python run.py\n\n"
            "  CLI Mode:\n"
            "    python run.py -i \"./data.xlsx\" -o \"./report.xlsx\" -s \"Sheet1\"\n"
        )
    )
    parser.add_argument(
        '-i', '--input',
        dest='input_file',
        help='Path to the source Excel file (e.g., "data/my_data.xlsx")'
    )
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        help='Path to save the generated report (e.g., "reports/my_report.xlsx")'
    )
    parser.add_argument(
        '-s', '--sheet',
        dest='sheet_name',
        help='Name of the Excel sheet to read (default: first sheet)'
    )
    
    args = parser.parse_args()

    # --- Controller Initialization ---
    try:
        controller = MainController()
    except Exception as e:
        print(f"Fatal Error: Failed to initialize application controller: {e}", file=sys.stderr)
        sys.exit(1)

    # --- Mode Selection ---
    if args.input_file and args.output_file:
        # --- CLI Mode ---
        # Both input and output are specified
        run_cli(controller, args)
        
    elif args.input_file or args.output_file:
        # Partial args, invalid for CLI
        parser.error("For CLI mode, both --input and --output are required.")
        
    else:
        # --- GUI Mode ---
        # No args provided
        run_gui(controller)

def run_cli(controller, args):
    """Launch the application in Command-Line Interface mode."""
    controller.run_cli(args.input_file, args.output_file, args.sheet_name)

def run_gui(controller):
    """Launch the application in Graphical User Interface mode."""
    if not GUI_AVAILABLE:
        print("\nGUI mode failed: tkinter is not installed.", file=sys.stderr)
        print("Please install tkinter (e.g., 'sudo apt-get install python3-tk')")
        print("or run in CLI mode (use 'python run.py --help' for info).")
        
        # Fallback to an interactive prompt
        print("\nGUI not available. Starting interactive CLI mode...")
        run_interactive_cli(controller)
        return

    try:
        # Use ThemedTk if available for a better look
        try:
            from ttkthemes import ThemedTk
            root = ThemedTk(theme="arc")
        except ImportError:
            root = tk.Tk()
            
        # Initialize the view and pass it the controller
        app = MainView(root, controller)
        
        # Start the Tkinter main loop
        root.mainloop()
        
    except Exception as e:
        print(f"\nAn error occurred while running the GUI: {e}", file=sys.stderr)
        traceback.print_exc()

def run_interactive_cli(controller: MainController):
    """
    A fallback interactive command-line prompt.
    """
    try:
        # Get Input File
        while True:
            input_file = input("Enter Excel file path: ").strip().strip('\"')
            if os.path.exists(input_file):
                break
            print(f"Error: File not found: {input_file}")

        # Get Sheet Name
        sheet = input("Enter sheet name (or press Enter for default): ").strip()
        sheet_name = sheet if sheet else None

        # Get Output File
        output = input("Enter output filename (or press Enter for default): ").strip().strip('\"')
        if not output:
            base, ext = os.path.splitext(input_file)
            output_file = f"{base}_Report.xlsx" # Force .xlsx
            print(f"Using default output: {output_file}")
        else:
            output_file = output

        # Run Analysis
        controller.run_cli(input_file, output_file, sheet_name)

    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")

if __name__ == "__main__":
    """
    Application entry point.
    """
    main()

