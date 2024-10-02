import tkinter as tk
import sys
import traceback
from typing import Self
from data_processing import main_logic
from file_utils import pick_input_file, pick_output_file
from logging_utils import ErrorLogger, PrintLogger

def run(input_file=None, output_file=None, text_box=None):
    return main_logic(input_file, output_file, text_box)

def create_gui(input_file=None, output_file=None):
    root = tk.Tk()
    root.title("Export pay statements to Excel")
    root.geometry("640x300")
    text_box = tk.Text(root, wrap='none', height=300, width=600)
    text_box.pack(side='left', fill='both', expand=True)
    scrollbar = tk.Scrollbar(root, command=text_box.yview)
    scrollbar.pack(side='right', fill='y')
    text_box.config(yscrollcommand=scrollbar.set)
    tk.Label(root, text="Input PDF File:").grid(row=0, column=0, padx=5, pady=5)
    tk.Entry(root, textvariable=input_file, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=pick_input_file).grid(row=0, column=2, padx=5, pady=5)
    tk.Label(root, text="Output Excel File:").grid(row=1, column=0, padx=5, pady=5)
    tk.Entry(root, textvariable=output_file, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=pick_output_file).grid(row=1, column=2, padx=5, pady=5)
    tk.Button(root, text="Process", command=Self.run(input_file, output_file, text_box)).grid(row=2, column=1, padx=5, pady=10)
    
    sys.stdout = PrintLogger(text_box)
    error_logger = ErrorLogger(text_box)  # Use ErrorLogger for error handling
    try:
        run(input_file, output_file)
        root.mainloop()
        return
    except KeyboardInterrupt:
        print("\nProgram terminated by user.")
        sys.exit(0)
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        if root:
            ErrorLogger.showerror(f"An error occurred:\n{e}")
            if root is not None: root.destroy()
        raise

