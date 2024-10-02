import os
import time
from tkinter import filedialog
from main_utils import command_line_mode
import keyboard
import tkinter as tk
from tkinter import messagebox
from main_utils import is_gui_mode, os_supports_gui, print_message

def is_pdf(file_path):
    if not os.path.exists(file_path):
        return False
    with open(file_path, 'rb') as file:
        header = file.read(8)
    return header.startswith(b'%PDF-')

def pick_input_file_gui():
    if command_line_mode():
        root = tk.Tk()
        root.withdraw()
    file_path = filedialog.askopenfilename(defaultextension=".pdf", title="Select the PDF file containing your pay statements", filetypes=[("PDF Files", "*.pdf")])
    if command_line_mode():
        if root is not None: root.destroy()
    return file_path

def pick_input_file_cli():
    file_path = input("Enter the path to the PDF file containing your pay statements: ")
    return file_path.strip('\"')

def pick_output_file_gui(input_file=None):
    if command_line_mode():
        root = tk.Tk()
        root.withdraw()
    file_path = filedialog.asksaveasfilename(confirmoverwrite=True, initialfile=os.path.basename(input_file).replace('.pdf', '.xlsx'), defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Select the name of the Excel you want to create")
    if command_line_mode():
        if root is not None: root.destroy()
    return file_path

def pick_output_file_cli(input_file=None):
    file_path = input("Enter the name of the Excel workbook you want to create, \n or press [Enter] to use the default name: ")
    if not file_path:
        if input_file is None:
            return None
        return input_file.replace('.pdf', '.xlsx').strip('\"')
    else:
        return file_path.strip('\"')

def pick_input_file(file_path):
    # global gui_mode
    while True:
        
        if file_path is None:
            # file_path = pick_input_file_gui() if os_supports_gui() else pick_input_file_cli()
            file_path = pick_input_file_gui() if os_supports_gui() or is_gui_mode() else pick_input_file_cli()
        
        if not file_path: return None
        
        if os.path.exists(file_path) and not is_pdf(file_path):
            print_message(f"{file_path} is not a PDF file.\nPlease {'select' if is_gui_mode() else 'enter the name of'} a PDF file.")
            file_path = None
            continue
           
        if os.path.exists(file_path):
            if not is_gui_mode(): print_message(f"File selected: {file_path}")
            return file_path
        else:
            if os_supports_gui():
                if command_line_mode():
                    root = tk.Tk()
                    root.withdraw()
                answer = messagebox.askretrycancel('File selected does not exist. Click Retry to try again, or Cancel to quit')
                if command_line_mode() and root is not None: root.destroy()
                if answer == 'retry':
                    file_path = ''
                    continue
                else:
                    return None
            else:
                print_message("File does not exist. Press any key to try again, or press [Enter], [ctrl/command]+[c/z], or [Esc] to quit.", user_input=True)
                event = keyboard.read_event()
                if event.event_type == keyboard.KEY_DOWN and event.name in ('ctrl+c', 'esc', 'enter'):
                    return None
                else:
                    file_path = ''
                    continue
                
def pick_output_file(input_file, file_path):
    # global gui_mode    
    while True:
        if file_path is None:
            if command_line_mode():
                root = tk.Tk()
                root.withdraw()
            file_path = pick_output_file_gui(input_file) if os_supports_gui() or is_gui_mode() else pick_output_file_cli(input_file)
            if command_line_mode() and root is not None: root.destroy()
        if not file_path:
            return None
        if not os.path.exists(file_path):
            return file_path
        else:
            if os_supports_gui():
                return file_path
            else:
                overwrite = input(f"File {file_path} already exists. Do you want to overwrite it? (Press [y]='yes' or [n]=no, default=[n]): ").lower()
                if overwrite == 'y':
                    return file_path
                else:
                    file_path = None
                    continue

def is_file_locked(filepath):
    if not os.path.exists(filepath):
        return False
    try:
        with open(filepath, 'a'):
            pass
    except IOError:
        return True
    return False

def file_lock_wait(filename):
    while is_file_locked(filename):
        if is_gui_mode():
            err_message = f'{os.path.basename(filename)} is locked by another application. Please close it and retry or cancel'
            response = messagebox.askretrycancel(title='Error', message=err_message)
            if response == 'cancel' or response == None:
                return False
        else:
            print_message(f'{os.path.basename(filename)} is locked by another application. Please close it then press a key, or press "ctrl+q" or "esc" to quit...', user_input=True)
            event = keyboard.read_event()
            if event.event_type == keyboard.KEY_DOWN and event.name in ('ctrl+q', 'esc'):
                return False
            time.sleep(5)
    return True
