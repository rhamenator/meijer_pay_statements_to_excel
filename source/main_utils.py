import os
import sys
import tkinter as tk
from tkinter import Label, Toplevel, messagebox

def os_supports_gui():
    return os.name in ['Windows', 'Darwin', 'Linux', 'FreeBSD', 'OpenBSD', 'NetBSD', 'SunOS']

def is_gui_mode():
    return os_supports_gui() and not sys.stdin.isatty()

def command_line_mode():
    return sys.stdin.isatty() or not os.environ.get('DISPLAY')

def clear_main_window_message(root):
    # Function to clear the message label in Tk.root
    # global root
    for widget in root.winfo_children():
        widget.destroy()

def show_in_main_window(root, title, message):
    # Update or create a label in Tk.root to display the message
    # global root
    if root is None:
        # print("Error: Tk.root is not initialized.")
        return
    # Clear any existing message label
    clear_main_window_message()

    # Create a label to display the message
    label_message = tk.Label(root, text=message, wraplength=300)
    label_message.pack(pady=20)

    # Optionally set a title for the main window
    root.title(title if title else "Messages:")

def show_temporary_message(message, duration=3000):
    top = Toplevel()
    top.title("Message")
    top.geometry("300x300")
    Label(top, text=message, wraplength=250).pack(expand=True)
    top.after(duration, top.destroy)

def print_message(message, title=None, user_input=False, dialog_type=None, duration=3000):
    # global gui_mode
    if command_line_mode():
        root = tk.Tk()
        root.withdraw()
    response = None
    if user_input:
        if is_gui_mode() or os_supports_gui():
            if dialog_type == "info" or dialog_type == None:
                messagebox.showinfo(title if title is not None else "Information", message)
            elif dialog_type == "showwarning":
                messagebox.showwarning(title if title is not None else "Warning", message)
            elif dialog_type == "showerror":
                messagebox.showerror(title if title is not None else "Error", message)
            elif dialog_type == "askokcancel":
                response = messagebox.askokcancel(title if title is not None else "Question", message)
            elif dialog_type == "askyesno":
                response = messagebox.askyesno(title if title is not None else "Question", message)
            elif dialog_type == "askyesnocancel":
                response = messagebox.askyesnocancel(title if title is not None else "Question", message)
            elif dialog_type == "askretrycancel":
                response = messagebox.askretrycancel(title if title is not None else "Problem", message)
            elif user_input:
                if root is None and is_gui_mode():
                    show_temporary_message(message, duration=3000)
                elif root is not None and is_gui_mode():
                    show_in_main_window(root, title, message)
                else: print(message)
            else: show_in_main_window(root, title, message)
        else:
            print(message)
    else:
        if root is None and is_gui_mode():
            show_temporary_message(message, duration=3000)
        elif root is not None and is_gui_mode():
            show_in_main_window(root, title, message)
        else: print(message)
    if command_line_mode() and root is not None: root.destroy()
    return response


def show_message(message, title=None, dialog_type=None):
    if dialog_type == "info" or dialog_type is None:
        messagebox.showinfo(title if title is not None else "Information", message)
    elif dialog_type == "showwarning":
        messagebox.showwarning(title if title is not None else "Warning", message)
    elif dialog_type == "showerror":
        messagebox.showerror(title if title is not None else "Error", message)