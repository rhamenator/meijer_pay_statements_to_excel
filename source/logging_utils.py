import tkinter as tk
from tqdm import tqdm
from tkinter import messagebox
import time

class TqdmToText(tqdm):
    def __init__(self, *args, **kwargs):
        self.text_widget = kwargs.pop('text_widget', None)
        super().__init__(*args, **kwargs)

    def display(self, msg=None, pos=None):
        if self.text_widget:
            self.text_widget.insert(tk.END, self.format_meter(self.n, self.total, time.time() - self.start_t) + '\n')
            self.text_widget.see(tk.END)
        else:
            super().display(msg, pos)

class PrintLogger:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)

    def flush(self):
        pass

class ErrorLogger(PrintLogger):
    def showerror(self, message):
        messagebox.showerror("Error", message)