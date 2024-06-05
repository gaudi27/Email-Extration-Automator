# -*- coding: utf-8 -*-
"""
Created on Tue Jun  4 13:38:50 2024

@author: George Audi
"""

import threading
import tkinter as tk
from tkinter import messagebox
import yaml
import os
import sys
import subprocess
from gmailParsing import start_parsing_emails

class EmailParserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Parser")
        self.root.geometry("500x300")

        self.label = tk.Label(root, text="Email Parser")
        self.label.pack(pady=10)

        self.user_label = tk.Label(root, text="Type Username with quotes around it:")
        self.user_label.pack(pady=5)
        self.user_entry = tk.Entry(root)
        self.user_entry.pack(pady=5)

        self.password_label = tk.Label(root, text="Type Password with quotes around it:")
        self.password_label.pack(pady=5)
        self.password_entry = tk.Entry(root, show="*")
        self.password_entry.pack(pady=5)

        self.save_button = tk.Button(root, text="Save Credentials", command=self.save_credentials)
        self.save_button.pack(pady=5)

        self.start_button = tk.Button(root, text="Start Parsing", command=self.start_parsing)
        self.start_button.pack(pady=5)

        self.open_button = tk.Button(root, text="Open Spreadsheet", command=self.open_spreadsheet)
        self.open_button.pack(pady=5)

        self.parsing = False
        self.thread = None

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def save_credentials(self):
        user = self.user_entry.get()
        password = self.password_entry.get()
        credentials = {"user": f'"{user}"', "password": f'"{password}"'}
        with open("usernameAndPassword.yml", "w") as f:
            yaml.dump(credentials, f)
        messagebox.showinfo("Email Parser", "Credentials saved.")

    def start_parsing(self):
        if not self.parsing:
            self.parsing = True
            self.thread = threading.Thread(target=start_parsing_emails)
            self.thread.daemon = True  # Allow thread to exit when main program exits
            self.thread.start()
            messagebox.showinfo("Email Parser", "Started parsing emails.")
        else:
            messagebox.showwarning("Email Parser", "Email parsing is already running.")

    def open_spreadsheet(self):
        file_path = "List of Emails to be used.xlsx"
        if os.path.exists(file_path):
            if os.name == 'nt':  # For Windows
                os.startfile(file_path)
            elif os.name == 'posix':  # For MacOS and Linux
                subprocess.call(('open' if sys.platform == 'darwin' else 'xdg-open', file_path))
        else:
            messagebox.showwarning("Email Parser", "Spreadsheet file not found.")

    def on_closing(self):
        self.parsing = False
        if self.thread and self.thread.is_alive():
            self.thread.join(1)
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailParserApp(root)
    root.mainloop()

