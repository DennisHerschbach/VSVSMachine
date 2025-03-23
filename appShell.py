#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Mar 20 16:42:29 2025

@author: Michael Herschbach
"""

import tkinter as tk
from tkinter import scrolledtext, filedialog
import sys
import time
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

# Import custom functions
from importer import sortImporter, mainImporter
from label import reportGen
from emailer import email_wizard

# Global variable to hold the dataframe
df = None

def email_instructions():
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get("https://www.microsoft.com/en-us/microsoft-365/outlook/log-in")
        time.sleep(5)  # Allow webpage to load
        
        # Click sign-in button
        driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div[2]/main/div/div/div/div[6]/div/div/div[2]/div[1]/div/div/div[1]/div/div/div[3]/div/a[1]").click()
        
    except Exception as e:
        print(f"Error: {e}")
        
    # Prompt user with message box, once hit, program proceeds
    messagebox.showinfo("Instructions", 'Sign in to Outlook in the new window. Once signed in, hit "OK"')
    
    email_wizard(root, driver, df)

# Print statements to show up in program console
class ConsoleRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)

    def flush(self):
        pass

# Sorter input
def button1_action():
    global df
    print("Beginning sorter import...")
    file_path = filedialog.askopenfilename(title="Select a file")
    if file_path:
        print(f"Selected file: '{file_path}'")
    df = sortImporter(file_path)
    print("Done with sorter import!")

# Main import
def button2_action():
    global df
    print("Beginning MAIN import...")
    file_path = filedialog.askopenfilename(title="Select a file")
    if file_path:
        print(f"Selected file: '{file_path}'")
    df = mainImporter(file_path)
    print("done!")

# Generate reports
def button3_action():
    global df
    if df is not None:
        print("Generating reports...")
        reportGen(df)
        print("Done! Check folder for reports.")
    else:
        print("No data to generate reports.")

#Start email Wizard
def button4_action():
    global df
    email_instructions()
    print("Opening Outlook sign-in...")

# placeholder
def button5_action():
    print("Button 5 clicked")

# Create main window
root = tk.Tk()
root.title("VSVS Machine")
root.geometry("600x400")

# Title
title_label = tk.Label(root, text="VSVS Machine", font=("Arial", 16))
title_label.pack()

# Frame for buttons on the left
button_frame = tk.Frame(root)
button_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.Y)

# Buttons and link to functions
buttons = [
    ("Import from Sorter", button1_action),
    ("Import from MAIN", button2_action),
    ("Generate Reports", button3_action),
    ("Email Wizard", button4_action),
    ("PLACEHOLDER", button5_action)
]
for text, command in buttons:
    btn = tk.Button(button_frame, text=text, command=command, width=15)
    btn.pack(pady=5)

# Frame for console output
console_frame = tk.Frame(root)
console_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

# Header label for the console
console_label = tk.Label(console_frame, text="Console", font=("Arial", 12))
console_label.pack()

# Scrolled text widget for console output
console_output = scrolledtext.ScrolledText(console_frame, width=50, height=20)
console_output.pack(fill=tk.BOTH, expand=True)

# Redirect print statements to console output
sys.stdout = ConsoleRedirector(console_output)

root.mainloop()
