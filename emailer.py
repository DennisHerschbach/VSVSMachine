#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Mar 20 16:42:29 2025

@author: Michael Herschbach
"""

import time
import tkinter as tk
from tkinter import messagebox
import keyboard

def send_email(driver, recipient, subject, message, cc_var, mac):
    # print(recipient)
    # print(subject)
    # print(message)
    
    try:
        time.sleep(2)  # Wait
        
        # Compose new email
        keyboard.press_and_release("n")
        time.sleep(3)
        
        # Enter each email followed by a comma
        for email in recipient:
            keyboard.write(email)
            time.sleep(1)
            keyboard.press_and_release(",")
            time.sleep(1)
        
        if cc_var.get():
            keyboard.press_and_release("tab")
            time.sleep(1)
            keyboard.press_and_release("tab")
            time.sleep(1)
                
        # Subject line
        keyboard.press_and_release("tab")
        time.sleep(1)
        keyboard.write(subject)
        time.sleep(1)
        
        # Message
        keyboard.press_and_release("tab")
        time.sleep(1)
        keyboard.write(message)
        time.sleep(1)
        
        tk.messagebox.showinfo("Confirmation", "Click OK when email has been sent.")
        
        #confirmation = tk.messagebox.askquestion("Confirmation",  "Send email?")
        
        # if confirmation:
        #     # Send
        #     if mac.get():
        #         keyboard.press_and_release("command+enter")  # macOS
        #     else:
        #         keyboard.press_and_release("ctrl+enter")  # Windows
                
        # else:
        #     raise Exception("User cancelled email.")

        #just wait to ensure email got sent 
        time.sleep(1)
        
    except Exception as e:
        print(f"Error: {e}")
    

def email_wizard(root, driver, df):
    
    def submit_email():
        group_list = entry_email.get().split(",")
        
        if all_teams.get():
            print("Emailing all teams")
            values_list = df['Group Number'].tolist()
            filtered_list = [x for x in values_list if x != 0]
            group_list = [int(x) for x in filtered_list]
            
        for group in group_list:
            get_emails = df.loc[df['Group Number'] == int(group), 'Email']
            teacher_email = df.loc[df['Group Number'] == int(group), 'Teacher Email'].values
            if not get_emails.empty:
                email_list = get_emails.values[0]
                
                # Add teacher email if box checked
                if include_teacher.get():
                    email_list.append(teacher_email[0])
                
                subject = entry_subject.get()
                message = text_message.get("1.0", tk.END).strip()
        
                # Setup variables for keywords
                team_number = group
                
                # Set up matrix (team leader, names, emails)
                first_names = df.loc[df['Group Number'] == int(group), 'First Name'].values
                last_names = df.loc[df['Group Number'] == int(group), 'Last Name'].values
                emails = df.loc[df['Group Number'] == int(group), 'Email'].values
                phones = df.loc[df['Group Number'] == int(group), 'Phone Number'].values
                
                matrix = "Team Leader: "
                for j in range(len(first_names[0])):
                    matrix += (first_names[0][j] + " " + last_names[0][j] + 
                               ", " + phones[0][j] + ", " + emails[0][j] + "\n")
                    
                # Set up teacher block
                grade_level = df.loc[df['Group Number'] == int(group), 'Grade Level'].values
                teacher = df.loc[df['Group Number'] == int(group), 'Teacher'].values
                teacher_email = df.loc[df['Group Number'] == int(group), 'Teacher Email'].values
                school = df.loc[df['Group Number'] == int(group), 'School'].values
                day = df.loc[df['Group Number'] == int(group), 'Day'].values
                start_time = df.loc[df['Group Number'] == int(group), 'Start Time'].values
                end_time = df.loc[df['Group Number'] == int(group), 'End Time'].values
                
                teacher_block = (
                    "Grade Level: " + str(int(grade_level[0])) + 
                    "\nTeacher: " + teacher[0] + "\nEmail: " + teacher_email[0] + 
                    "\nSchool: " + school[0] + "\nDay: " + day[0] + 
                    "\nTime: " + start_time[0] + " - " + end_time[0]
                )
    
                # Subject can ONLY handle team #
                subject = subject.format(team_number=team_number)
                message = message.format(team_number=team_number, matrix=matrix, tblock=teacher_block)
                
                if email_list and subject and message:
                    driver.switch_to.window(driver.window_handles[-1])
                    send_email(driver, email_list, subject, message, cc_var, mac)
                else:
                    messagebox.showwarning("Warning", "All fields must be filled out!")
            else:
                print(f"No emails found for group {group}")
                return
        
    
    # Create new window
    new_window = tk.Toplevel(root)
    new_window.title("Email Wizard")
    new_window.geometry("500x700")
    
    tk.Label(new_window, text="Group Number(s):").pack()
    entry_email = tk.Entry(new_window, width=50)
    entry_email.pack()

    tk.Label(new_window, text="Subject:").pack()
    entry_subject = tk.Entry(new_window, width=50)
    entry_subject.pack()

    tk.Label(new_window, text="Message:").pack()
    text_message = tk.Text(new_window, width=50, height=15)
    text_message.pack()

    tk.Button(new_window, text="Send Email", command=submit_email).pack()

    tk.Label(new_window, text="WARNING: Do not touch your keyboard or mouse while the email is being written").pack()
    
    # Options section
    tk.Label(new_window, text="Options:").pack(pady=(10, 0))
    
    # Teacher
    include_teacher = tk.BooleanVar(value=True)
    teach_checkbox = tk.Checkbutton(new_window, text="Include teacher", variable=include_teacher)
    teach_checkbox.pack()
    
    # Send email to ALL teams
    all_teams = tk.BooleanVar(value=False)
    all_checkbox = tk.Checkbutton(new_window, text="Email ALL teams (overrides Group # submission)", variable=all_teams)
    all_checkbox.pack()
    
    cc_var = tk.BooleanVar(value=True)
    cc_checkbox = tk.Checkbutton(new_window, text="CC/BCC skip (see documentation)", variable=cc_var)
    cc_checkbox.pack()
    
    # Mac
    mac = tk.BooleanVar(value=False)
    mac_checkbox = tk.Checkbutton(new_window, text="Check if using MacOS", variable=mac)
    mac_checkbox.pack()
    

####### FOR TESTING:

# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
# from selenium.webdriver.common.by import By
    
# from importer import mainImporter
# df = mainImporter('25AccessMASTER.csv')
        
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
# # driver.get("https://www.microsoft.com/en-us/microsoft-365/outlook/log-in")
# # time.sleep(5)  # Allow webpage to load
# # driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div[2]/main/div/div/div/div[6]/div/div/div[2]/div[1]/div/div/div[1]/div/div/div[3]/div/a[1]").click()

# # # Allow user to sign in manually
# # input("Press Enter after signing in...")
# # driver.switch_to.window(driver.window_handles[-1])
# # time.sleep(10)  # Allow time for inbox to load

# root = tk.Tk()
# root.title("Email Sender")
# root.geometry("500x300")

# def emailbutton():
#     email_wizard(root, driver, df)
#     print("Button clicked")

# button = tk.Button(root, text="Proceed", command=emailbutton)
# button.pack(pady=20)

# root.mainloop()
