#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Feb 25 20:38:09 2025

@author: aptitude
"""

import pandas as pd
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

def reportGen(df):
    
    #sort mon-fri, by times
    day_order = {'Monday': 1, 'Tuesday': 2, 'Wednesday': 3, 'Thursday': 4, 'Friday': 5}

    # Convert 'Start Time' to a datetime object
    df['Converted Time'] = pd.to_datetime(df['Start Time'], format='%I:%M %p')
    
    # Sort by 'Day' using a mapping, then by 'Converted Time'
    df = df.sort_values(by=['Day', 'Converted Time'], key=lambda x: x.map(day_order) if x.name == 'Day' else x)
    
    # Drop the converted time column
    df = df.drop(columns=['Converted Time'])

    def truncate_text(text, max_length):
        text = str(text)
        return text[:max_length] if len(text) > max_length else text
    
    def format_name(name):
        parts = str(name).split()
        return f"{parts[0][0]}. {parts[-1]}" if len(parts) > 1 else name
    
    def format_initial_lastname(first_names, last_names):
        if isinstance(first_names, list) and isinstance(last_names, list) and first_names and last_names:
            return f"{first_names[0][0].upper()}. {last_names[0]}"
        return ""
    
    PAGE_WIDTH, PAGE_HEIGHT = 612, 792
    LABEL_WIDTH, LABEL_HEIGHT = 288, 96
    MARGIN_X, MARGIN_Y, GAP_Y = 12, 60, 0
    LABELS_PER_COLUMN, LABELS_PER_PAGE = 7, 14
    
    def draw_label(c, x, y, row_data):
        c.rect(x, y, LABEL_WIDTH, LABEL_HEIGHT, stroke=1, fill=0)
        c.setFont("Helvetica", 10)
        text_x, text_y = x + 10, y + LABEL_HEIGHT - 20
        
        for line in row_data:
            if "**" in line:
                c.setFont("Helvetica-Bold", 10)
                c.drawString(text_x, text_y, line.replace("**", ""))
                c.setFont("Helvetica", 10)
            else:
                c.drawString(text_x, text_y, line)
            text_y -= 12
    
    def create_labels_from_csv(df, output_file, week):
        c = canvas.Canvas(output_file, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    
        for i, row in enumerate(df.itertuples(index=False)):
            nameHolder = []
            for j in range(len(row[1])):
                if j % 2 == 0:
                    nameString = ""
                    nameString = f"{nameString + row[1][j][0]}. {row[2][j]}, " 
                    # handle odd number teams
                    if j == len(row[1])-1:
                        nameString = nameString[:-2] 
                        nameHolder.append(nameString)
                else: 
                    nameString = f"{nameString + row[1][j][0]}. {row[2][j]}, " 
                    #remove last comma
                    nameString = nameString[:-2] 
                    nameHolder.append(nameString)
            
            nameHolder.extend(["", ""])
            tab = "            "
            label_text = [
                f"**Team {int(row[0])}**{tab}{tab}{row[8]} @ {row[9]}",
                f"**Lesson:** {row[14+week]}",
                f"Teacher: {format_name(row[7])}{tab}School: {truncate_text(row[13],10)}",
                f"Group: {nameHolder[0]}",
                f"{tab}{nameHolder[1]}",
                f"{tab}{nameHolder[2]}"
            ]
            
            label_index = i % LABELS_PER_PAGE
            row_in_page = label_index % LABELS_PER_COLUMN
            col = label_index // LABELS_PER_COLUMN
            x, y = MARGIN_X + col * (LABEL_WIDTH + MARGIN_X), PAGE_HEIGHT - MARGIN_Y - (row_in_page + 1) * (LABEL_HEIGHT + GAP_Y)
            
            if label_index == 0 and i != 0:
                c.showPage()
            
            draw_label(c, x, y, label_text)
        
        c.save()
        
    for week in range(6):
        create_labels_from_csv(df, f"label_w{week+1}.pdf", week)
    
    def create_table_pdf(df, output_pdf):
        df_copy = df.copy()
        
        df_copy[df_copy.columns[0]] = pd.to_numeric(df_copy[df_copy.columns[0]], errors='coerce').fillna(0).astype(int)
        df_copy.iloc[:, 7] = df_copy.iloc[:, 7].apply(format_name)
        df_copy.iloc[:, 2] = df_copy.iloc[:, 1].str[0].str[0] + ". " + df_copy.iloc[:, 2].str[0]
        df_copy.rename(columns={"Group Number": "#", "Last Name": "Team Lead", "Lesson 1" : "Week 1", 
                                "Lesson 2" : "Week 2", "Lesson 3" : "Week 3", "Lesson 4" : "Week 4", 
                                "Lesson 5" : "Week 5", "Lesson 6" : "Week 6"}, inplace=True)
        df_copy = df_copy.applymap(lambda x: truncate_text(x, 9)) #truncate all to 9 char
        
        df_copy.iloc[:, 8] = df_copy.iloc[:,8].astype(str).str[:3]   #specifically truncate days 
        df_copy = df_copy.iloc[:, [0, 2, 7, 8, 9, 10, 13, 14, 15, 16, 17, 18, 19]]
        
        data = [df_copy.columns.tolist()] + df_copy.values.tolist()
        pdf = SimpleDocTemplate(output_pdf, pagesize=landscape(letter))
        elements = [Table(data)]
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
        elements[0].setStyle(style)
        pdf.build(elements)
    
    create_table_pdf(df, "teams_table.pdf")
    
    def create_full_names_table_pdf(df, output_pdf):
        df_copy = df.copy()
        
        df_copy[df_copy.columns[0]] = pd.to_numeric(df_copy[df_copy.columns[0]], errors='coerce').fillna(0).astype(int)
        
        # Create a new DataFrame to store expanded rows
        expanded_rows = []
        
        # Iterate through each row (team)
        for _, row in df_copy.iterrows():
            group_num = row["Group Number"]
            teacher = format_name(row['Teacher'])
            day = row['Day']
            start_time = row['Start Time']
            end_time = row['End Time']
            school = truncate_text(row['School'], 10)
            week1 = row['Lesson 1']
            week2 = row['Lesson 2']
            week3 = row['Lesson 3']
            week4 = row['Lesson 4']
            week5 = row['Lesson 5']
            week6 = row['Lesson 6']
            
            # Format the team lead's name in "F. Lastname" format
            lead_name = f"{row['First Name'][0][0]}. {row['Last Name'][0]}"
            
            # First row contains all team information with the team lead
            team_row = {
                '#': group_num,
                'Name': lead_name,
                'Teacher': teacher,
                'Day': day,
                'Start': start_time,
                'End': end_time,
                'School': school,
                'Week 1': week1,
                'Week 2': week2,
                'Week 3': week3,
                'Week 4': week4,
                'Week 5': week5,
                'Week 6': week6
            }
            expanded_rows.append(team_row)
            
            # Add additional rows for each team member (except the first one/team lead)
            # with empty values for everything except names
            for i in range(1, len(row['First Name'])):
                # Format the team member's name in "F. Lastname" format
                member_name = f"{row['First Name'][i][0]}. {row['Last Name'][i]}"
                
                member_row = {
                    '#': '',
                    'Name': member_name,
                    'Teacher': '',
                    'Day': '',
                    'Start': '',
                    'End': '',
                    'School': '',
                    'Week 1': '',
                    'Week 2': '',
                    'Week 3': '',
                    'Week 4': '',
                    'Week 5': '',
                    'Week 6': ''
                }
                expanded_rows.append(member_row)
        
        # new DataFrame from the expanded rows
        expanded_df = pd.DataFrame(expanded_rows)
        
        expanded_df = expanded_df.applymap(lambda x: truncate_text(x, 9)) #truncate all to 9 char
        expanded_df.iloc[:, 3] = expanded_df.iloc[:,3].astype(str).str[:3]   #specifically truncate days 
        
        data = [expanded_df.columns.tolist()] + expanded_df.values.tolist() # prep for pdf
        
        pdf = SimpleDocTemplate(output_pdf, pagesize=landscape(letter))
        elements = [Table(data)]
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
        elements[0].setStyle(style)
        
        pdf.build(elements)
        
    create_full_names_table_pdf(df, "teams_table_full.pdf")
    
    def create_checkout(df):
        df_copy = df.copy()
        df_copy[df_copy.columns[0]] = pd.to_numeric(df_copy[df_copy.columns[0]], errors='coerce').fillna(0).astype(int)
        df_copy.iloc[:, 7] = df_copy.iloc[:, 7].apply(format_name)
        df_copy.iloc[:, 2] = df_copy.iloc[:, 1].str[0].str[0] + ". " + df_copy.iloc[:, 2].str[0]
        df_copy.rename(columns={"Group Number": "#", "Lesson 1": "Lesson", "Lesson 2": "Lesson", 
                                 "Lesson 3": "Lesson", "Lesson 4": "Lesson", "Lesson 5": "Lesson", "Lesson 6": "Lesson"}, inplace=True)
        df_copy = df_copy.applymap(lambda x: truncate_text(x, 15))
        
        df_copy["Box #"] = ""
        df_copy["Time In"] = ""
        df_copy["Additional Information                               "] = ""
        
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday"]
        styles = getSampleStyleSheet()
        
        for i in range(6):
            output_pdf = f"checkout_week{i+1}.pdf"
            pdf = SimpleDocTemplate(output_pdf, pagesize=landscape(letter))
            elements = []
            
            for weekday in weekdays:
                df_weekday = df_copy[df_copy.iloc[:, 8] == weekday]  # Assuming the weekday column is index 8
                
                if not df_weekday.empty:
                    df_checkout = df_weekday.iloc[:, [0, 9, 7, 13, 14 + i, 20, 21, 22]]
                    
                    data = [df_checkout.columns.tolist()] + df_checkout.values.tolist()
                    table = Table(data, hAlign='LEFT')
                    
                    style = TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ])
                    table.setStyle(style)
                    
                    table_title = "Week " + str(i+1) + " " + weekday
                    elements.append(Paragraph(table_title, styles['Title']))
                    elements.append(Spacer(1, 12))
                    elements.append(table)
                    elements.append(PageBreak())  # Start a new page for each weekday
            
            pdf.build(elements)

        
    create_checkout(df)
    

##### TEST CODE


from importer import mainImporter

df = mainImporter('25AccessMASTER.csv')
df.to_csv('test.csv', index=False)

reportGen(df)
