#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Feb 24 08:45:39 2025

@author: Michael Herschbach
"""

import pandas as pd

def sortImporter(file_path):
    
    df = pd.read_csv(file_path)
    
    # Grouping and aggregating data into lists for each team
    aggregated_df = df.groupby("Group Number").agg({
        "First Name": list,
        "Last Name": list,
        "Email": list,
        "Phone Number": list,
        "Team Leader": list,
        "Board Member": list,
        "Teacher": "first",
        "Day": "first", 
        "Start Time": "first",  
        "End Time": "first"
    }).reset_index()
    
    # Create extra columns
    aggregated_df["Teacher Email"] = ""
    aggregated_df["Teacher Phone"] = ""
    aggregated_df["School"] = ""
    
    # 6 lesson columns for each week
    for i in range(1, 7):  
        aggregated_df[f"Lesson {i}"] = ""  # Creates empty columns
    
    # Expanding function
    def expand_dataframe(df):
        expanded_rows = []
        
        for _, row in df.iterrows():
            max_len = max(len(row[col]) if isinstance(row[col], list) else 1 for col in df.columns)
            
            for i in range(max_len):
                new_row = {}
                
                for col in df.columns:
                    if isinstance(row[col], list):
                        new_row[col] = row[col][i] if i < len(row[col]) else None
                    else:
                        new_row[col] = "" if i > 0 else row[col]
                
                expanded_rows.append(new_row)
        
        return pd.DataFrame(expanded_rows)
    
    # Expanding the aggregated DataFrame to MAIN format
    expanded_df = expand_dataframe(aggregated_df)
    
    expanded_df.to_csv('MAIN_SHEET.csv', index=False)
    print("saved new MAIN sheet to 'MAIN_SHEET.csv'")
    
    return(expanded_df)
    
# import dataframes in the MAIN format (from above function) and aggregate
def mainImporter(file_path):
    df = pd.read_csv(file_path)
    df["Group Number"].fillna(method="ffill", inplace=True)
    grouped = df.groupby("Group Number")
    
    aggregated_data = []
    
    for group, data in grouped:
        aggregated_row = {"Group Number": group}
        
        for col in df.columns:
            if col == "Group Number":
                continue
            elif col in ["First Name", "Last Name", "Email", "Phone Number"]:
                aggregated_row[col] = list(data[col].dropna()) if data[col].notna().any() else []
            else:
                aggregated_row[col] = data[col].iloc[0]  # Keep single-entry columns as strings
        
        aggregated_data.append(aggregated_row)
    
    return pd.DataFrame(aggregated_data)
