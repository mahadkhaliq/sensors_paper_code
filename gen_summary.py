
"""
author: Mahad 
(https://www.linkedin.com/in/mahad-khaliq/)
(https://github.com/mahadkhaliq)

PM2.5 Data Analysis Tool
This script processes air quality data from Excel files to identify extreme PM2.5 values (e.g., 0 or >1000 µg/m³) and summarizes the results for each sensor. It provides insights into sensors that exhibit extreme readings on multiple days.

Features:
Data Loading: Reads PM2.5 data from multiple Excel files in a folder.
Extreme Value Detection: Identifies days where PM2.5 levels are 0 or exceed 1000 µg/m³.
Summary Table Generation: Produces a comprehensive summary of sensors with extreme readings over multiple days.
Output: Saves the summary as a CSV file for further analysis.
Usage:
Place all Excel files containing PM2.5 data in a folder.
Run the script and provide the folder path as input.
View or save the generated summary table.
This tool is designed for quick and efficient analysis of air quality sensor data, offering valuable insights for monitoring and reporting.
"""

import os
import pandas as pd

def load_and_find_extreme_values(file_path):
    try:
        data = pd.read_excel(file_path)
        data['Timestamp (UTC)'] = pd.to_datetime(data['Timestamp (UTC)'], errors='coerce')
        return data[['Timestamp (UTC)', 'PM2.5 (ug/m3)']].assign(File=os.path.basename(file_path))
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return pd.DataFrame()

def process_folder(folder_path):
    all_data = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            file_path = os.path.join(folder_path, file)
            data = load_and_find_extreme_values(file_path)
            all_data.append(data)

    return pd.concat(all_data, ignore_index=True)

def calculate_days_with_dates(data, condition):
    """
    Calculate the number of unique days and list specific days for each sensor based on a condition.
    """
    filtered_data = data[condition]
    filtered_data['Date'] = filtered_data['Timestamp (UTC)'].dt.date
    # Group by file and date
    daily_counts = filtered_data.groupby(['File', 'Date']).size().reset_index(name='Count')
    # Count unique days and list specific dates
    grouped = daily_counts.groupby('File')['Date']
    unique_day_counts = grouped.nunique()
    specific_dates = grouped.apply(lambda x: ', '.join(map(str, sorted(x.unique()))))
    return unique_day_counts[unique_day_counts > 1], specific_dates

def generate_summary_table(data):
    """
    Generate a summary table showing:
    - Number of days each sensor exceeded PM2.5 > 1000 for more than one day.
    - Number of days each sensor had PM2.5 = 0 for more than one day.
    - Specific dates for each condition.
    """
    # Calculate days and dates for PM2.5 > 1000
    high_values, high_dates = calculate_days_with_dates(data, data['PM2.5 (ug/m3)'] > 1000)
    # Calculate days and dates for PM2.5 = 0
    zero_values, zero_dates = calculate_days_with_dates(data, data['PM2.5 (ug/m3)'] == 0)

    # Combine results into a single table
    summary = pd.DataFrame({
        'Sensor': high_values.index,
        'Days with PM2.5 > 1000': high_values.values,
        'Dates with PM2.5 > 1000': high_dates
    }).merge(
        pd.DataFrame({
            'Sensor': zero_values.index,
            'Days with PM2.5 = 0': zero_values.values,
            'Dates with PM2.5 = 0': zero_dates
        }),
        on='Sensor',
        how='outer'
    ).fillna({'Days with PM2.5 > 1000': 0, 'Days with PM2.5 = 0': 0, 
              'Dates with PM2.5 > 1000': '', 'Dates with PM2.5 = 0': ''})
    
    # Convert to integer for clarity
    summary[['Days with PM2.5 > 1000', 'Days with PM2.5 = 0']] = summary[
        ['Days with PM2.5 > 1000', 'Days with PM2.5 = 0']
    ].astype(int)
    
    return summary

def main():
    folder_path = input("Enter the folder path containing the Excel files: ").strip()
    print("Processing files...")
    combined_data = process_folder(folder_path)

    if combined_data.empty:
        print("No data found in the provided folder.")
    else:
        summary_table = generate_summary_table(combined_data)
        if summary_table.empty:
            print("No sensors found with PM2.5 > 1000 or PM2.5 = 0 spanning more than one day.")
        else:
            print("\nSummary Table: PM2.5 > 1000 and PM2.5 = 0 for More Than One Day")
            print(summary_table)
            summary_table.to_csv("summary_table_with_dates.csv", index=False)
            print("\nSummary table saved as 'summary_table_with_dates.csv'.")

if __name__ == "__main__":
    main()
