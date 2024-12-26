"""
author: Mahad
(https://www.linkedin.com/in/mahad-khaliq/)
(https://github.com/mahadkhaliq)

PM2.5 Extreme Value Analysis Tool
This Python script is designed to analyze air quality data by identifying extreme PM2.5 levels (values of 0 or >1000 µg/m³). It processes Excel files containing PM2.5 data, extracts relevant extreme values, and visualizes the results using scatter plots.

Features
Data Loading: Automatically reads Excel files (.xls or .xlsx) from a specified folder.
Extreme Value Detection: Filters and identifies rows where PM2.5 levels are either 0 or greater than 1000 µg/m³.
Visualization: Generates a scatter plot with clear markers to visualize extreme values over time for each file.
Summary: Outputs a combined DataFrame with the detected extreme values, including timestamps and file names.
How It Works
Input: Specify the folder path containing the Excel files.
Processing: The script reads each file, extracts rows with extreme PM2.5 values, and combines the data from all files.
Output: Displays a summary of the extreme values and generates a scatter plot for visual analysis.
"""
import os
import pandas as pd
import matplotlib.pyplot as plt

def load_and_find_extreme_values(file_path):
    try:
        data = pd.read_excel(file_path)
        data['Timestamp (UTC)'] = pd.to_datetime(data['Timestamp (UTC)'], errors='coerce')
        extreme_values = data[(data['PM2.5 (ug/m3)'] == 0) | (data['PM2.5 (ug/m3)'] > 1000)]
        return extreme_values[['Timestamp (UTC)', 'PM2.5 (ug/m3)']].assign(File=os.path.basename(file_path))
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return pd.DataFrame()

def process_folder(folder_path):
    all_extreme_values = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') or file.endswith('.xls'):
            file_path = os.path.join(folder_path, file)
            extreme_values = load_and_find_extreme_values(file_path)
            all_extreme_values.append(extreme_values)

    combined_extreme_values = pd.concat(all_extreme_values, ignore_index=True)
    return combined_extreme_values

def plot_extreme_values(data):
    plt.figure(figsize=(14, 7))
    for file_name in data['File'].unique():
        file_data = data[data['File'] == file_name]
        plt.scatter(file_data['Timestamp (UTC)'], file_data['PM2.5 (ug/m3)'], label=file_name, s=10)

    plt.axhline(y=1000, color='r', linestyle='--', label='Threshold > 1000')
    plt.axhline(y=0, color='b', linestyle='--', label='PM2.5 = 0')
    plt.xlabel('Timestamp')
    plt.ylabel('PM2.5 (ug/m3)')
    plt.title('PM2.5 Extreme Values (0 or >1000) Across Files')
    plt.legend()
    plt.show()

def main():
    folder_path = input("Enter the folder path containing the Excel files: ").strip()
    print("Processing files...")
    combined_extreme_values = process_folder(folder_path)

    if combined_extreme_values.empty:
        print("No extreme PM2.5 values found in the provided folder.")
    else:
        print("\nSummary of Extreme PM2.5 Values:")
        print(combined_extreme_values)
        print("\nGenerating plot...")
        plot_extreme_values(combined_extreme_values)

if __name__ == "__main__":
    main()
