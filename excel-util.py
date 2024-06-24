import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
from wordcloud import WordCloud
import os
import re

# Columns to ignore
IGNORE_COLUMNS = ['Department', 'Institution', 'Submitted on:', 'Username', 'Full name', 'Group', 'Course']

# Version
version = "0.0.3"

# Function to sanitize filenames
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename)

# Function to clean column titles for pie charts
def clean_title(title):
    return re.sub(r'->.*', '', title)

# Define a function to determine the predominant type of a column
def predominant_type(series):
    num_count = 0
    str_count = 0
    yes_no_count = 0

    for value in series:
        if pd.isnull(value):
            continue
        elif isinstance(value, (int, float)):
            num_count += 1
        elif isinstance(value, str):
            if value.lower() in ['yes', 'no', 'not yet']:
                yes_no_count += 1
            else:
                try:
                    float(value)
                    num_count += 1
                except ValueError:
                    str_count += 1

    if num_count > max(str_count, yes_no_count):
        return 'int'
    elif yes_no_count > max(num_count, str_count):
        return 'yes_no'
    else:
        return 'str'

# Define a function to clean and categorize columns
def categorize_columns(df):
    int_cols = []
    yes_no_cols = []
    str_cols = []

    for column in df.columns:
        if column in IGNORE_COLUMNS:
            continue
        col_type = predominant_type(df[column])
        if col_type == 'int':
            int_cols.append(column)
        elif col_type == 'yes_no':
            yes_no_cols.append(column)
        else:
            str_cols.append(column)

    int_data = df[int_cols].apply(pd.to_numeric, errors='coerce').dropna()
    yes_no_data = df[yes_no_cols]
    str_data = df[str_cols]
    
    return int_data, yes_no_data, str_data

def combine_excel_files(file_list, output_dir):
    # Initialize empty dataframes for each category
    int_df = pd.DataFrame()
    yes_no_df = pd.DataFrame()
    str_df = pd.DataFrame()

    # Process each file
    for file in file_list:
        if file.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        
        int_data, yes_no_data, str_data = categorize_columns(df)
        
        int_df = pd.concat([int_df, int_data], ignore_index=True)
        yes_no_df = pd.concat([yes_no_df, yes_no_data], ignore_index=True)
        str_df = pd.concat([str_df, str_data], ignore_index=True)

    # Write the combined data to a new Excel file with separate sheets
    output_file = os.path.join(output_dir, 'combined_data.xlsx')
    with pd.ExcelWriter(output_file) as writer:
        int_df.to_excel(writer, sheet_name='Integers', index=False)
        yes_no_df.to_excel(writer, sheet_name='YesNo', index=False)
        str_df.to_excel(writer, sheet_name='Strings', index=False)

    messagebox.showinfo("Success", f"Data combined successfully into {output_file}")

    # Generate charts
    generate_charts(int_df, yes_no_df, str_df, output_dir)

def generate_charts(int_df, yes_no_df, str_df, output_dir):
    # Create a folder to store the charts
    chart_folder = os.path.join(output_dir, 'charts')
    if not os.path.exists(chart_folder):
        os.makedirs(chart_folder)

    # Pie chart for each integer column
    for column in int_df.columns:
        if 'response' in column.lower() or 'id' in column.lower():
            continue
        sanitized_column = sanitize_filename(column)
        cleaned_title = clean_title(column)
        int_counts = int_df[column].value_counts()
        int_counts_percent = (int_counts / int_counts.sum()) * 100
        plt.figure(figsize=(8, 8))
        int_counts_percent.plot(kind='pie', autopct='%1.1f%%')
        plt.title(f'Distribution of {cleaned_title}')
        plt.savefig(os.path.join(chart_folder, f'{sanitized_column}_pie_chart.png'))
        plt.close()

    # Word cloud for each string column
    for column in str_df.columns:
        text = ' '.join(str(value) for value in str_df[column].dropna())
        if text.strip():  # Check if there is any text to generate the word cloud
            wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text)
            plt.figure(figsize=(10, 5))
            plt.imshow(wordcloud, interpolation='bilinear')
            plt.axis('off')
            plt.title(f'Word Cloud for {column}')
            sanitized_column = sanitize_filename(column)
            plt.savefig(os.path.join(chart_folder, f'{sanitized_column}_word_cloud.png'))
            plt.close()

    # Pie chart for each yes/no column
    for column in yes_no_df.columns:
        sanitized_column = sanitize_filename(column)
        cleaned_title = clean_title(column)
        yes_no_counts = yes_no_df[column].apply(lambda x: 'Yes' if str(x).lower() == 'yes' else 'No').value_counts()
        yes_no_counts_percent = (yes_no_counts / yes_no_counts.sum()) * 100
        plt.figure(figsize=(8, 8))
        yes_no_counts_percent.plot(kind='pie', autopct='%1.1f%%')
        plt.title(f'Distribution of {cleaned_title}')
        plt.savefig(os.path.join(chart_folder, f'{sanitized_column}_pie_chart.png'))
        plt.close()

    messagebox.showinfo("Success", "Charts generated successfully and saved in the 'charts' folder")

def select_files():
    file_list = filedialog.askopenfilenames(filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv")])
    if file_list:
        for file in file_list:
            selected_files.append(file)
        files_label.config(text="\n".join(selected_files))

def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory(title="Select Output Directory")
    if output_dir:
        output_dir_label.config(text=output_dir)

def start_work():
    if not selected_files:
        messagebox.showwarning("Warning", "No files selected!")
        return
    if not output_dir:
        messagebox.showwarning("Warning", "Output directory not selected!")
        return
    combine_excel_files(selected_files, output_dir)

# Initialize global variables
selected_files = []
output_dir = ""

# Create the main window
root = tk.Tk()
root.title("Excel and CSV Combiner")

# Add instructions
instructions = tk.Label(root, text="Welcome to the Excel and CSV Combiner.\n\n"
                                   "1. Click 'Select Excel and CSV Files' to choose your files.\n"
                                   "2. Click 'Select Output Directory' to choose where to save the combined file and charts.\n"
                                   "3. Click 'Start Work' to process the files and generate the combined Excel file and charts.\n"
                                   "\nMade by RealYusufIsmail",
                        justify=tk.LEFT, padx=10, pady=10)
instructions.pack()

# Version number
version_label = tk.Label(root, text="Version: " + version, justify=tk.LEFT, padx=10)
version_label.pack()

# Create and place buttons and labels
files_label = tk.Label(root, text="No files selected", justify=tk.LEFT, padx=10, pady=10)
files_label.pack()

button_select_files = tk.Button(root, text="Select Excel and CSV Files", command=select_files)
button_select_files.pack(pady=5)

output_dir_label = tk.Label(root, text="No output directory selected", justify=tk.LEFT, padx=10, pady=10)
output_dir_label.pack()

button_select_output_dir = tk.Button(root, text="Select Output Directory", command=select_output_dir)
button_select_output_dir.pack(pady=5)

button_start_work = tk.Button(root, text="Start Work", command=start_work)
button_start_work.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
