import os
import zipfile
import re
import shutil
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading
import webbrowser
from bs4 import BeautifulSoup
import pandas as pd
import sys

# Function to extract the content from the EPUB file
def extract_content_from_epub(epub_path):
    with zipfile.ZipFile(epub_path, 'r') as zip_ref:
        # Extract to a temporary directory
        temp_dir = 'extracted_epub'
        zip_ref.extractall(temp_dir)
    return temp_dir

# Function to extract the weekly programs from the extracted EPUB content
def extract_all_weekly_programs(extracted_folder):
    oebps_folder = os.path.join(extracted_folder, 'OEBPS')
    # List all XHTML files, skipping the cover (assumed to be the first file)
    xhtml_files = [f for f in os.listdir(oebps_folder) if f.endswith('.xhtml')][1:]

    # Initialize an empty dictionary to hold the programs
    all_weekly_programs = {}

    # Extract the date pattern for filtering relevant sections
    date_pattern = re.compile(r'\d{1,2}-\d{1,2}\sDE\s[A-ZÑ]+')

    # Process each XHTML file which contains the weekly program
    for file_name in xhtml_files:
        file_path = os.path.join(oebps_folder, file_name)
        # Read the content of the file
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        # Parse the file content using BeautifulSoup
        soup = BeautifulSoup(content, 'html.parser')
        # Find all header tags that could contain dates or section titles
        headers = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        # Temporary storage for current week's program
        current_week_program = []
        current_week_title = None
        for header in headers:
            header_text = header.get_text(strip=True)
            # Check if the header contains a date, which signifies a new week's program
            if date_pattern.search(header_text):
                # If we already have a week's program, store it before continuing to the next
                if current_week_title and current_week_program:
                    all_weekly_programs[current_week_title] = current_week_program
                current_week_title = header_text  # Update the current week's title
                current_week_program = []  # Reset the program list
            elif current_week_title:
                # If we are within a week's program, add the sections to the list
                current_week_program.append(header_text)
        # After finishing the file, store the last week's program
        if current_week_title and current_week_program:
            all_weekly_programs[current_week_title] = current_week_program

    return all_weekly_programs

# Function to format the weekly programs into a columnar structure for Excel
def format_weekly_programs_for_excel(all_weekly_programs):
    # Sort the weeks based on the date
    sorted_weeks = sorted(all_weekly_programs.keys(), key=lambda x: re.search(r'\d{1,2}-\d{1,2}', x).group())

    # Use zip to combine the weeks into a columnar format for Excel
    columns = zip(*[all_weekly_programs.get(week, [''] * max(len(p) for p in all_weekly_programs.values())) for week in sorted_weeks])

    # Create a DataFrame from the columns
    df = pd.DataFrame(columns, columns=sorted_weeks)

    return df

# Main function to process the EPUB and extract the programs in a formatted way for Excel
def extract_weekly_schedules_to_excel(epub_path, output_excel_file_path):
    # Extract the EPUB content to a temporary directory
    extracted_folder = extract_content_from_epub(epub_path)
    
    # Extract the weekly programs from the extracted content
    all_weekly_programs = extract_all_weekly_programs(extracted_folder)
    
    # Clean up the extracted files
    shutil.rmtree(extracted_folder)
    
    # Format the programs for Excel
    df_weekly_programs = format_weekly_programs_for_excel(all_weekly_programs)

    # Save the DataFrame to an Excel file
    df_weekly_programs.to_excel(output_excel_file_path, index=False)

# GUI functions
def handle_extraction():
    button['state'] = tk.DISABLED
    epub_file_path = filedialog.askopenfilename(title="Select EPUB file", filetypes=[("EPUB files", "*.epub")])
    if not epub_file_path:
        button['state'] = tk.NORMAL
        return
    
    # Create the thread and set it as a daemon
    extraction_thread = threading.Thread(target=extract_and_open_excel_file, args=(epub_file_path,))
    extraction_thread.daemon = True
    extraction_thread.start()

def extract_and_open_excel_file(epub_file_path):
    try:
        output_excel_file_path = 'weekly_programs.xlsx'
        extract_weekly_schedules_to_excel(epub_file_path,output_excel_file_path)
        print("Attempting to open Excel file")  # Debug print
        webbrowser.open(output_excel_file_path)
        print("Excel file should be open now")  # Debug print
    except Exception as e:
        error_message = str(e)
        print(f"Caught an exception: {error_message}")  # Debug print
        root.after(1, lambda: messagebox.showerror("Error", error_message))
    finally:
        print("Attempting to close the GUI")  # Debug print
        root.after(1, lambda: root.destroy() or sys.exit())   

root = tk.Tk()
root.title("VyM Extractor")
button = tk.Button(root, text="Selecciona EPUB", command=handle_extraction, height=3, width=30)
button.pack()  
root.mainloop()
