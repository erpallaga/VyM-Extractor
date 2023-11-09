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

# Function to save the weekly programs to a text file
def save_weekly_programs_to_text(weekly_programs, text_file_path):
    with open(text_file_path, 'w', encoding='utf-8') as file:
        for week, program in weekly_programs.items():
            file.write(f'{week}:\n')
            file.write('\n'.join(program))
            file.write('\n\n')

# Main function to process the EPUB and extract the programs
def extract_weekly_schedules(epub_path, output_file_path):
    # Extract the EPUB content to a temporary directory
    extracted_folder = extract_content_from_epub(epub_path)
    
    # Extract the weekly programs from the extracted content
    all_weekly_programs = extract_all_weekly_programs(extracted_folder)
    
    # Clean up the extracted files
    shutil.rmtree(extracted_folder)
    
    # Save the extracted programs to a text file
    save_weekly_programs_to_text(all_weekly_programs, output_file_path)

# GUI functions
def handle_extraction():
    button['state'] = tk.DISABLED
    epub_file_path = filedialog.askopenfilename(title="Select EPUB file", filetypes=[("EPUB files", "*.epub")])
    if not epub_file_path:
        button['state'] = tk.NORMAL
        return
    
    # Create the thread and set it as a daemon
    extraction_thread = threading.Thread(target=extract_and_open_text_file, args=(epub_file_path,))
    extraction_thread.daemon = True
    extraction_thread.start()

def extract_and_open_text_file(epub_file_path):
    try:
        output_text_file_path = 'weekly_programs.txt'
        extract_weekly_schedules(epub_file_path, output_text_file_path)
        webbrowser.open(output_text_file_path)
    except Exception as e:
        # If there's an error, show it in a message box in the main thread
        root.after(1, lambda: messagebox.showerror("Error", str(e)))
    finally:
        # Schedule the GUI to close in the main thread
        root.after(1, lambda: root.destroy() or sys.exit())

root = tk.Tk()
root.title("VyM Extractor")
button = tk.Button(root, text="Selecciona EPUB", command=handle_extraction, height=3, width=30)
button.pack()  
root.mainloop()
