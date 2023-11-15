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
    xhtml_files = [f for f in os.listdir(oebps_folder) if f.endswith('.xhtml') and not f.endswith('-extracted.xhtml')][1:]
    all_weekly_programs = {}
    date_pattern = re.compile(r'\d{1,2}(\sDE\s[A-ZÑ]+|-\d{1,2})\s(?:A\s\d{1,2}\sDE\s[A-ZÑ]+)?')
    song_pattern = re.compile(r'Canción \d+')

    for file_name in xhtml_files:
        file_path = os.path.join(oebps_folder, file_name)
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        soup = BeautifulSoup(content, 'html.parser')
        headers = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        current_week_program = []
        current_week_title = None

        for header in headers:
            header_text = header.get_text(strip=True)
            if date_pattern.search(header_text):
                if current_week_title and current_week_program:
                    all_weekly_programs[current_week_title] = adjust_program_length(current_week_program)
                current_week_title = header_text
                current_week_program = []
            elif current_week_title:
                # Extract only the song part if it's a song line
                if 'Canción' in header_text and 'oración' in header_text:
                    song_match = song_pattern.search(header_text)
                    header_text = song_match.group() if song_match else header_text
                current_week_program.append(header_text)

        if current_week_title and current_week_program:
            all_weekly_programs[current_week_title] = adjust_program_length(current_week_program)

    return all_weekly_programs

# Helper function to adjust the length of each program section
def adjust_program_length(program):
    # Define the fixed lengths for each section
    fixed_lengths = {
        'SEAMOS MEJORES MAESTROS': 4,
        'NUESTRA VIDA CRISTIANA': 4  # Cambio de 5 a 4
    }

    for section, length in fixed_lengths.items():
        # Find the start and end index of the section
        section_indices = [i for i, item in enumerate(program) if section in item]
        if section_indices:
            start_idx = section_indices[0] + 1
            end_idx = start_idx
            section_items = []

            while end_idx < len(program) and not any(sec in program[end_idx] for sec in fixed_lengths if sec != section):
                item = program[end_idx]
                # For 'NUESTRA VIDA CRISTIANA', only include items with a leading number
                if section == 'NUESTRA VIDA CRISTIANA' and not item.startswith('Canción'):
                    if re.match(r'^\d+\.', item):
                        section_items.append(item)
                else:
                    section_items.append(item)
                end_idx += 1

            # Para 'NUESTRA VIDA CRISTIANA', si el último elemento es una canción, lo sacamos del conteo
            if section == 'NUESTRA VIDA CRISTIANA' and section_items[-1].startswith('Canción'):
                final_song = section_items.pop()

            # Calculate the number of items in the section
            section_length = len(section_items)
            # Insert empty cells if needed
            if section_length < length:
                section_items += [''] * (length - section_length)
            program[start_idx:end_idx] = section_items

            # Añade la canción al final de la sección 'NUESTRA VIDA CRISTIANA'
            if section == 'NUESTRA VIDA CRISTIANA':
                program.insert(end_idx, final_song)

    return program


# Function to format the weekly programs into a columnar structure for Excel
def format_weekly_programs_for_excel(all_weekly_programs):
    # Use the order of the weeks as they are in the dictionary
    sorted_weeks = list(all_weekly_programs.keys())

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
