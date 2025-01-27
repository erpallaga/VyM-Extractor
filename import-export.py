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
from datetime import datetime, timedelta

# Function to extract the content from the EPUB file
def extract_content_from_epub(epub_path):
    with zipfile.ZipFile(epub_path, 'r') as zip_ref:
        # Extract to a temporary directory
        temp_dir = 'extracted_epub'
        zip_ref.extractall(temp_dir)
    return temp_dir

# Helper function to adjust the length of each program section
def adjust_program_length(program):
    # Define the fixed lengths for each section
    fixed_lengths = {
        'SEAMOS MEJORES MAESTROS': 4,
        'NUESTRA VIDA CRISTIANA': 4
    }

    for section, length in fixed_lengths.items():
        # Find the start and end index of the section
        section_indices = [i for i, item in enumerate(program) if section in item]
        if not section_indices:
            continue
        
        start_idx = section_indices[0] + 1
        end_idx = start_idx

        # Collect items for this section until next known section or the end
        while end_idx < len(program) and not any(
            sec in program[end_idx] for sec in fixed_lengths if sec != section
        ):
            end_idx += 1

        # Extract the lines that belong to this section
        section_items = program[start_idx:end_idx]

        # Look specifically for a final line that matches: Palabras de conclusión...|Canción nn
        final_song = None
        if section == 'NUESTRA VIDA CRISTIANA':
            final_song_pattern = re.compile(r'Palabras de conclusión.*\|Canción\s+(\d+)', re.IGNORECASE)
            for idx_line, line in enumerate(section_items):
                match = final_song_pattern.search(line)
                if match:
                    # Make sure group(1) actually exists
                    if match.lastindex and match.lastindex >= 1:
                        # Extract just "Canción NNN"
                        final_song = f"Canción {match.group(1)}"
                    # Remove that entire line from the section
                    section_items.pop(idx_line)
                    break

        # Pad or truncate items to the desired length
        if len(section_items) < length:
            section_items += [''] * (length - len(section_items))
        else:
            section_items = section_items[:length]

        # If we found a final song, add it after the main block
        if final_song:
            section_items.append(final_song)

        # Replace the old slice in 'program' with our adjusted items
        program[start_idx:end_idx] = section_items

    # Clean up: strip any "Canción..." lines to the simpler form
    for i, item in enumerate(program):
        if item.startswith('Canción'):
            m = re.match(r'(Canción\s*\d+)', item, re.IGNORECASE)
            if m:
                program[i] = m.group(1)

    return program

# Function to extract the weekly programs from the extracted EPUB content
def extract_all_weekly_programs(extracted_folder, target_weekday=1):  # 1 = martes por defecto
    """
    Parses each xhtml file to find blocks that match the headings for each
    week. For each heading that includes a date, we find the associated lines
    until the next heading. Then store them in a dict with keys = date string.
    """
    oebps_folder = os.path.join(extracted_folder, 'OEBPS')
    xhtml_files = [
        f for f in os.listdir(oebps_folder)
        if f.endswith('.xhtml') and not f.endswith('-extracted.xhtml')
    ][1:]
    all_weekly_programs = {}

    # Regex to extract the date from headings
    date_pattern = re.compile(r'(\d{1,2})\sDE\s([A-ZÑ]+)|(\d{1,2})-(\d{1,2})\sDE\s([A-ZÑ]+)', re.IGNORECASE)
    
    # Mapping from month name to month number
    month_mapping = {
        "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
        "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
        # If your EPUB might contain other months/abbreviations, add them here
    }
    current_year = datetime.now().year
    current_month = datetime.now().month

    for file_name in xhtml_files:
        file_path = os.path.join(oebps_folder, file_name)
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        soup = BeautifulSoup(content, 'html.parser')
        headers = soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
        current_week_program = []
        current_week_title = None
        formatted_date = None

        for header in headers:
            header_text = header.get_text(strip=True)
            date_match = date_pattern.search(header_text)
            if date_match:
                # This header seems to contain a date
                if date_match.group(3):  # e.g. "1-7 DE ENERO"
                    day = date_match.group(3)
                    month_str = date_match.group(5)
                else:  # e.g. "1 DE ENERO"
                    day = date_match.group(1)
                    month_str = date_match.group(2)

                if not month_str:
                    continue  # If something's off and we didn't capture the month at all, skip

                # Convert month string to uppercase
                upper_month = month_str.upper()
                # Attempt to retrieve month number
                program_month = month_mapping.get(upper_month)
                if not program_month:
                    # We found an unknown month - skip this entire "week"
                    print(f"Skipping unknown month '{month_str}' in heading: '{header_text}'")
                    current_week_title = None
                    current_week_program = []
                    continue

                # If we get here, we have a valid month
                program_year = current_year + 1 if program_month < current_month else current_year
                initial_date = datetime(program_year, program_month, int(day))
                
                # Calculate which weekday that date is (Mon=0, Tue=1, etc.)
                day_of_week = initial_date.weekday()
                # Adjust to the desired target weekday
                days_to_add = (target_weekday - day_of_week) % 7
                adjusted_date = initial_date + timedelta(days=days_to_add)
                
                formatted_date = adjusted_date.strftime('%d/%m/%Y')
                sortable_date = adjusted_date.strftime('%Y-%m-%d')  # for sorting
                current_week_title = sortable_date
                current_week_program = []

            elif current_week_title:
                # If we're already inside a recognized week, gather the lines
                current_week_program.append(header_text)

        # Once we exit the loop, if we found a week block, store it
        if current_week_title and current_week_program and formatted_date:
            all_weekly_programs[current_week_title] = {
                'date': formatted_date,
                'program': adjust_program_length(current_week_program)
            }

    # Sort the week blocks by date
    sorted_programs = {
        v['date']: v['program'] 
        for k, v in sorted(all_weekly_programs.items())
    }

    return sorted_programs

# Function to format the weekly programs into a columnar structure for Excel
def format_weekly_programs_for_excel(all_weekly_programs):
    """
    Takes a dict of { date_str: [items] } and makes a DataFrame.
    Each column is one date/week, each row is the Nth line item.
    """
    # Use the order of the weeks as sorted by date
    sorted_weeks = list(all_weekly_programs.keys())

    # Figure out the longest "program" among the weeks
    max_length = max(len(p) for p in all_weekly_programs.values()) if all_weekly_programs else 0

    # Build columns by zipping each program's lines
    columns = []
    for week in sorted_weeks:
        program_lines = all_weekly_programs[week]
        # Pad short programs so all columns are equal length
        if len(program_lines) < max_length:
            program_lines += [''] * (max_length - len(program_lines))
        columns.append(program_lines)

    # Transpose columns -> rows
    rows = list(zip(*columns))
    df = pd.DataFrame(rows, columns=sorted_weeks)

    return df

# Main function to process the EPUB and extract the programs in a formatted way for Excel
def extract_weekly_schedules_to_excel(epub_path, output_excel_file_path):
    # Extract the EPUB content to a temporary directory
    extracted_folder = extract_content_from_epub(epub_path)
    
    # Extract the weekly programs from the content
    all_weekly_programs = extract_all_weekly_programs(extracted_folder)
    
    # Clean up the extracted files
    shutil.rmtree(extracted_folder)
    
    # Format the programs for Excel
    df_weekly_programs = format_weekly_programs_for_excel(all_weekly_programs)

    # Save the DataFrame to an Excel file
    df_weekly_programs.to_excel(output_excel_file_path, index=False)

# GUI-related functions
def handle_extraction():
    button['state'] = tk.DISABLED
    epub_file_path = filedialog.askopenfilename(
        title="Select EPUB file",
        filetypes=[("EPUB files", "*.epub")]
    )
    if not epub_file_path:
        button['state'] = tk.NORMAL
        return
    
    extraction_thread = threading.Thread(target=extract_and_open_excel_file, args=(epub_file_path,))
    extraction_thread.daemon = True
    extraction_thread.start()

def extract_and_open_excel_file(epub_file_path):
    try:
        output_excel_file_path = 'weekly_programs.xlsx'
        extract_weekly_schedules_to_excel(epub_file_path, output_excel_file_path)
        print("Attempting to open Excel file")
        webbrowser.open(output_excel_file_path)
        print("Excel file should be open now")
    except Exception as e:
        error_message = str(e)
        print(f"Caught an exception: {error_message}")
        root.after(1, lambda: messagebox.showerror("Error", error_message))
    finally:
        print("Attempting to close the GUI")
        root.after(1, lambda: root.destroy() or sys.exit())

# Build the simple Tkinter GUI
root = tk.Tk()
root.title("VyM Extractor")
button = tk.Button(root, text="Selecciona EPUB", command=handle_extraction, height=3, width=30)
button.pack()
root.mainloop()
