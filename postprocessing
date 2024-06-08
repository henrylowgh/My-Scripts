import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import re

# Function to extract sequence names
def parse_identifier(full_sequence_name):
    match = re.match(r'^#\s?([\w-]+?)-(b|a)(\d+)', full_sequence_name)
    if match:
        prefix = match.group(1)
        chain_type = match.group(2)
        number = match.group(3)
        return f"{prefix}-{number}", f"{prefix}-{chain_type}{number}", chain_type
    return None, None, None

def parse_and_reorder_blocks(file_path, output_directory):
    try:
        with open(file_path, 'r') as file:
            content = file.read().strip()

        if not content:
            raise ValueError("The file is empty or not properly formatted.")

        blocks = content.split('//')
        heavy_chains = {}
        light_chains = {}
        unmatched_heavy = []
        unmatched_light = []
        missing_sequence_data = []
        excluded_pairs = []
        pair_count = 0  # Counter for matched pairs

        # Identify and categorize each block
        for block in blocks:
            if not block.strip():
                continue  # Skip empty blocks
            lines = block.strip().split('\n')
            header = lines[0]  # The header line with the sequence name and type

            print(f"Processing block with header: {header}")  # Debug print

            # Check if there's sequence data
            sequence_present = any(" " in line and line.split()[1].isalpha() for line in lines if line.strip())
            print(f"Sequence present: {sequence_present}")  # Debug print

            # Extract IDs using the provided regular expression
            main_id, full_id, chain_type = parse_identifier(header)
            # print(f"Extracted IDs - Main ID: {main_id}, Full ID: {full_id}, Chain Type: {chain_type}")  # Debug print

            if chain_type == 'b':
                if sequence_present:
                    heavy_chains[main_id] = {'block': block + '//', 'name': header}
                else:
                    missing_sequence_data.append((header, 'Header present but empty sequence for heavy chain (meaning ANARCI failed to annotate)'))
            elif chain_type == 'a':
                if sequence_present:
                    light_chains[main_id] = {'block': block + '//', 'name': header}
                else:
                    missing_sequence_data.append((header, 'Header present but empty sequence for light chain (meaning ANARCI failed to annotate)'))

        # Ensure that only complete, annotated pairs are included
        reordered_content = ''
        for light_id, light_data in light_chains.items():
            if light_id in heavy_chains:
                reordered_content += light_data['block'].strip() + '\n' + heavy_chains[light_id]['block'].strip() + '\n'
                del heavy_chains[light_id]  # Remove matched heavy chain
                pair_count += 1
            else:
                unmatched_light.append((light_data['name'], 'Missing annotated heavy chain'))
                excluded_pairs.append((light_data['name'], 'Missing annotated heavy chain'))

        for heavy_id, heavy_data in heavy_chains.items():
            unmatched_heavy.append((heavy_data['name'], 'Missing annotated light chain'))
            excluded_pairs.append((heavy_data['name'], 'Missing annotated light chain'))

        # Create output filename based on input filename
        base_filename = os.path.basename(file_path)
        new_filename = os.path.splitext(base_filename)[0] + '_reordered' + os.path.splitext(base_filename)[1]
        output_path = os.path.join(output_directory, new_filename)

        # Write reordered content to a new file
        with open(output_path, 'w') as output_file:
            output_file.write(reordered_content)

        # Log the output
        log_filename = os.path.splitext(base_filename)[0] + '_log.txt'
        log_path = os.path.join(output_directory, log_filename)
        with open(log_path, 'w') as log_file:
            if unmatched_light or unmatched_heavy or missing_sequence_data:
                if unmatched_light:
                    log_file.write("\nUnmatched Light Chains:\n")
                    for name, reason in unmatched_light:
                        log_file.write(f"{name} - {reason}\n")
                if unmatched_heavy:
                    log_file.write("\nUnmatched Heavy Chains:\n")
                    for name, reason in unmatched_heavy:
                        log_file.write(f"{name} - {reason}\n")
                if missing_sequence_data:
                    log_file.write("\nEntries with Missing Sequence Data:\n")
                    for name, reason in missing_sequence_data:
                        log_file.write(f"{name} - {reason}\n")
                if excluded_pairs:
                    log_file.write("\nExcluded Pairs:\n")
                    for name, reason in excluded_pairs:
                        log_file.write(f"{name} - {reason}\n")
            log_file.write(f"\nTotal sequence pairs included in the output: {pair_count}\n")

        # Output unmatched chains, excluded pairs, and missing sequence data to the terminal
        if unmatched_light or unmatched_heavy or missing_sequence_data or excluded_pairs:
            if unmatched_light:
                print("\nUnmatched Light Chains:")
                for name, reason in unmatched_light:
                    print(f"{name} - {reason}")
            if unmatched_heavy:
                print("\nUnmatched Heavy Chains:")
                for name, reason in unmatched_heavy:
                    print(f"{name} - {reason}")
            if missing_sequence_data:
                print("\nEntries with Missing Sequence Data:")
                for name, reason in missing_sequence_data:
                    print(f"{name} - {reason}")
            if excluded_pairs:
                print("\nExcluded Pairs:")
                for name, reason in excluded_pairs:
                    print(f"{name} - {reason}")
        print(f"\nTotal sequence pairs included in the output: {pair_count}")

        messagebox.showinfo("Success", f"File reordered successfully. {pair_count} pairs included. Check the terminal and log file for issues.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def open_file_dialog():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title="Select the antibody sequence file")
    if file_path:
        output_directory = filedialog.askdirectory(title="Select output directory")
        if output_directory:
            parse_and_reorder_blocks(file_path, output_directory)
        else:
            messagebox.showwarning("Warning", "Output directory not selected.")
    else:
        messagebox.showwarning("Warning", "File not selected.")

# Trigger the file selection dialog
open_file_dialog()
