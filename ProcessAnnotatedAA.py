import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

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
        pair_count = 0  # Counter for matched pairs

        # Identify and categorize each block
        for block in blocks:
            if not block.strip():
                continue  # Skip empty blocks
            lines = block.strip().split('\n')
            header = lines[0]  # The header line with the sequence name and type

            # Check if there's sequence data
            sequence_present = any(" " in line and line.split()[1].isalpha() for line in lines if line.strip())

            if 'TDM-1-H' in header:
                heavy_id = header.split('-H')[1].split('-')[0]
                if sequence_present:
                    heavy_chains[heavy_id] = {'block': block + '//', 'name': header}
                else:
                    missing_sequence_data.append((header, 'Empty sequence for heavy chain'))
            elif 'TDM-1-L' in header:
                light_id = header.split('-L')[1].split('-')[0]
                if sequence_present:
                    light_chains[light_id] = {'block': block + '//', 'name': header}
                else:
                    missing_sequence_data.append((header, 'Empty sequence for light chain'))

        # Ensure that only complete pairs are included
        reordered_content = ''
        for light_id, light_data in light_chains.items():
            if light_id in heavy_chains:
                reordered_content += light_data['block'].strip() + '\n' + heavy_chains[light_id]['block'].strip() + '\n'
                del heavy_chains[light_id]  # Remove matched heavy chain
                pair_count += 1
            else:
                unmatched_light.append((light_data['name'], 'Missing heavy chain'))

        unmatched_heavy.extend((h['name'], 'Missing light chain') for h in heavy_chains.values())

        # Create output filename based on input filename
        base_filename = os.path.basename(file_path)
        new_filename = os.path.splitext(base_filename)[0] + '_reordered' + os.path.splitext(base_filename)[1]
        output_path = os.path.join(output_directory, new_filename)

        # Write reordered content to a new file
        with open(output_path, 'w') as output_file:
            output_file.write(reordered_content)

        # Output unmatched chains and missing sequence data to the terminal
        if unmatched_light or unmatched_heavy or missing_sequence_data:
            print("Unmatched or Missing Data Detected:")
            if unmatched_light:
                print("Unmatched Light Chains:")
                for name, reason in unmatched_light:
                    print(f"{name} - {reason}")
            if unmatched_heavy:
                print("Unmatched Heavy Chains:")
                for name, reason in unmatched_heavy:
                    print(f"{name} - {reason}")
            if missing_sequence_data:
                print("Entries with Missing Sequence Data:")
                for name, reason in missing_sequence_data:
                    print(f"{name} - {reason}")
        print(f"Total sequence pairs included in the output: {pair_count}")

        messagebox.showinfo("Success", f"File reordered successfully. {pair_count} pairs included. Check the terminal for issues.")
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
