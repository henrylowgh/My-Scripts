import pandas as pd
import tkinter as tk
from tkinter import filedialog

def load_file(prompt):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=prompt)
    return file_path

def save_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    directory = filedialog.askdirectory(title="Select output directory")
    return directory

def process_files(hybridoma_path, qc_path):
    # Load data
    hybridoma_data = pd.read_excel(hybridoma_path)
    qc_data = pd.read_excel(qc_path)

    # Clean up the data for accurate matching
    hybridoma_data['Azenta sequence ID'] = hybridoma_data['Azenta sequence ID'].str.strip()
    hybridoma_data['Unnamed: 6'] = hybridoma_data['Unnamed: 6'].str.strip()
    qc_data['DNAName'] = qc_data['DNAName'].str.strip()

    # Create a dictionary for Light Chain IDs
    light_chain_map = hybridoma_data.set_index('Azenta sequence ID')['Unnamed: 3'].to_dict()
    # Create a dictionary for Heavy Chain IDs
    heavy_chain_map = hybridoma_data.set_index('Unnamed: 6')['Unnamed: 3'].to_dict()

    # Append Clone# using Light Chain ID and Heavy Chain ID maps
    qc_data['Clone#'] = qc_data['DNAName'].apply(lambda x: light_chain_map.get(x, heavy_chain_map.get(x)))

    return qc_data

def main():
    hybridoma_path = load_file("Select the Hybridoma Excel File")
    qc_path = load_file("Select the QC Excel File")
    if not hybridoma_path or not qc_path:
        return print("Directory selection incomplete or incorrect file format, exiting the script.")

    updated_qc_data = process_files(hybridoma_path, qc_path)

    output_dir = save_file()
    output_path = f"{output_dir}/QC_Data_with_CloneNumbers.xlsx"
    updated_qc_data.to_excel(output_path, index=False)
    print(f"Updated file saved to {output_path}")

if __name__ == "__main__":
    main()
