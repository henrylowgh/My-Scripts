import pandas as pd
import plotly.graph_objects as go
import tkinter as tk
from tkinter import filedialog

def create_sortable_table(df):
    fig = go.Figure(data=[go.Table(
        header=dict(values=list(df.columns),
                    fill_color='paleturquoise',
                    align='left'),
        cells=dict(values=[df[col] for col in df.columns],
                   fill_color='lavender',
                   align='left'))
    ])
    fig.show()

def main():
    # Set up the root for Tkinter
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    # Open file dialog and get the file path
    file_path = filedialog.askopenfilename(
        title="Select file", 
        filetypes=[("CSV Files", "*.csv")]
    )

    # Load and plot the data if a file was selected
    if file_path:
        df = pd.read_csv(file_path)

        # Creating an interactive table
        create_sortable_table(df)
    else:
        print("No file was selected")

if __name__ == "__main__":
    main()
