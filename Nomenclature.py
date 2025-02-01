import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys

class SplashScreen():
    def __init__(self, root):
        self.root = root
        self.root.title("Welcome")
        self.root.geometry("400x220")
        self.root.attributes('-topmost', True)
        self.root.lift()

        tk.Label(self.root, text="Nomenclature- Excel", font=("Arial", 24)).pack(pady=20)
        tk.Label(self.root, text="Credits:", font=("Arial", 16)).pack()
        tk.Label(self.root, text="Developed by Ananth Narayan", font=("Arial", 12)).pack()
        tk.Label(self.root, text="with help of ChatGPT and DeepSeek", font=("Arial", 8)).pack()

        self.progress = tk.Label(self.root, text="Loading...", font=("Arial", 12))
        self.progress.pack(pady=20)

        self.root.after(4000, self.close_splash)  # Close splash screen after 4 seconds

    def close_splash(self):
        self.root.destroy()
        open_main_window()  # Open main application window

def load_mapping():
    """Load the nomenclature mapping from an Excel file."""
    file_path = filedialog.askopenfilename(title="Select Nomenclature File", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return
    try:
        mapping_df = pd.read_excel(file_path, usecols=[0, 1], header=None)
        mapping_dict = dict(zip(mapping_df.iloc[:, 0], mapping_df.iloc[:, 1]))
        # messagebox.showinfo("Success", "Mapping file loaded successfully!")
        return mapping_dict
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load mapping: {e}")
        return None

def process_file(root):
    """Process the data file, replace headers, and save output."""
    mapping_dict = load_mapping()
    if not mapping_dict:
        return

    file_path = filedialog.askopenfilename(title="Select Excel data (.csv or .xlsx) to process", filetypes=[("Excel files", "*.xlsx"),("CSV files", "*.csv")])
    if not file_path:
        return

    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, header=0)
        else:
            df = pd.read_excel(file_path, header=0)

        # Flatten MultiIndex if necessary
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col for col in df.columns]

        # Filter and reorder columns based on the mapping
        filtered_columns = [col for col in mapping_dict if col in df.columns]
        df_filtered = df[filtered_columns]
        df_filtered.columns = [mapping_dict[col] for col in filtered_columns]

        save_path = filedialog.asksaveasfilename(title="Save converted excel data as", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            # Save without formatting
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                df_filtered.to_excel(writer, index=False, header=True)  # Write plain data without any formatting
            messagebox.showinfo("Success", f"File saved successfully and excel sheet will open now...")
            os.startfile(save_path)  # Open the file in Excel
            root.destroy()  # Close the main window
            sys.exit()  # Kill the script
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process file: {e}")
        root.destroy()  # Ensure the main window closes on error
        sys.exit()  # Kill the script

def open_main_window():
    """Open the main application window."""
    root = tk.Tk()
    root.title("Nomenclature Editor")
    root.geometry("400x75")

    btn_load_mapping = tk.Button(root, text="Load Nomenclature and Excel File", command=lambda: process_file(root))
    btn_load_mapping.pack(pady=20)

    root.protocol("WM_DELETE_WINDOW", lambda: (root.destroy(), sys.exit()))  # Ensure script exits when window is closed
    root.mainloop()

# Run the splash screen
splash_root = tk.Tk()
app = SplashScreen(splash_root)
splash_root.mainloop()
