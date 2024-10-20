import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Function to load Excel file
def load_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {e}")
        return None

# Function to show a preview of files to be renamed
def preview_files(directory, excel_data, file_types):
    preview_listbox.delete(0, tk.END)  # Clear the preview list
    if excel_data is None:
        messagebox.showerror("Error", "No Excel data found")
        return

    for index, row in excel_data.iterrows():
        try:
            prd_code = row['prd_code']
            prd_name = row['prd_name']
            number = row['number']
            new_name = f"{prd_code} {prd_name} {number}"

            # Search for the matching file in the directory based on file type
            for file in os.listdir(directory):
                file_ext = os.path.splitext(file)[1].lower()
                if prd_name in file and file_ext in file_types:
                    preview_listbox.insert(tk.END, f"{file} -> {new_name}{file_ext}")
                    break
        except Exception as e:
            messagebox.showerror("Error", f"Failed to preview files: {e}")
            return

# Function to rename files in the directory
def rename_files(directory, excel_data, file_types):
    if excel_data is None:
        messagebox.showerror("Error", "No Excel data found")
        return

    progress_bar['value'] = 0
    total_files = len(excel_data)
    
    for index, row in excel_data.iterrows():
        try:
            prd_code = row['prd_code']
            prd_name = row['prd_name']
            number = row['number']
            new_name = f"{prd_code} {prd_name} {number}"

            # Search for the matching file in the directory based on file type
            for file in os.listdir(directory):
                file_ext = os.path.splitext(file)[1].lower()
                if prd_name in file and file_ext in file_types:
                    old_file_path = os.path.join(directory, file)
                    new_file_path = os.path.join(directory, f"{new_name}{file_ext}")
                    os.rename(old_file_path, new_file_path)
                    break

            # Update the progress bar
            progress_bar['value'] = ((index + 1) / total_files) * 100
            app.update_idletasks()  # Refresh the UI

        except Exception as e:
            messagebox.showerror("Error", f"Failed to rename file: {e}")
            return
    
    messagebox.showinfo("Success", "Files renamed successfully!")

# Function to open file dialog for Excel selection
def select_excel():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
    )
    if file_path:
        excel_path_entry.delete(0, tk.END)
        excel_path_entry.insert(0, file_path)

# Function to open directory dialog
def select_directory():
    directory = filedialog.askdirectory(title="Select Directory")
    if directory:
        directory_path_entry.delete(0, tk.END)
        directory_path_entry.insert(0, directory)

# Main function to handle the preview and renaming process
def start_preview():
    excel_path = excel_path_entry.get()
    directory_path = directory_path_entry.get()
    file_types = file_type_entry.get().split()

    if not excel_path or not directory_path:
        messagebox.showerror("Error", "Please select both Excel file and Directory")
        return

    excel_data = load_excel(excel_path)
    preview_files(directory_path, excel_data, file_types)

def start_renaming():
    excel_path = excel_path_entry.get()
    directory_path = directory_path_entry.get()
    file_types = file_type_entry.get().split()

    if not excel_path or not directory_path:
        messagebox.showerror("Error", "Please select both Excel file and Directory")
        return

    excel_data = load_excel(excel_path)
    rename_files(directory_path, excel_data, file_types)

# Create the main application window
app = tk.Tk()
app.title("File Renaming Application with Preview")
app.geometry("500x400")

# Create labels, entries, and buttons for the UI
tk.Label(app, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=10)
excel_path_entry = tk.Entry(app, width=40)
excel_path_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=select_excel).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Select Directory:").grid(row=1, column=0, padx=10, pady=10)
directory_path_entry = tk.Entry(app, width=40)
directory_path_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=select_directory).grid(row=1, column=2, padx=10, pady=10)

tk.Label(app, text="File Types (e.g. .jpg .png):").grid(row=2, column=0, padx=10, pady=10)
file_type_entry = tk.Entry(app, width=40)
file_type_entry.grid(row=2, column=1, padx=10, pady=10)

# Create a preview listbox to show old and new file names
tk.Label(app, text="Preview of Renamed Files:").grid(row=3, column=0, columnspan=2, padx=10, pady=10)
preview_listbox = tk.Listbox(app, width=60, height=10)
preview_listbox.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

# Progress Bar
progress_bar = ttk.Progressbar(app, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# Create Start buttons for preview and renaming
tk.Button(app, text="Preview", command=start_preview).grid(row=6, column=0, pady=20)
tk.Button(app, text="Start Renaming", command=start_renaming).grid(row=6, column=2, pady=20)

# Start the Tkinter main loop
app.mainloop()