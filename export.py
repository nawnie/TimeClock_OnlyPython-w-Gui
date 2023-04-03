import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime, timedelta
import os
import shutil



def nearest_sunday(date):
    return date - ((timedelta(days=date.weekday() + 1)) % timedelta(days=7))

def get_current_payroll_filename():
    start_date = nearest_sunday(datetime.now().date())
    current_week = start_date.strftime("%Y-%U")
    filename = f"Payroll_Week_{current_week}.xlsx"
    return filename

def copy_to_desktop(filename):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = os.path.basename(filename)  # Get the file name without the directory path
    destination = os.path.join(desktop, filename)
    shutil.copy2(filename, destination)

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    filename = get_current_payroll_filename()  # Get the full file path

    # Check if the file exists in the working directory (current script directory)
    if not os.path.isfile(filename):
        messagebox.showerror("Error", "Payroll file not found.")
    else:
        try:
            # Copy the file from the working directory to the desktop
            copy_to_desktop(filename)  # Pass the full file path to the copy_to_desktop function
            messagebox.showinfo("Success", "Copied to desktop complete.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while copying the file: {e}")

    root.destroy()

if __name__ == "__main__":
    main()
