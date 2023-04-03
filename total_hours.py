import tkinter as tk
from tkinter import simpledialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import os

def nearest_sunday(date):
    return date - ((timedelta(days=date.weekday() + 1)) % timedelta(days=7))
    
def get_current_payroll_filename():
    start_date = nearest_sunday(datetime.now().date())
    current_week = start_date.strftime("%Y-%U")
    filename = f"Payroll_Week_{current_week}.xlsx"
    return filename

def calculate_total_hours(PIN):
    filename = get_current_payroll_filename()

    if not os.path.isfile(filename):
        print("Payroll file not found.")
        return

    df = pd.read_excel(filename)
    employee_row = df.loc[df['PIN'] == int(PIN)]  # Cast the entered PIN to an integer

    if employee_row.empty:
        print("PIN not found.")
        return

    index = employee_row.index[0]
    dates = employee_row.columns[3:]
    total_hours = timedelta()

    for date in dates:
        timestamps = employee_row[date].values[0]
        if pd.isnull(timestamps):
            continue

        timestamps_list = timestamps.split(", ")
        for i in range(0, len(timestamps_list) - 1, 2):
            time_in_str = timestamps_list[i].split(" ")[1]
            time_out_str = timestamps_list[i + 1].split(" ")[1]
            time_in = datetime.strptime(time_in_str, "%H:%M:%S")
            time_out = datetime.strptime(time_out_str, "%H:%M:%S")
            total_hours += time_out - time_in

    return total_hours

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    PIN = simpledialog.askstring("Enter PIN", "Please enter your PIN:")

    if PIN is not None:
        total_hours = calculate_total_hours(PIN)
        if total_hours is not None:
            messagebox.showinfo("Total Hours", f"Total hours worked: {total_hours}")
        else:
            messagebox.showerror("Error", "An error occurred while calculating total hours.")
    else:
        print("No PIN entered.")

    root.destroy()

if __name__ == "__main__":
    main()
