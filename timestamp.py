import tkinter as tk
from tkinter import simpledialog
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

def write_timestamp(PIN):
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
    today = str(datetime.now().date())
    current_time = datetime.now().strftime("%H:%M:%S")
    existing_timestamps = employee_row[today].values[0]

    if pd.isnull(existing_timestamps):
        updated_timestamps = f"timein {current_time}"
    else:
        timestamps = existing_timestamps.split(", ")
        last_entry = timestamps[-1]

        if "timein" in last_entry:
            updated_timestamps = f"{existing_timestamps}-timeout {current_time}"
        else:
            updated_timestamps = f"{existing_timestamps}, timein {current_time}"

    df.at[index, today] = updated_timestamps
    df.to_excel(filename, index=False)
    print(f"Timestamp updated: {current_time}")


def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    PIN = simpledialog.askstring("Enter PIN", "Please enter your PIN:")

    if PIN is not None:
        write_timestamp(PIN)
    else:
        print("No PIN entered.")

    root.destroy()

if __name__ == "__main__":
    main()
