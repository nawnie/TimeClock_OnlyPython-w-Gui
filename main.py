import tkinter as tk
import subprocess
import os
from datetime import datetime

def launch_script(script_name):
    subprocess.Popen(['python', script_name])

def check_payroll_file():
    current_week = datetime.now().strftime("%Y-%U")
    payroll_file = f"Payroll_Week_{current_week}.xlsx"

    if not os.path.isfile(payroll_file):
        launch_script('create_timesheet.py')

def main():
    check_payroll_file()
    
    root = tk.Tk()
    root.title("Simple Tkinter Window")
    root.geometry("300x200")

    button1 = tk.Button(root, text="Clock_in-out", command=lambda: launch_script('timestamp.py'))
    button2 = tk.Button(root, text="Manage Employees", command=lambda: launch_script('Add_remove.py'))
    button3 = tk.Button(root, text="View Hours", command=lambda: launch_script('total_hours.py'))
    button4 = tk.Button(root, text="Export time sheet", command=lambda: launch_script('export.py'))

    button1.pack(pady=10)
    button2.pack(pady=10)
    button3.pack(pady=10)
    button4.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
