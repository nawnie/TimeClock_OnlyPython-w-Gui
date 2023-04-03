import os
import pandas as pd
from datetime import datetime, timedelta

def read_employees(file_path):
    employees = []

    with open(file_path, 'r') as f:
        for line in f:
            first_name, last_name, pin = line.strip().split()
            employees.append({'First Name': first_name, 'Last Name': last_name, 'PIN': pin})

    return employees

def nearest_sunday(date):
    return date - ((timedelta(days=date.weekday() + 1)) % timedelta(days=7))

def create_timesheet(employees):
    start_date = nearest_sunday(datetime.now().date())
    current_week = start_date.strftime("%Y-%U")
    filename = f"Payroll_Week_{current_week}.xlsx"

    # Create a 14-day date range
    end_date = start_date + timedelta(days=13)
    dates = pd.date_range(start_date, end_date)

    # Create an empty DataFrame with columns for employee information and dates
    columns = ['First Name', 'Last Name', 'PIN'] + [str(date.date()) for date in dates]
    df = pd.DataFrame(columns=columns)

    # Fill in the employee information
    for employee in employees:
        df = df.append(employee, ignore_index=True)

    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)

def main():
    employees = read_employees("employees.txt")
    create_timesheet(employees)

if __name__ == "__main__":
    main()
