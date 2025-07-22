# Fleet Utilization Report

This Python script reads fleet data from an Excel file (`fleet_data.xlsx`) and creates a utilization report showing the percentage of days each vehicle was rented.

## What it does

- Loads data from `fleet_data.xlsx`  
- Skips rows with missing data  
- Calculates utilization percentage (`days rented / total days * 100`)  
- Writes a new Excel file `utilization_report.xlsx` with Vehicle ID, Branch, and Utilization %

## How to use

1. Place your fleet data in `fleet_data.xlsx` with columns: Vehicle ID, Branch, Total Days, Days Rented (starting from row 2)  
2. Run the script with Python 3 installed:

```bash
python fleet.py
