import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font

# Create data for the Savings sections
savings_data = [
    {"Goal": 3000, "Deadline": "2023-10-28", "Amount Saved": 424.24, "Days Left": 99, "Percentage Saved": 14,
     "Need to Save Per Day": 30.30, "Need to Save Per Week": 212.12, "Need to Save Per Month": 909.09},
    {"Goal": 2500, "Deadline": "2023-12-01", "Amount Saved": 0, "Days Left": 72, "Percentage Saved": 0,
     "Need to Save Per Day": 34.72, "Need to Save Per Week": 243.06, "Need to Save Per Month": 1041.67},
    {"Goal": 20000, "Deadline": "2024-09-01", "Amount Saved": 0, "Days Left": 408, "Percentage Saved": 0,
     "Need to Save Per Day": 49.02, "Need to Save Per Week": 343.16, "Need to Save Per Month": 1471.57},
    {"Goal": 1500, "Deadline": "2023-09-23", "Amount Saved": 0, "Days Left": 64, "Percentage Saved": 0,
     "Need to Save Per Day": 23.43, "Need to Save Per Week": 164.00, "Need to Save Per Month": 703.00},
]

# Create a DataFrame from the savings data
df = pd.DataFrame(savings_data)

# Add columns for Plan & Track (Assuming tracking plan data)
plan_data = {
    "Plan & Track Date": [
        "2023-07-20", "2023-07-27", "2023-08-03", "2023-08-10", "2023-08-17", "2023-08-24",
        "2023-08-31", "2023-09-07", "2023-09-14", "2023-09-21", "2023-09-28", "2023-10-05"
    ],
    "Amount to Save": [
        212.12, 212.12, 212.12, 212.12, 212.12, 212.12, 212.12, 212.12, 212.12, 212.12, 212.12, 212.12
    ]
}

# Convert the plan data to a DataFrame
plan_df = pd.DataFrame(plan_data)

# Create Excel writer
file_name = 'savings_plan.xlsx'
with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Savings Plan', index=False, startrow=1)
    plan_df.to_excel(writer, sheet_name='Savings Plan', index=False, startcol=12, startrow=1)

    # Load the workbook and select the sheet
    workbook = writer.book
    worksheet = workbook['Savings Plan']

    # Apply some styling
    green_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    font = Font(bold=True)

    # Fill header rows with color and apply bold font
    for cell in worksheet["2:2"]:
        cell.fill = green_fill
        cell.font = font

# Notify completion
print(f"{file_name} created successfully!")
