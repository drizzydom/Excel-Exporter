import pandas as pd
import json
import os
import numpy as np
import glob

# Update this with your Excel filename
EXCEL_FILE = 'JadePMSchedule.xlsm'

PRIMARY_FIELDS = [
    'Machine Number',
    'Serial Number',
    'Machine',
    'Next PM Date',
    'Days Until Next PM',
    'Sheet Link',
    'Comments for Maintenance',
    'Comments for Parts'
]

VERTICAL_MACHINE_FIELDS = [
    'Days Until Next PM',
    'Last PM Done',
    'Recommended Date of Next PM',
    'Maintenance Type',
    'Maintenance Done',
    'Required Materials',
    'Qty.',
    'Frequency'
]


def find_header_row_dynamic(df, expected_fields, min_matches=3):
    for i, row in df.iterrows():
        matches = sum(any(str(cell).strip().lower() == ef.lower() for ef in expected_fields) for cell in row)
        if matches >= min_matches:
            return i
    return 0


def extract_primary_sheet(xls):
    primary_sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=primary_sheet, header=None)
    header_row = find_header_row_dynamic(df, PRIMARY_FIELDS, min_matches=3)
    df.columns = df.iloc[header_row]
    df = df.drop(range(header_row + 1))
    df = df.reset_index(drop=True)
    machines = []
    for _, row in df.iterrows():
        machine = {field: row.get(field, None) for field in PRIMARY_FIELDS}
        # Convert NaN to None for JSON
        for k, v in machine.items():
            if pd.isna(v):
                machine[k] = None
        machine_sheet = row.get('Sheet Link')
        if pd.notna(machine_sheet) and str(machine_sheet) in xls.sheet_names:
            machine['Sheet Name'] = str(machine_sheet)
        machines.append(machine)
    return machines


def extract_parts_and_maintenance(df):
    # Dynamically find the vertical field names (headers) in the first 20 rows and 5 columns
    header_candidates = {}
    for col in range(3, min(6, df.shape[1])):  # D is col 3
        for row in range(0, min(20, df.shape[0])):
            cell = str(df.iloc[row, col]).strip()
            if cell in VERTICAL_MACHINE_FIELDS:
                header_candidates[cell] = row
    field_names = [f for f in VERTICAL_MACHINE_FIELDS if f in header_candidates]
    parts = []
    max_col = df.shape[1]
    for col in range(4, max_col):
        # Check if the entire column for the relevant rows is empty
        is_empty = True
        for field in field_names:
            row_idx = header_candidates[field]
            if df.shape[0] > row_idx:
                value = df.iloc[row_idx, col]
                if not (pd.isna(value) or str(value).strip() == ''):
                    is_empty = False
                    break
        if is_empty:
            break  # Stop iteration at the first empty column
        # Find the part name from the 'Maintenance Done' field if available
        part_name = None
        if 'Maintenance Done' in header_candidates:
            row_idx = header_candidates['Maintenance Done']
            part_name = str(df.iloc[row_idx, col]).strip()
        if not part_name or part_name.lower() == 'nan':
            continue
        maintenance = {}
        for field in field_names:
            row_idx = header_candidates[field]
            value = df.iloc[row_idx, col] if df.shape[0] > row_idx else None
            if pd.isna(value):
                value = None
            maintenance[field] = value
        parts.append({
            'Part Name': part_name,
            'Maintenance': maintenance
        })
    return parts


def extract_vertical_machine_sheet(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    # Extract vertical fields as before
    data = {}
    for field in VERTICAL_MACHINE_FIELDS:
        found = False
        for i, row in df.iterrows():
            for j, cell in enumerate(row):
                if str(cell).strip() == field:
                    value = row[j+1] if j+1 < len(row) else None
                    if pd.isna(value):
                        value = None
                    data[field] = value
                    found = True
                    break
            if found:
                break
        if not found:
            data[field] = None
    # Extract parts and their maintenance
    data['Parts'] = extract_parts_and_maintenance(df)
    return data


def extract_machine_sheets(xls, machines):
    for machine in machines:
        sheet_name = machine.get('Sheet Name')
        if not sheet_name:
            continue
        machine['MaintenanceData'] = extract_vertical_machine_sheet(xls, sheet_name)
    return machines


def main():
    # Find all Excel files in the current directory (xls, xlsx, xlsm), skip temp files
    excel_files = [f for f in glob.glob('*.xls*') if not f.startswith('~$')]
    if not excel_files:
        print("No Excel files found in the current directory.")
        return
    for excel_file in excel_files:
        print(f"Processing {excel_file}...")
        try:
            xls = pd.ExcelFile(excel_file)
            machines = extract_primary_sheet(xls)
            machines = extract_machine_sheets(xls, machines)
            output_name = f"output_{os.path.splitext(excel_file)[0]}.json"
            with open(output_name, 'w') as f:
                json.dump(machines, f, indent=2, default=str)
            print(f"Extraction complete. Data saved to {output_name}.")
        except Exception as e:
            print(f"Failed to process {excel_file}: {e}")


if __name__ == "__main__":
    main()
