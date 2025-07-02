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


def extract_maintenance_history(df):
    """Extract maintenance history table and return historical maintenance records mapped to parts"""
    historical_records = {}
    
    # Find the maintenance history table by looking for "Date" header around row 16-20
    date_header_row = None
    for row in range(15, min(25, df.shape[0])):
        for col in range(0, min(5, df.shape[1])):
            cell = str(df.iloc[row, col]).strip().lower()
            if cell == 'date':
                date_header_row = row
                break
        if date_header_row is not None:
            break
    
    if date_header_row is None:
        return historical_records
    
    # Get the part names and their maintenance types from the maintenance data header rows
    part_names = {}
    part_maintenance_types = {}
    for col in range(4, df.shape[1]):  # Start from column E (index 4)
        for row in range(5, min(15, df.shape[0])):  # Look in the maintenance section
            cell = str(df.iloc[row, col]).strip()
            field_name = str(df.iloc[row, 3]).strip().lower()
            
            if cell and cell.lower() != 'nan':
                if 'maintenance done' in field_name:
                    part_names[col] = cell
                elif 'maintenance type' in field_name:
                    part_maintenance_types[col] = cell
    
    # Extract data rows until we hit an empty row
    for row in range(date_header_row + 1, df.shape[0]):
        # Check if the row is entirely empty
        row_data = []
        is_empty_row = True
        for col in range(df.shape[1]):
            cell_value = df.iloc[row, col] if row < df.shape[0] and col < df.shape[1] else None
            if not pd.isna(cell_value) and str(cell_value).strip():
                is_empty_row = False
            row_data.append(cell_value if not pd.isna(cell_value) else None)
        
        if is_empty_row:
            break
        
        # Extract basic maintenance info
        date_val = row_data[0] if len(row_data) > 0 else None
        technician_val = row_data[1] if len(row_data) > 1 else None
        work_order_val = row_data[2] if len(row_data) > 2 else None
        po_number_val = row_data[3] if len(row_data) > 3 else None
        
        # Check each part column for completion
        for col_idx, part_name in part_names.items():
            if col_idx < len(row_data) and row_data[col_idx]:
                completion_status = str(row_data[col_idx]).strip().lower()
                if completion_status in ['yes', 'completed', 'y', 'true']:
                    # Get the actual maintenance type for this part
                    maintenance_type = part_maintenance_types.get(col_idx, 'Maintenance')
                    
                    # Create historical record for this part
                    historical_record = {
                        'Last PM Done': date_val,
                        'Technician': technician_val,
                        'Work Order': work_order_val,
                        'Po Number': po_number_val,
                        'Maintenance Type': maintenance_type,
                        'Completion Status': completion_status
                    }
                    
                    # Add to the part's historical records
                    if part_name not in historical_records:
                        historical_records[part_name] = []
                    historical_records[part_name].append(historical_record)
    
    return historical_records


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
    current_parts = extract_parts_and_maintenance(df)
    
    # Extract historical maintenance records mapped to parts
    historical_records = extract_maintenance_history(df)
    
    # Add historical records to existing parts
    for part in current_parts:
        part_name = part['Part Name']
        if part_name in historical_records:
            # Add historical maintenance records to this part
            part['Historical_Maintenance'] = historical_records[part_name]
    
    data['Parts'] = current_parts
    
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
