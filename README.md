# Datamapping-excel
Data mapping from excel to excel
import pandas as pd
import openpyxl
import os
import re
from openpyxl.utils import column_index_from_string, get_column_letter

# === File Paths ===
data_path = r"C:\Users\sreer\OneDrive\Documents\Project\KPI\NHCL KPI Code and Data\NHCL-Database\NHCL-KPI-Q1-2025.xlsx"
asm_template = r"C:\Users\sreer\OneDrive\Documents\Project\KPI\NHCL KPI Code and Data\KPI Templates\ASMKPI-Template.xlsx"
me_template = r"C:\Users\sreer\OneDrive\Documents\Project\KPI\NHCL KPI Code and Data\KPI Templates\ME-KPI-Report-Template.xlsx"
rsm_zsm_template= r"C:\Users\sreer\OneDrive\Documents\Project\KPI\NHCL KPI Code and Data\KPI Templates\RSM_ZSM_KPI_Template.xlsx"
nsm_template= r"C:\Users\sreer\OneDrive\Documents\Project\KPI\NHCL KPI Code and Data\KPI Templates\NSM-KPI-Template.xlsx"
output_folder = r"C:\Users\sreer\OneDrive\Documents\Project\KPI\NHCL KPI Code and Data\Automated Reports\KPI-Reports-Q1-2025"
os.makedirs(output_folder, exist_ok=True)

# === Helper: Only write to top-left of merged cells ===
def is_top_left_of_merged_cell(ws, row, col):
    cell_coord = f"{get_column_letter(col)}{row}"
    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            return cell_coord == merged_range.coord.split(":")[0]
    return True

# === Load data ===
df = pd.read_excel(data_path, sheet_name="Master4code")  
df.columns = df.columns.str.strip()
df = df[df['HQ'].notna()].reset_index(drop=True)

# === ASM Cell Mapping ===
asm_mapping = {
"Name": "B2",
"Designation": "B3",
"Department": "B4",
"Email": "B5",
"HQ": "B6",
"HQ Type": "B7",
"FY": "E2",
"Quarter/ Month": "E3",
"Region": "E4",
"Area": "E5",
"HQ Count": "E6",
"Target Current Year (INR)": "F9",
"Overall Sales Achievement": "G9",
"Focus Brand Sales Target": "F10",
"Focus Brand Sales Achievement": "G10",
"Rising Brand Sales Target": "F11",
"Rising Brand Sales Achievement": "G11",
"Growth Target (%)":"F12",
"Growth(%)": "G12",
"Payment Collection": "G13",
"Call Average Self": "G14",
"Active Customer Addition": "G15",
"Personal Order Booking": "G16",
"Call Average Std": "G17", 
"HQ wise Sales": "G18", 
"HQ wise FB Sales": "G19",
"HQ wise RB Sales": "G20",
"HQ wise Growth": "G21"
}

# === ME/SE Cell Mapping ===
me_mapping = {
"Name": "B2", 
"Designation": "B3", 
"Department": "B4", 
"Email": "B5",
"HQ": "B6",
"FY": "E2",
"Quarter/ Month": "E3",
"Area": "E4",
"Region": "E5",
"HQ Type": "E6",
"Target Current Year (INR)": "F8",
"Overall Sales Achievement": "G8",
"Focus Brand Sales Target": "F9", 
"Focus Brand Sales Achievement": "G9",
"Rising Brand Sales Target": "F10", 
"Rising Brand Sales Achievement": "G10",
"Growth Target (%)":"F11",
"Growth(%)": "G11",
"Payment Collection": "G12",
"Active Customer Addition": "G13",
"Call Average Self": "G14",
"Samples Target":"F15",
"Sample Quantity Issued": "G15", 
"Personal Order Booking": "G16",
"Consistency": "G17", 
}

# === RSM/ZSM Cell Mapping ===
rsm_zsm_mapping = {
"Name": "B2", 
"Designation": "B3", 
"Department": "B4",
"Email": "B5",
"HQ Count": "B6",
"FY": "E2", 
"Quarter/ Month": "E3",
"Region": "E4",
"Count of ASM": "E5",
"HQ Type": "E6",
"Target Current Year (INR)": "F9",
"Overall Sales Achievement": "G9",
"Focus Brand Sales Target": "F10",
"Focus Brand Sales Achievement": "G10",
"Rising Brand Sales Target": "F11",
"Rising Brand Sales Achievement": "G11",
"Growth Target (%)":"F12",
"Growth(%)": "G12",
"Call Average Self": "G13",
"Active Customer Addition": "G14",
"Call Average Std": "G15", 
"HQ wise Sales": "G16", 
"HQ wise FB Sales": "G17",
"HQ wise RB Sales": "G18",
"HQ wise Growth": "G19",
"Brand Growth": "G20"
}

# === NSM Cell Mapping ===
nsm_mapping = {
"Name": "B2", 
"Designation": "B3", 
"Department": "B4", 
"Region": "B5",
"Email": "B6",
"FY": "E2", 
"Quarter/ Month": "E3",
"Count of ASM": "E4",
"HQ Count": "E5",
"Target Current Year (INR)": "G8",
"Overall Sales Achievement": "H8",
"Focus Brand Sales Target": "G9",
"Focus Brand Sales Achievement": "H9",
"Rising Brand Sales Target": "G10",
"Rising Brand Sales Achievement": "H10",
"Growth Target (%)":"G11",
"Growth(%)": "H11",
"Active Customer Addition": "H13",
"ASM TARGET":"H17",
"Brand Growth": "H18"
}


# === Report Generation ===
for _, row_data in df.iterrows():
    designation = str(row_data['Designation']).strip().upper()
    hq = str(row_data['HQ']).strip()

    if designation == "ASM":
        template_path = asm_template
        cell_mapping = asm_mapping
    elif designation in "ME":
        template_path = me_template
        cell_mapping = me_mapping
    elif designation in "RSM":
        template_path = rsm_zsm_template
        cell_mapping = rsm_zsm_mapping
    elif designation in "ZSM":
        template_path = rsm_zsm_template
        cell_mapping = rsm_zsm_mapping
    elif designation in "NSM":
        template_path = nsm_template
        cell_mapping = nsm_mapping
    else:
        print(f"Skipped: Unknown '{designation}' for HQ: {hq}")
        continue

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # === Fill all mapped fields ===
    for field, cell_address in cell_mapping.items():
        if field in row_data:
            col_letter = ''.join(filter(str.isalpha, cell_address))
            row_number = int(''.join(filter(str.isdigit, cell_address)))
            col_index = column_index_from_string(col_letter)
            if is_top_left_of_merged_cell(ws, row_number, col_index):
                ws.cell(row=row_number, column=col_index).value = row_data[field]

    # === Save the file ===
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", hq).replace(" ", "_")
    output_path = os.path.join(output_folder, f"{safe_name}_{designation}_KPI.xlsx")
    wb.save(output_path)
    print(f" Saved: {output_path}")

print("\n All KPI files generated correctly based on Designation.")
