import pandas as pd

# Load the Excel file
file_path = r'C:\Users\Lenovo\Desktop\New folder (2)\Nokia script all olts.xlsx'
xls = pd.ExcelFile(file_path)

# Loop through each sheet and save it as a separate Excel file
for sheet_name in xls.sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    output_file = f'{sheet_name}.xlsx'  # Name the output file after the sheet
    df.to_excel(output_file, index=False)
    print(f'Saved {output_file}')
