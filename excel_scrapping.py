import pandas as pd
import os

original_file_path = 'xls/Chelsea Ng_Java-B.xlsx'

output_directory = 'report card'

if not os.path.exists(output_directory):
    os.makedirs(output_directory)

excel_file = pd.ExcelFile(original_file_path)

for sheet_name in excel_file.sheet_names:
    df = pd.read_excel(original_file_path, sheet_name=sheet_name)
    new_file_path = os.path.join(output_directory, f'Chelsea Ng_Java-B {sheet_name}.xlsx')
    df.to_excel(new_file_path, index=False)

print("Success")