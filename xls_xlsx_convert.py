import os
import glob
import pandas as pd

def convert_xls_to_xlsx(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xls'):
                xls_path = os.path.join(root, file)
                xlsx_path = os.path.splitext(xls_path)[0] + '.xlsx'
                
                # Read the .xls file
                df = pd.read_excel(xls_path, sheet_name=None)
                
                # Write to .xlsx file
                with pd.ExcelWriter(xlsx_path) as writer:
                    for sheet_name, data in df.items():
                        data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"Converted {xls_path} to {xlsx_path}")

if __name__ == "__main__":
    current_directory = os.path.dirname(os.path.abspath(__file__))
    convert_xls_to_xlsx(current_directory)