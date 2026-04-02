import pandas as pd
import os

# List your Excel files here
excel_files = [
    "12mmweekly_averaged_data.xlsx",
    "PRIMARY_extracted.xlsx",
    "billet_averaged.xlsx",
    "ingot_averaged.xlsx",
    "CRC_extracted.xlsx",
    "pig_iron_averaged_data.xlsx",
    "HRC_extracted.xlsx",
    "HR PLATE_extracted.xlsx",
    "IMPORT_INDIA_extracted_averaged_data.xlsx",
    "MELTING SCRAP_HMS(80-20)_averaged_data.xlsx",
    "MELTING SCRAP_End Cutting_averaged_data.xlsx",
    "MELTING SCRAP_CR Busheling (Loose)_averaged_data.xlsx"
]

folder_path = "."

combined_data = []

for file_name in excel_files:
    file_path = os.path.join(folder_path, file_name)
    try:
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Add filename and sheet name columns
            df.insert(0, "File Name", file_name)
            df.insert(1, "Sheet Name", sheet_name)

            combined_data.append(df)
    except Exception as e:
        print(f"Error processing file {file_name}: {e}")

final_df = pd.concat(combined_data, ignore_index=True)

# Add Serial Numbers
final_df.insert(0, "S.No", range(1, len(final_df) + 1))

output_path = os.path.join(folder_path, "final_output.xlsx")
final_df.to_excel(output_path, index=False)

print(f"Successfully saved combined data to {output_path}")
