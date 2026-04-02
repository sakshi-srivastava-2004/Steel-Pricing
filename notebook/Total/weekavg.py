import pandas as pd
import numpy as np

input_file = 'MELTING SCRAP_HMS(80-20).xlsx'   # your input Excel file
output_file = 'MELTING SCRAP_HMS(80-20)_averaged_data.xlsx'  # output file

def average_weeks_only(df):
    df = df.replace(['-', '', ' '], np.nan).reset_index(drop=True)

    averaged_weeks = []
    start_idx = 0
    week_number = 1


    first_col = df.iloc[:, 0].astype(str).str.strip().str.upper()

    for i, val in enumerate(first_col):
        if val == 'H':
            week_df = df.iloc[start_idx:i]
            if not week_df.empty:
                numeric_avg = week_df.select_dtypes(include=[np.number]).mean()

                
                non_numeric_avg = week_df.select_dtypes(exclude=[np.number]).apply(lambda col: col.dropna().iloc[0] if not col.dropna().empty else np.nan)

                avg_row = pd.concat([non_numeric_avg, numeric_avg])

              
                avg_row = pd.Series([week_number] + avg_row.tolist(), index=['Week_Number'] + df.columns.tolist())

                averaged_weeks.append(avg_row)

                week_number += 1
            start_idx = i + 1

    week_df = df.iloc[start_idx:]
    if not week_df.empty:
        numeric_avg = week_df.select_dtypes(include=[np.number]).mean()
        non_numeric_avg = week_df.select_dtypes(exclude=[np.number]).apply(lambda col: col.dropna().iloc[0] if not col.dropna().empty else np.nan)
        avg_row = pd.concat([non_numeric_avg, numeric_avg])
        avg_row = pd.Series([week_number] + avg_row.tolist(), index=['Week_Number'] + df.columns.tolist())
        averaged_weeks.append(avg_row)


    result_df = pd.DataFrame(averaged_weeks)
    return result_df


def main():
    xl = pd.ExcelFile(input_file)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for sheet_name in xl.sheet_names:
        print(f"Processing sheet: {sheet_name}")
        df = xl.parse(sheet_name)

        if df.columns.duplicated().any():
            df.columns = pd.io.parsers.ParserBase({'names': df.columns})._maybe_dedup_names(df.columns)
            print(f"Duplicate columns detected and renamed in sheet: {sheet_name}")

        averaged_df = average_weeks_only(df)
        averaged_df.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.close()
    print(f"\n Weekly averages saved to '{output_file}'")


if __name__ == '__main__':
    main()
