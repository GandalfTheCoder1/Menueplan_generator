import pandas as pd
import os

def convert_excel_to_csv():
    excel_file = "Menues/menueplan.xlsx"
    output_folder = "csv_files"
    os.makedirs(output_folder, exist_ok=True)

    df_dict = pd.read_excel(excel_file, sheet_name=None, header=None)

    for sheet_name, df in df_dict.items():
        # Use 4th row (index 3) as header
        new_columns = df.iloc[3].tolist()

        # Clean header: replace NaN or blank headers
        cleaned_columns = []
        for idx, col in enumerate(new_columns):
            if pd.isna(col) or str(col).strip() == '':
                cleaned_columns.append(f"*")
            else:
                cleaned_columns.append(str(col))

        # Skip the first 4 rows, use the rest as data
        new_df = df.iloc[5:].reset_index(drop=True)
        new_df.columns = cleaned_columns

        # Replace NaNs and blank strings with '*'
        new_df = new_df.fillna('*')

        # Save to CSV
        csv_file_path = os.path.join(output_folder, f"{sheet_name}.csv")
        new_df.to_csv(csv_file_path, index=False)

    return
