import pandas as pd
import os


def convert_excel_to_csv():
    """
    Convert Excel file with multiple sheets to individual CSV files.
    Uses the 4th row as headers and replaces missing values with '*'.
    """
    excel_file = "Menues/menueplan.xlsx"
    output_folder = "csv_files"
    
    # Create output directory
    os.makedirs(output_folder, exist_ok=True)
    
    # Read all sheets from Excel file
    df_dict = pd.read_excel(excel_file, sheet_name=None, header=None)
    
    for sheet_name, df in df_dict.items():
        # Process each sheet
        processed_df = _process_sheet(df)
        
        # Save to CSV
        csv_file_path = os.path.join(output_folder, f"{sheet_name}.csv")
        processed_df.to_csv(csv_file_path, index=False)


def _process_sheet(df):
    """
    Process a single Excel sheet by setting headers and cleaning data.
    
    Args:
        df (pd.DataFrame): Raw DataFrame from Excel sheet
        
    Returns:
        pd.DataFrame: Processed DataFrame with cleaned headers and data
    """
    # Use 4th row (index 3) as header
    header_row = df.iloc[3].tolist()
    
    # Clean headers: replace NaN or blank with '*'
    cleaned_columns = _clean_column_names(header_row)
    
    # Skip the first 4 rows, use the rest as data
    new_df = df.iloc[5:].reset_index(drop=True)
    new_df.columns = cleaned_columns
    
    # Replace NaNs and blank strings with '*'
    new_df = new_df.fillna('*')
    
    return new_df


def _clean_column_names(column_names):
    """
    Clean column names by replacing NaN or empty values with '*'.
    
    Args:
        column_names (list): List of column names from Excel
        
    Returns:
        list: Cleaned column names
    """
    cleaned_columns = []
    
    for col in column_names:
        if pd.isna(col) or str(col).strip() == '':
            cleaned_columns.append('*')
        else:
            cleaned_columns.append(str(col))
    
    return cleaned_columns