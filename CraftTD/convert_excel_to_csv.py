import pandas as pd
import sys
import os

def convert_excel_to_csv(excel_file=None, csv_file=None, separator=None, sheet_name=None):
    """
    Convert Excel localization file to CSV format
    
    Args:
        excel_file: Path to Excel file (default: Localization.xlsx)
        csv_file: Path to output CSV file (default: same name as Excel with .csv extension)
        separator: CSV separator character (default: comma)
        sheet_name: Name of the sheet to convert (default: all sheets combined)
    """
    # Default parameters
    default_excel_file = 'Localization.xlsx'
    default_csv_file = "Localization.csv"
    default_separator = '|'
    
    # Set parameters with defaults
    excel_file = excel_file if excel_file is not None else default_excel_file
    csv_file = csv_file if csv_file is not None else default_csv_file
    separator = separator if separator is not None else default_separator
    
    if csv_file is None:
        csv_file = os.path.splitext(excel_file)[0] + '.csv'
    
    try:
        # Check if Excel file exists
        if not os.path.exists(excel_file):
            print(f"Error: File '{excel_file}' not found")
            return None
        
        # Read Excel file
        if sheet_name:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            print(f"Loaded sheet '{sheet_name}' from {excel_file} with {len(df)} entries")
        else:
            # Load all sheets and combine them
            wb = pd.ExcelFile(excel_file)
            print(f"Found sheets: {wb.sheet_names}")s
            
            dfs = []
            for sheet in wb.sheet_names:
                df_sheet = pd.read_excel(excel_file, sheet_name=sheet)
                print(f"Loaded sheet '{sheet}' with {len(df_sheet)} entries")
                dfs.append(df_sheet)
            
            # Combine all sheets
            df = pd.concat(dfs, ignore_index=True)
            print(f"Combined total: {len(df)} entries")
        
        # Remove rows where all values are NaN (empty rows)
        df = df.dropna(how='all')
        print(f"After removing empty rows: {len(df)} entries")
        
        # Remove rows where the key column is empty or NaN
        if 'key' in df.columns:
            df = df[df['key'].notna() & (df['key'] != '')]
            print(f"After removing rows with empty keys: {len(df)} entries")
        
        print(f"Columns: {df.columns.tolist()}")
        
        # Save to CSV with custom separator
        df.to_csv(csv_file, index=False, encoding='utf-8', sep=separator)
        print(f"Successfully converted to {csv_file} with separator '{separator}'")
        
        # Show preview
        print("\nPreview of CSV file:")
        print(df.head().to_string())
        
        return df
        
    except Exception as e:
        print(f"Error: {e}")
        return None

if __name__ == "__main__":
    # Command line usage
    if len(sys.argv) == 1:
        # Default: convert all sheets from Localization.xlsx to single CSV
        convert_excel_to_csv()
    elif len(sys.argv) == 2:
        # Convert specific file (all sheets to single CSV)
        convert_excel_to_csv(sys.argv[1])
    elif len(sys.argv) == 3:
        # Convert specific file with custom output name (all sheets)
        convert_excel_to_csv(sys.argv[1], sys.argv[2])
    elif len(sys.argv) == 4:
        # Convert specific file with custom output name and separator (all sheets)
        convert_excel_to_csv(sys.argv[1], sys.argv[2], sys.argv[3])
    elif len(sys.argv) == 5:
        # Convert specific file with custom output name, separator, and sheet name (single sheet)
        convert_excel_to_csv(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        print("Usage:")
        print("  python convert_excel_to_csv.py")
        print("  python convert_excel_to_csv.py <excel_file>")
        print("  python convert_excel_to_csv.py <excel_file> <csv_file>")
        print("  python convert_excel_to_csv.py <excel_file> <csv_file> <separator>")
        print("  python convert_excel_to_csv.py <excel_file> <csv_file> <separator> <sheet_name>")
        print("\nExamples:")
        print("  python convert_excel_to_csv.py")
        print("  python convert_excel_to_csv.py Localization.xlsx")
        print("  python convert_excel_to_csv.py Localization.xlsx custom_output.csv")
        print("  python convert_excel_to_csv.py Localization.xlsx custom_output.csv ';'")
        print("  python convert_excel_to_csv.py Localization.xlsx custom_output.csv '|' spell")