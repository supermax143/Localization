import pandas as pd
import sys
import os

def convert_excel_to_csv(excel_file=None, csv_file=None, separator=None):
    """
    Convert Excel localization file to CSV format
    
    Args:
        excel_file: Path to Excel file (default: Localization.xlsx)
        csv_file: Path to output CSV file (default: same name as Excel with .csv extension)
        separator: CSV separator character (default: comma)
    """
    # Default parameters
    default_excel_file = 'BriansKeeper.xlsx'
    default_csv_file = "Localization.csv"  # Will be generated from excel_file if None
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
        df = pd.read_excel(excel_file)
        print(f"Loaded {excel_file} with {len(df)} entries")
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
        # Default: convert Localization.xlsx
        convert_excel_to_csv()
        
        # Also convert BrianSkipper.xlsx if it exists
        if os.path.exists('BrianSkipper.xlsx'):
            print("\n" + "="*50)
            convert_excel_to_csv('BrianSkipper.xlsx')
    elif len(sys.argv) == 2:
        # Convert specific file
        convert_excel_to_csv(sys.argv[1])
    elif len(sys.argv) == 3:
        # Convert specific file with custom output name
        convert_excel_to_csv(sys.argv[1], sys.argv[2])
    elif len(sys.argv) == 4:
        # Convert specific file with custom output name and separator
        convert_excel_to_csv(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        print("Usage:")
        print("  python convert_excel_to_csv.py")
        print("  python convert_excel_to_csv.py <excel_file>")
        print("  python convert_excel_to_csv.py <excel_file> <csv_file>")
        print("  python convert_excel_to_csv.py <excel_file> <csv_file> <separator>")
        print("\nExamples:")
        print("  python convert_excel_to_csv.py BrianSkipper.xlsx")
        print("  python convert_excel_to_csv.py BrianSkipper.csv BrianSkipper.csv")
        print("  python convert_excel_to_csv.py BrianSkipper.csv BrianSkipper.csv ';'")
        print("  python convert_excel_to_csv.py Localization.xlsx custom_output.csv '|'")
