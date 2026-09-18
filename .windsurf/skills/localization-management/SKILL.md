# Localization Management Skill

## Overview
This skill provides comprehensive guidance for managing game localization files, including Excel file structure, CSV conversion, and data management. It covers the complete workflow from adding new translations to converting localization data for game integration.

## ⚠️ CRITICAL RULE

**NEVER edit the CSV file directly!** 

When adding or modifying localizations:
- **ONLY edit `Localization.xlsx`** - this is the source of truth
- **CSV files are automatically generated** by running the conversion script
- **Manual CSV edits will be overwritten** on next conversion
- **Always run the conversion script** after Excel changes

**Workflow:**
1. Make changes in `Localization.xlsx`
2. Run `convert_excel_to_csv.py` or `convert_to_csv.bat`
3. Use the generated CSV file in your project

## File Structure

### Excel File Organization
- **Primary file**: `Localization.xlsx`
- **Multi-sheet structure**: Each game category has its own sheet
- **Standard columns**: `key`, `ru`, `en` (expandable for additional languages)
- **Sheet naming convention**: Use descriptive names (Sheet1, spell, Shop, etc.)

### Common Sheet Types
- **Sheet1**: General UI elements, common phrases
- **spell**: Spell names and descriptions
- **Shop**: Shop items, prices, descriptions
- **Units**: Unit names, abilities, descriptions
- **Menus**: Menu items, buttons, labels

## Scripts and Tools

### Conversion Script: `convert_excel_to_csv.py`

**Purpose**: Convert Excel localization files to CSV format for game integration

**Usage**:
```bash
# Convert all sheets to single CSV file
python convert_excel_to_csv.py

# Convert specific Excel file
python convert_excel_to_csv.py Localization.xlsx

# Convert with custom output name
python convert_excel_to_csv.py Localization.xlsx custom_output.csv

# Convert with custom separator
python convert_excel_to_csv.py Localization.xlsx custom_output.csv ';'

# Convert specific sheet only
python convert_excel_to_csv.py Localization.xlsx output.csv '|' spell
```

**Features**:
- Automatically combines all Excel sheets into single CSV file
- Filters empty rows to avoid `||` artifacts
- Custom separator support (default: `|`)
- UTF-8 encoding for proper character support
- Removes rows with empty keys

### Batch Script: `convert_to_csv.bat`

**Purpose**: Quick conversion via double-click

**Content**:
```batch
@echo off
echo Converting Excel localization files to CSV format...
echo.

python convert_excel_to_csv.py

echo.
echo Conversion complete!
```

## Adding New Localizations

### Method 1: Direct Excel Editing
1. Open `Localization.xlsx`
2. Navigate to appropriate sheet (or create new one)
3. Add new row with format: `key | ru | en`
4. Save file
5. Run conversion script

### Method 2: Python Script Addition
```python
import openpyxl

wb = openpyxl.load_workbook('Localization.xlsx')

# Create or get sheet
if 'YourSheet' not in wb.sheetnames:
    ws = wb.create_sheet('YourSheet')
else:
    ws = wb['YourSheet']

# Add headers if needed
ws.cell(row=1, column=1, value='key')
ws.cell(row=1, column=2, value='ru')
ws.cell(row=1, column=3, value='en')

# Add new localization
ws.cell(row=2, column=1, value='your_key')
ws.cell(row=2, column=2, value='Ваш текст')
ws.cell(row=2, column=3, value='Your text')

wb.save('Localization.xlsx')
```



## Key Naming Conventions

### Patterns by Category
- **Common**: `common_no`, `common_hello`, `common_yes`
- **Spells**: `spell_name_fireball`, `spell_description_fireball`
- **Shop**: `shopitem_name_crystal_pack_1`, `shopitem_description_crystal_pack_1`
- **UI**: `ui_button_start`, `ui_label_score`, `ui_menu_settings`

### Best Practices
- Use underscores instead of spaces
- Include category prefix
- Use descriptive names
- Maintain consistency within categories
- Avoid special characters

## Troubleshooting

### Empty Rows in CSV
**Problem**: CSV file contains many `||` lines

**Solution**: The conversion script now automatically filters empty rows. If problem persists:
1. Check Excel file for hidden rows
2. Ensure no formatting in empty cells
3. Recreate Excel file with clean structure

### Character Encoding Issues
**Problem**: Russian characters display as `���`

**Solution**: 
1. Ensure UTF-8 encoding in conversion script
2. Check Excel file encoding
3. Verify Python script uses `encoding='utf-8'`

### Sheet Structure Issues
**Problem**: Excel file has 1000+ rows with mostly empty data

**Solution**: Recreate Excel file with clean structure:
```python
import openpyxl
wb = openpyxl.Workbook()
ws = wb.create_sheet('YourSheet')
# Add only necessary data
wb.save('Localization.xlsx')
```

### Conversion Script Issues
**Problem**: Script not combining all sheets into single CSV

**Solution**: Ensure script uses the updated version with sheet concatenation:
```python
# Load all sheets and combine them
wb = pd.ExcelFile(excel_file)
dfs = []
for sheet in wb.sheet_names:
    df_sheet = pd.read_excel(excel_file, sheet_name=sheet)
    dfs.append(df_sheet)

# Combine all sheets
df = pd.concat(dfs, ignore_index=True)
```



## Batch Operations

### Adding Multiple Localizations
```python
import openpyxl

wb = openpyxl.load_workbook('Localization.xlsx')
ws = wb['YourSheet']

items = [
    ('key1', 'Русский текст 1', 'English text 1'),
    ('key2', 'Русский текст 2', 'English text 2'),
    ('key3', 'Русский текст 3', 'English text 3'),
]

row = ws.max_row + 1
for key, ru, en in items:
    ws.cell(row=row, column=1, value=key)
    ws.cell(row=row, column=2, value=ru)
    ws.cell(row=row, column=3, value=en)
    row += 1

wb.save('Localization.xlsx')
```

### Creating New Sheet with Data
```python
import openpyxl

wb = openpyxl.load_workbook('Localization.xlsx')
ws = wb.create_sheet('NewCategory')

# Add headers
ws.cell(row=1, column=1, value='key')
ws.cell(row=1, column=2, value='ru')
ws.cell(row=1, column=3, value='en')

# Add data
# ... your data ...

wb.save('Localization.xlsx')
```

## Best Practices

1. **File Organization**
   - Keep Excel file in version control
   - Maintain backup before major changes
   - Use descriptive sheet names

2. **Data Quality**
   - Avoid empty rows in Excel
   - Use consistent key naming
   - Keep translations concise and accurate

3. **Workflow**
   - Edit in Excel first
   - Convert to CSV after changes
   - Verify CSV output before deployment
   - Keep CSV in version control

4. **Performance**
   - Convert only when needed
   - Keep file sizes manageable
   - Avoid excessive empty rows

## Project Structure Example

```
ExternalGameData/Localization/
├── Localization.xlsx          # Master Excel file
├── Localization.csv           # Generated CSV for Unity
├── convert_excel_to_csv.py    # Conversion script
├── convert_to_csv.bat         # Quick conversion
└── CraftTD/                   # Project-specific directory
    ├── Localization.xlsx      # Project-specific Excel
    ├── Localization.csv       # Project-specific CSV
    ├── convert_excel_to_csv.py # Project-specific script
    └── convert_to_csv.bat      # Project-specific batch
```

## Maintenance

### Regular Tasks
1. Update translations after content changes
2. Add new keys for new features
3. Review and improve existing translations
4. Backup Excel files regularly

### Version Control
- Commit Excel files with meaningful messages
- Include CSV files in repository
- Document major localization changes
- Use branches for experimental translations

## Advanced Features

### Adding New Languages
1. Add new column to Excel file (e.g., `de`, `fr`, `es`)
2. Update conversion script to handle additional columns
3. Verify CSV output includes new language column
4. Test that translations are properly formatted

### Localization Validation
```python
import pandas as pd

def validate_localization(file_path):
    df = pd.read_csv(file_path, sep='|')
    
    # Check for missing keys
    missing_keys = df[df['key'].isna() | (df['key'] == '')]
    if not missing_keys.empty:
        print("Warning: Rows with missing keys found")
    
    # Check for missing translations
    missing_ru = df[df['ru'].isna() | (df['ru'] == '')]
    missing_en = df[df['en'].isna() | (df['en'] == '')]
    
    print(f"Missing Russian translations: {len(missing_ru)}")
    print(f"Missing English translations: {len(missing_en)}")
    
    # Check for duplicate keys
    duplicates = df[df.duplicated(['key'], keep=False)]
    if not duplicates.empty:
        print("Warning: Duplicate keys found:")
        print(duplicates['key'].unique())
```

## Troubleshooting Commands

### Check Excel File Structure
```bash
python -c "import openpyxl; wb = openpyxl.load_workbook('Localization.xlsx'); print('Sheets:', wb.sheetnames); [print(f'{sheet}: {wb[sheet].max_row} rows') for sheet in wb.sheetnames]"
```

### Verify CSV Content
```bash
python -c "import pandas as pd; df = pd.read_csv('Localization.csv', sep='|'); print(f'Total: {len(df)} entries'); print(df.head(10).to_string())"
```

### Quick Conversion Check
```bash
cd path/to/localization && python convert_excel_to_csv.py
```

## Support and Resources

- **Excel Format**: Standard .xlsx with UTF-8 encoding
- **CSV Format**: Pipe-separated (`|`) with UTF-8 encoding
- **Python Requirements**: pandas, openpyxl