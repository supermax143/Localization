import pandas as pd

def add_shop_select_localization():
    """Add 'shop_select' localization to BrianSkipper.xlsx"""
    
    # Read existing file
    df = pd.read_excel('BrianSkipper.xlsx')
    
    # New shop_select localization entry
    new_entry = {"key": "shop_select", "ru": "Выбрать", "en": "Select"}
    
    # Create DataFrame for new entry
    new_df = pd.DataFrame([new_entry])
    
    # Concatenate existing and new entries
    updated_df = pd.concat([df, new_df], ignore_index=True)
    
    # Sort by key for better organization
    updated_df = updated_df.sort_values('key').reset_index(drop=True)
    
    # Save to Excel file
    updated_df.to_excel('BrianSkipper.xlsx', index=False)
    
    print(f"Added 'shop_select' localization entry to BrianSkipper.xlsx")
    print(f"Total entries: {len(updated_df)}")
    print("\nUpdated entries:")
    print(updated_df.to_string())
    
    return updated_df

if __name__ == "__main__":
    add_shop_select_localization()
