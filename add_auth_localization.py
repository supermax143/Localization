"""Add Authentication localization to BriansKeeper.xlsx"""
import pandas as pd

# Read existing file
df = pd.read_excel('BriansKeeper.xlsx')

# New Authentication localization entries
new_entries = [
    {"key": "auth_upgrade_weapon_message", "ru": "Ща мы их накажем! Давай улучшим оружие", "en": "Now we'll punish them! Let's upgrade the weapon"},
    {"key": "auth_tap_anywhere_to_continue", "ru": "Нажми где угодно, чтобы продолжить", "en": "Tap anywhere to continue"}
]

# Create DataFrame for new entries
new_df = pd.DataFrame(new_entries)

# Append to existing data
df = pd.concat([df, new_df], ignore_index=True)

# Save back to Excel
df.to_excel('BriansKeeper.xlsx', index=False)

print("Authentication localization added successfully!")
print(f"Added {len(new_entries)} new entries")
