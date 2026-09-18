---
description: Add localization for new languages to Excel localization files
---

# Add Localization Workflow

This workflow helps add new language columns to Excel localization files and automatically translates keys.

## Usage
Run this workflow when you need to:
- Add a new language column to existing localization file
- Automatically generate translations based on keys
- Update localization files with new translations

## Steps

1. **Check current localization file structure**
   ```bash
   python -c "import pandas as pd; df = pd.read_excel('Localization.xlsx'); print('Current structure:'); print(df.to_string()); print(f'\nShape: {df.shape}'); print('Columns:', df.columns.tolist())"
   ```

2. **Create localization update script**
   // turbo
   ```bash
   cat > update_localization.py << 'EOF'
import pandas as pd
import sys

def add_localization_language(file_path='Localization.xlsx', language_code='en', translations_dict=None):
    """
    Add new language column to localization file and fill with translations
    
    Args:
        file_path: Path to Excel localization file
        language_code: Language code for new column (e.g., 'en', 'de', 'fr')
        translations_dict: Dictionary of key->translation mappings
    """
    try:
        # Read existing file
        df = pd.read_excel(file_path)
        print(f"Loaded {file_path} with {len(df)} entries")
        print(f"Current columns: {df.columns.tolist()}")
        
        # Add new language column if not exists
        if language_code not in df.columns:
            df[language_code] = ''
            print(f"Added column '{language_code}' for {language_code} language")
        else:
            print(f"Column '{language_code}' already exists")
        
        # Default translations based on keys if not provided
        if translations_dict is None:
            translations_dict = generate_default_translations(df['key'].tolist(), language_code)
        
        # Fill translations
        updated_count = 0
        for idx, row in df.iterrows():
            key = row['key']
            if key in translations_dict:
                df.at[idx, language_code] = translations_dict[key]
                updated_count += 1
        
        print(f"Updated {updated_count} translations for {language_code}")
        
        # Save updated file
        df.to_excel(file_path, index=False)
        print(f"Successfully saved updated {file_path}")
        
        # Show result
        print("\nUpdated localization file:")
        print(df.to_string())
        
        return df
        
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

def generate_default_translations(keys, language_code):
    """Generate default translations based on key patterns"""
    translations = {}
    
    # Common patterns for different languages
    if language_code == 'en':
        patterns = {
            'initialization_title': 'Initialization',
            'main_menu_title': 'Main Menu',
            'main_menu_button_start': 'Start Game',
            'game_title': 'Game',
            'show_window_button': 'Show Window',
            'init_localization': 'Initialize Localization',
            'load_assets': 'Load Assets',
            'login_state': 'Login State',
            'health': 'Health',
            'settings': 'Settings',
            'exit': 'Exit',
            'save': 'Save',
            'load': 'Load',
            'cancel': 'Cancel',
            'confirm': 'Confirm',
            'yes': 'Yes',
            'no': 'No',
            'ok': 'OK',
            'error': 'Error',
            'warning': 'Warning',
            'info': 'Information',
            'success': 'Success',
            'failed': 'Failed',
            'loading': 'Loading',
            'please_wait': 'Please Wait',
            'connection_lost': 'Connection Lost',
            'retry': 'Retry',
            'back': 'Back',
            'next': 'Next',
            'previous': 'Previous',
            'finish': 'Finish',
            'help': 'Help',
            'about': 'About',
            'version': 'Version',
            'player': 'Player',
            'score': 'Score',
            'time': 'Time',
            'level': 'Level',
            'pause': 'Pause',
            'resume': 'Resume',
            'restart': 'Restart',
            'quit': 'Quit'
        }
    elif language_code == 'de':
        patterns = {
            'initialization_title': 'Initialisierung',
            'main_menu_title': 'Hauptmenü',
            'main_menu_button_start': 'Spiel Starten',
            'game_title': 'Spiel',
            'show_window_button': 'Fenster Zeigen',
            'init_localization': 'Lokalisierung Initialisieren',
            'load_assets': 'Assets Laden',
            'login_state': 'Anmeldestatus',
            'health': 'Gesundheit',
            'settings': 'Einstellungen',
            'exit': 'Beenden',
            'save': 'Speichern',
            'load': 'Laden',
            'cancel': 'Abbrechen',
            'confirm': 'Bestätigen',
            'yes': 'Ja',
            'no': 'Nein',
            'ok': 'OK',
            'error': 'Fehler',
            'warning': 'Warnung',
            'info': 'Information',
            'success': 'Erfolg',
            'failed': 'Fehlgeschlagen',
            'loading': 'Laden',
            'please_wait': 'Bitte Warten',
            'retry': 'Wiederholen',
            'back': 'Zurück',
            'next': 'Weiter',
            'previous': 'Vorherige',
            'finish': 'Beenden',
            'help': 'Hilfe',
            'about': 'Über',
            'version': 'Version',
            'player': 'Spieler',
            'score': 'Punkte',
            'time': 'Zeit',
            'level': 'Level',
            'pause': 'Pause',
            'resume': 'Fortsetzen',
            'restart': 'Neustarten',
            'quit': 'Verlassen'
        }
    elif language_code == 'fr':
        patterns = {
            'initialization_title': 'Initialisation',
            'main_menu_title': 'Menu Principal',
            'main_menu_button_start': 'Démarrer le Jeu',
            'game_title': 'Jeu',
            'show_window_button': 'Afficher la Fenêtre',
            'init_localization': 'Initialiser la Localisation',
            'load_assets': 'Charger les Assets',
            'login_state': 'État de Connexion',
            'health': 'Santé',
            'settings': 'Paramètres',
            'exit': 'Quitter',
            'save': 'Sauvegarder',
            'load': 'Charger',
            'cancel': 'Annuler',
            'confirm': 'Confirmer',
            'yes': 'Oui',
            'no': 'Non',
            'ok': 'OK',
            'error': 'Erreur',
            'warning': 'Avertissement',
            'info': 'Information',
            'success': 'Succès',
            'failed': 'Échec',
            'loading': 'Chargement',
            'please_wait': 'Veuillez Patienter',
            'retry': 'Réessayer',
            'back': 'Retour',
            'next': 'Suivant',
            'previous': 'Précédent',
            'finish': 'Terminer',
            'help': 'Aide',
            'about': 'À Propos',
            'version': 'Version',
            'player': 'Joueur',
            'score': 'Score',
            'time': 'Temps',
            'level': 'Niveau',
            'pause': 'Pause',
            'resume': 'Reprendre',
            'restart': 'Redémarrer',
            'quit': 'Quitter'
        }
    
    # Generate translations for keys
    for key in keys:
        if key in patterns:
            translations[key] = patterns[key]
        else:
            # Generate basic translation by converting key to readable format
            translations[key] = key.replace('_', ' ').title()
    
    return translations

if __name__ == "__main__":
    # Example usage for English
    add_localization_language('Localization.xlsx', 'en')
EOF
   ```

3. **Run the localization update script**
   ```bash
   python update_localization.py
   ```

4. **Verify the updated file**
   ```bash
   python -c "import pandas as pd; df = pd.read_excel('Localization.xlsx'); print('Final result:'); print(df.to_string()); print(f'\nFile format: {df.shape}'); print('Columns:', df.columns.tolist())"
   ```

5. **Clean up temporary script**
   ```bash
   del update_localization.py
   ```

## Customization

To add a different language, modify step 3:
```bash
python -c "import sys; sys.path.append('.'); from update_localization import add_localization_language; add_localization_language('Localization.xlsx', 'de')"  # For German
python -c "import sys; sys.path.append('.'); from update_localization import add_localization_language; add_localization_language('Localization.xlsx', 'fr')"  # For French
```

## Custom Translations

To provide custom translations, create a dictionary and pass it to the function:
```python
custom_translations = {
    'your_key': 'Your Custom Translation',
    'another_key': 'Another Translation'
}
add_localization_language('Localization.xlsx', 'en', custom_translations)
```
