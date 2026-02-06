# ============================================================================
# memoQ TERMBASE - MISSING TRANSLATIONS FINDER
# ============================================================================
# This script finds entries in your termbase that are missing translations
# and creates separate Excel files for each language that needs translation.
# ============================================================================

import pandas as pd
import os

# SETTINGS - Change these if needed

INPUT_FILE = "TB General.xlsx"  # Change if different file name
OUTPUT_FOLDER = "missing_translations"  # Folder where results will be saved

SOURCE_LANGUAGE = "German"  # The language you're translating FROM
TARGET_LANGUAGES = ["English", "French", "Italian"]  # Languages you're translating TO

# MAIN PROGRAM - You don't need to change anything below

def main():
    """
    Main function that does all the work
    """
    
    print("\n" + "="*70)
    print("  memoQ TERMBASE - MISSING TRANSLATIONS FINDER")
    print("="*70)
    
    # Step 1: Check if the input file exists
    print(f"\nStep 1: Looking for file '{INPUT_FILE}'...")
    if not os.path.exists(INPUT_FILE):
        print(f"❌ ERROR: Cannot find '{INPUT_FILE}'")
        print("\nPlease make sure:")
        print("  - The file is in the same folder as this script")
        print("  - The filename is spelled correctly")
        input("\nPress Enter to exit...")
        return
    print("✓ File found!")
    
    # Step 2: Read the Excel file
    print(f"\nStep 2: Reading the Excel file...")
    try:
        data = pd.read_excel(INPUT_FILE)
        print(f"✓ Successfully read {len(data)} entries")
    except Exception as e:
        print(f"❌ ERROR reading file: {e}")
        input("\nPress Enter to exit...")
        return
    
    # Step 3: Find the columns for each language
    print(f"\nStep 3: Finding language columns...")
    
    # Get all column names
    all_columns = data.columns.tolist()
    
    # Find where German columns start (first German column is the term)
    german_positions = [i for i, col in enumerate(all_columns) if col == SOURCE_LANGUAGE]
    
    if len(german_positions) == 0:
        print(f"❌ ERROR: Cannot find '{SOURCE_LANGUAGE}' columns in the file")
        input("\nPress Enter to exit...")
        return
    
    german_term_position = german_positions[0]
    print(f"✓ Found {SOURCE_LANGUAGE} term in column {german_term_position}")
    
    # Find positions for each target language
    target_positions = {}
    for language in TARGET_LANGUAGES:
        positions = [i for i, col in enumerate(all_columns) if col == language]
        if len(positions) > 0:
            target_positions[language] = positions[0]
            print(f"✓ Found {language} term in column {positions[0]}")
        else:
            print(f"⚠ Warning: Cannot find '{language}' columns")
    
    # Step 4: Show statistics
    print(f"\nStep 4: Analyzing termbase...")
    print("-" * 70)
    
    # Count how many German terms exist
    german_column = data.iloc[:, german_term_position]
    german_not_empty = (german_column.notna()) & (german_column.astype(str).str.strip() != '')
    total_german = german_not_empty.sum()
    
    print(f"Total entries with {SOURCE_LANGUAGE}: {total_german}")
    print("\nMissing translations:")
    
    # For each target language, count missing translations
    missing_counts = {}
    for language, position in target_positions.items():
        target_column = data.iloc[:, position]
        target_is_empty = (target_column.isna()) | (target_column.astype(str).str.strip() == '')
        
        # Missing = German exists BUT target language is empty
        missing = german_not_empty & target_is_empty
        count = missing.sum()
        missing_counts[language] = count
        
        if count > 0:
            print(f"  {language}: {count} entries missing")
        else:
            print(f"  {language}: All entries translated ✓")
    
    # Step 5: Create output folder
    print(f"\nStep 5: Creating output folder...")
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"✓ Created folder '{OUTPUT_FOLDER}'")
    else:
        print(f"✓ Folder '{OUTPUT_FOLDER}' already exists")
    
    # Step 6: Create Excel files for each language with missing translations
    print(f"\nStep 6: Creating Excel files...")
    print("-" * 70)
    
    files_created = 0
    
    for language, position in target_positions.items():
        
        # Skip if no missing translations
        if missing_counts[language] == 0:
            print(f"  {language}: Skipped (no missing translations)")
            continue
        
        print(f"  {language}: Processing...")
        
        # Find rows where German exists but this language is empty
        german_column = data.iloc[:, german_term_position]
        target_column = data.iloc[:, position]
        
        german_not_empty = (german_column.notna()) & (german_column.astype(str).str.strip() != '')
        target_is_empty = (target_column.isna()) | (target_column.astype(str).str.strip() == '')
        
        missing_rows = german_not_empty & target_is_empty
        
        # Get only the rows with missing translations
        data_to_export = data[missing_rows].copy()
        
        # Create filename
        output_filename = f"Missing_{language}.xlsx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        # Save to Excel
        data_to_export.to_excel(output_path, index=False, engine='openpyxl')
        
        print(f"  {language}: ✓ Created '{output_filename}' ({len(data_to_export)} entries)")
        files_created += 1
    
    # Step 7: Done!
    print("\n" + "="*70)
    print("✅ COMPLETED!")
    print("="*70)
    
    if files_created > 0:
        print(f"\n✓ Created {files_created} Excel file(s) in '{OUTPUT_FOLDER}' folder")
        print("\nNext steps:")
        print(f"  1. Open the '{OUTPUT_FOLDER}' folder")
        print("  2. Open the Excel files")
        print("  3. Fill in the missing translations")
        print("  4. Import the completed files back into memoQ")
    else:
        print("\n✓ No files created - all translations are complete!")
    
    print("\n" + "="*70)
    input("\nPress Enter to exit...")

# RUN THE PROGRAM

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + "="*70)
        print("❌ AN ERROR OCCURRED")
        print("="*70)
        print(f"\nError message: {e}")
        print("\nIf you need help, copy this entire error message.")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
