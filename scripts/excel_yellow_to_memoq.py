import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import argparse
import os
import warnings
warnings.filterwarnings('ignore')

def detect_yellow_highlighted_cells(file_path):
    """Detect individual cells with yellow highlighting"""
    
    print("🔍 Scanning for yellow-highlighted cells...")
    
    workbook = openpyxl.load_workbook(file_path, data_only=False)
    worksheet = workbook.active
    
    yellow_cells = {}  # Dictionary: {row_num: {col_num: cell_value}}
    total_yellow_cells = 0
    
    for row_num, row in enumerate(worksheet.iter_rows(min_row=1), 1):
        for col_num, cell in enumerate(row, 1):
            if cell.fill and cell.fill.fill_type:
                color_found = False
                
                if hasattr(cell.fill, 'start_color') and cell.fill.start_color:
                    if hasattr(cell.fill.start_color, 'index'):
                        index_color = str(cell.fill.start_color.index)
                        if 'FFFF' in index_color.upper():
                            color_found = True
                    
                    if hasattr(cell.fill.start_color, 'rgb'):
                        rgb_color = str(cell.fill.start_color.rgb)
                        if any(yellow_rgb in rgb_color.upper() for yellow_rgb in [
                            'FFFF00', 'FFFFE0', 'FFFACD', 'FFFF99', 'FFCC99', 
                            'FFD700', 'FFFFCC', 'FFF2CC']):
                            color_found = True
                
                if color_found:
                    if row_num not in yellow_cells:
                        yellow_cells[row_num] = {}
                    yellow_cells[row_num][col_num] = cell.value
                    total_yellow_cells += 1
    
    workbook.close()
    
    print(f"✅ Found {total_yellow_cells} individual yellow cells in {len(yellow_cells)} rows")
    return yellow_cells

def prepare_yellow_cells_for_memoq(input_file, output_file, target_languages=None):
    """Extract only yellow-highlighted cells and prepare for memoQ with correct column mapping"""
    
    if target_languages is None:
        target_languages = ['FR', 'IT', 'EN']
    
    print(f"📂 Processing: {input_file}")
    
    yellow_cells = detect_yellow_highlighted_cells(input_file)
    
    if not yellow_cells:
        print("⚠️  No yellow-highlighted cells found!")
        return None
    
    # Load the data
    df = pd.read_excel(input_file, engine='openpyxl')
    print(f"✅ Loaded Excel file: {len(df)} total rows")
    print(f"📋 Available columns: {list(df.columns)}")
    
    # Find the correct source column (DE or DEU)
    source_col = None
    for col in df.columns:
        if str(col).upper() in ['DE', 'DEU', 'GERMAN', 'DEUTSCH']:
            source_col = col
            break
    
    if not source_col:
        print("❌ No German source column found (looking for DE, DEU, GERMAN, DEUTSCH)")
        return None
    
    print(f"✅ German source column: {source_col}")
    
    # Map target language columns
    lang_col_mapping = {}
    for target_lang in target_languages:
        found_col = None
        for col in df.columns:
            col_upper = str(col).upper()
            if target_lang == 'FR' and col_upper in ['FR', 'FRA', 'FRENCH', 'FRANÇAIS']:
                found_col = col
                break
            elif target_lang == 'IT' and col_upper in ['IT', 'ITA', 'ITALIAN', 'ITALIANO']:
                found_col = col
                break
            elif target_lang == 'EN' and col_upper in ['EN', 'ENG', 'ENGLISH']:
                found_col = col
                break
        
        if found_col:
            lang_col_mapping[target_lang] = found_col
            print(f"✅ {target_lang} column: {found_col}")
        else:
            print(f"⚠️ {target_lang} column not found")
    
    # Map column names to Excel column numbers (1-based)
    col_name_to_num = {}
    for i, col_name in enumerate(df.columns):
        col_name_to_num[col_name] = i + 1
    
    # Create memoQ-ready dataframe
    memoq_rows = []
    
    for row_num, yellow_cols in yellow_cells.items():
        if row_num <= 1:  # Skip header row
            continue
            
        df_index = row_num - 2  # Convert to DataFrame index
        if df_index < 0 or df_index >= len(df):
            continue
            
        original_row = df.iloc[df_index]
        row_data = {}
        
        # Check if we should include Komponente
        komponente_val = original_row.get('Komponente', '') if 'Komponente' in df.columns else ''
        if komponente_val and str(komponente_val).strip() and str(komponente_val).upper() != 'NAN':
            row_data['Komponente'] = komponente_val
        
        # Add German source (always from DE/DEU column)
        source_text = original_row[source_col]
        if source_text and str(source_text).strip():
            row_data['Source'] = source_text
            print(f"   Row {row_num}: Source (DE) = '{source_text}'")
        else:
            continue  # Skip rows without German source
        
        # Add target languages - only include content from yellow-highlighted cells
        for target_lang in target_languages:
            if target_lang in lang_col_mapping:
                target_col_name = lang_col_mapping[target_lang]
                target_col_num = col_name_to_num[target_col_name]
                
                if target_col_num in yellow_cols:
                    # This target language cell has yellow highlighting
                    yellow_text = yellow_cols[target_col_num]
                    row_data[target_lang] = yellow_text
                    print(f"     {target_lang} (YELLOW) = '{yellow_text}'")
                else:
                    # No yellow highlighting in this target language - leave empty
                    row_data[target_lang] = ''
                    print(f"     {target_lang} (no yellow) = empty")
            else:
                # Column not found - leave empty
                row_data[target_lang] = ''
        
        memoq_rows.append(row_data)
    
    if not memoq_rows:
        print("❌ No valid rows found!")
        return None
    
    # Convert to DataFrame with correct column order
    memoq_df = pd.DataFrame(memoq_rows)
    
    # Reorder columns: Komponente (if exists), Source, then target languages in order
    column_order = []
    if 'Komponente' in memoq_df.columns:
        column_order.append('Komponente')
    column_order.append('Source')
    for lang in target_languages:
        if lang in memoq_df.columns:
            column_order.append(lang)
    
    memoq_df = memoq_df[column_order]
    
    print(f"\n✅ FINAL SEGMENTS FOR MEMOQ:")
    print(f"Column order: {list(memoq_df.columns)}")
    for i, row in memoq_df.iterrows():
        source = row['Source']
        highlights = []
        for lang in target_languages:
            if lang in row and str(row[lang]).strip():
                highlights.append(f"{lang}='{row[lang]}'")
        highlight_text = " | ".join(highlights) if highlights else "No target highlights"
        print(f"   {i+1}. Source: '{source}' | Yellow: {highlight_text}")
    
    # Save simple Excel
    save_simple_excel(memoq_df, output_file)
    
    return len(memoq_df)

def save_simple_excel(df, output_file):
    """Save simple Excel without colors or special formatting"""
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Translation', index=False)
        
        worksheet = writer.sheets['Translation']
        
        # Adjust column widths based on actual columns
        col_widths = {
            'Komponente': 30,
            'Source': 80,
            'FR': 80,
            'IT': 80,
            'EN': 80
        }
        
        for col_idx, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_idx)
            width = col_widths.get(col_name, 50)
            worksheet.column_dimensions[col_letter].width = width
        
        # Freeze panes
        if 'Komponente' in df.columns:
            worksheet.freeze_panes = 'C2'  # Freeze after Komponente and Source
        else:
            worksheet.freeze_panes = 'B2'  # Freeze after Source

def process_all_rows_fallback(input_file, output_file):
    """Process all rows if no yellow highlighting found"""
    
    print("🔄 Processing all rows as fallback...")
    df = pd.read_excel(input_file, engine='openpyxl')
    
    # Find source column
    source_col = None
    for col in df.columns:
        if str(col).upper() in ['DE', 'DEU', 'GERMAN', 'DEUTSCH']:
            source_col = col
            break
    
    if not source_col:
        print("❌ No German source column found")
        return None
    
    memoq_df = pd.DataFrame()
    
    # Add Komponente if not empty
    if 'Komponente' in df.columns:
        komponente_series = df['Komponente']
        if komponente_series.notna().any() and (komponente_series.astype(str).str.strip() != '').any():
            memoq_df['Komponente'] = df['Komponente']
    
    # Add source
    memoq_df['Source'] = df[source_col]
    
    # Add target languages
    lang_mapping = {
        'FR': ['FR', 'FRA', 'FRENCH'],
        'IT': ['IT', 'ITA', 'ITALIAN'],
        'EN': ['EN', 'ENG', 'ENGLISH']
    }
    
    for target_lang, possible_cols in lang_mapping.items():
        found_col = None
        for col in df.columns:
            if str(col).upper() in possible_cols:
                found_col = col
                break
        
        if found_col:
            memoq_df[target_lang] = df[found_col]
        else:
            memoq_df[target_lang] = ''
    
    # Clean data
    memoq_df = memoq_df[memoq_df['Source'].notna() & (memoq_df['Source'].astype(str).str.strip() != '')]
    memoq_df = memoq_df.reset_index(drop=True)
    
    save_simple_excel(memoq_df, output_file)
    return len(memoq_df)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Extract yellow-highlighted cells with correct column mapping')
    parser.add_argument('input_file', help='Input Excel file')
    parser.add_argument('--output', help='Output file name')
    parser.add_argument('--languages', nargs='+', default=['FR', 'IT', 'EN'])
    parser.add_argument('--all', action='store_true', help='Process all rows if no yellow found')
    
    args = parser.parse_args()
    
    if not args.output:
        name, ext = os.path.splitext(args.input_file)
        args.output = f"{name}_memoQ.xlsx"
    
    print(f"🟡 EXCEL TO MEMOQ - COLUMN-CORRECTED VERSION")
    print(f"{'='*50}")
    print(f"Input: {args.input_file}")
    print(f"Output: {args.output}")
    print()
    
    try:
        segment_count = prepare_yellow_cells_for_memoq(
            args.input_file, args.output, args.languages
        )
        
        if segment_count is None and args.all:
            segment_count = process_all_rows_fallback(args.input_file, args.output)
        
        if segment_count:
            print(f"\n✅ SUCCESS! Prepared {segment_count} segments")
            print(f"💾 File: {args.output}")
            print(f"\n💡 Column mapping:")
            print(f"   DE/DEU → Source")
            print(f"   FR/FRA → FR")  
            print(f"   IT/ITA → IT")
            print(f"   EN/ENG → EN")
            print(f"   Komponente → removed if empty")
        else:
            print("\n⚠️ No segments processed")
            
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()