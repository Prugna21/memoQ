import pandas as pd
import openpyxl
import argparse
import os

def detect_yellow_cells(file_path):
    """Find cells with yellow highlighting"""
    workbook = openpyxl.load_workbook(file_path, data_only=False)
    worksheet = workbook.active
    
    yellow_cells = {}
    for row_num, row in enumerate(worksheet.iter_rows(min_row=1), 1):
        for col_num, cell in enumerate(row, 1):
            if cell.fill and cell.fill.start_color:
                color_index = str(cell.fill.start_color.index)
                if 'FFFF' in color_index.upper():
                    if row_num not in yellow_cells:
                        yellow_cells[row_num] = {}
                    yellow_cells[row_num][col_num] = cell.value
    
    workbook.close()
    print(f"Found yellow cells in {len(yellow_cells)} rows")
    return yellow_cells

def prepare_for_memoq(input_file, output_file):
    """Extract yellow cells and prepare for memoQ"""
    
    yellow_cells = detect_yellow_cells(input_file)
    if not yellow_cells:
        print("No yellow cells found")
        return None
    
    df = pd.read_excel(input_file)
    
    # Find column mappings
    source_col = None
    for col in df.columns:
        if str(col).upper() in ['DE', 'DEU']:
            source_col = col
            break
    
    if not source_col:
        print("No German source column found")
        return None
    
    lang_mapping = {}
    lang_patterns = {
        'FR': ['FR', 'FRA'], 
        'IT': ['IT', 'ITA'], 
        'EN': ['EN', 'ENG']
    }
    
    for target_lang, patterns in lang_patterns.items():
        for col in df.columns:
            if str(col).upper() in patterns:
                lang_mapping[target_lang] = col
                break
    
    # Map column positions
    col_positions = {col: i + 1 for i, col in enumerate(df.columns)}
    
    # Process rows with yellow cells
    results = []
    for row_num, yellow_cols in yellow_cells.items():
        if row_num <= 1:
            continue
            
        df_index = row_num - 2
        if df_index < 0 or df_index >= len(df):
            continue
            
        original_row = df.iloc[df_index]
        row_data = {}
        
        # Add Komponente if not empty
        if 'Komponente' in df.columns:
            komponente = original_row['Komponente']
            if pd.notna(komponente) and str(komponente).strip():
                row_data['Komponente'] = komponente
        
        # Add German source (always)
        row_data['Source'] = original_row[source_col]
        
        # Add target languages (only if yellow)
        for target_lang, source_col_name in lang_mapping.items():
            col_pos = col_positions[source_col_name]
            if col_pos in yellow_cols:
                row_data[target_lang] = yellow_cols[col_pos]
            else:
                row_data[target_lang] = ''
        
        if row_data['Source'] and str(row_data['Source']).strip():
            results.append(row_data)
    
    if not results:
        print("No valid segments found")
        return None
    
    # Save results
    result_df = pd.DataFrame(results)
    result_df.to_excel(output_file, index=False)
    
    return len(results)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Extract yellow cells for memoQ')
    parser.add_argument('input_file', help='Input Excel file')
    parser.add_argument('--output', help='Output file name')
    
    args = parser.parse_args()
    
    if not args.output:
        name, ext = os.path.splitext(args.input_file)
        args.output = f"{name}_memoQ.xlsx"
    
    print(f"Processing: {args.input_file}")
    
    count = prepare_for_memoq(args.input_file, args.output)
    
    if count:
        print(f"SUCCESS! Processed {count} segments")
        print(f"Output: {args.output}")
    else:
        print("No segments processed")
