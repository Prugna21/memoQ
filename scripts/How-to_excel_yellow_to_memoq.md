# **Excel Yellow Script**

## Table of Contents
- [What the Script Does](#what-the-script-does)
- [Use Case](#use-case)
- [Step-by-Step Breakdown](#step-by-step-breakdown)
  - [1. Import Libraries](#1-import-libraries)
  - [2. Function: detect_yellow_cells(file_path)](#2-function-detect_yellow_cellsfile_path)
  - [3. Function: prepare_for_memoq(input_file, output_file)](#3-function-prepare_for_memoqinput_file-output_file)
    - [3a: Get Yellow Cells](#3a-get-yellow-cells)
    - [3b: Read Excel Data](#3b-read-excel-data)
    - [3c: Find Source Column (German)](#3c-find-source-column-german)
    - [3d: Map Target Language Columns](#3d-map-target-language-columns)
    - [3e: Process Rows with Yellow Cells](#3e-process-rows-with-yellow-cells)
    - [3f: Create Output Excel File](#3f-create-output-excel-file)
    - [3g: Format Output File](#3g-format-output-file)
  - [4. Main Execution Block](#4-main-execution-block)
- [Steps to Run the Script](#steps-to-run-the-script-excel_yellow_to_memoqpy)
  - [1. Save the script](#1-save-the-script)
  - [2. Open Command Prompt](#2-open-command-prompt)
  - [3. Navigate to the folder](#3-navigate-to-the-folder-where-you-saved-the-file-and-the-script)
  - [4. Install required packages](#4-install-required-packages-if-not-already-installed)
  - [5. Run the script](#5-run-the-script)
- [Expected Output](#expected-output)
- [Compatibility](#compatibility)
  - [What works universally](#what-works-universally)
  - [What might need adaptation](#what-might-need-adaptation-for-other-users)
  - [To make it more user-friendly](#to-make-it-more-user-friendly-for-others)

## **What the Script Does**
1. **Input**: Takes an Excel file with multilingual content
2. **Detection**: Finds all cells highlighted in yellow
3. **Extraction**: Extracts only rows that contain yellow-highlighted cells
4. **Processing**:
   - Always includes German source text
   - Only includes target language columns that are highlighted yellow
   - Preserves component information if available
5. **Output**: Creates a new Excel file formatted for memoQ translation tool

## **Use Case**
- Translators have to update/revise only the cells in yellow
- Project managers need to extract only the highlighted segments
- The extracted content needs to be imported into memoQ for translation

## **Step-by-Step Breakdown**

### **1. Import Libraries**
```python
import pandas as pd
import openpyxl
import argparse
import os
```
- **pandas**: For data manipulation and Excel file handling
- **openpyxl**: For working with Excel formatting (colors, styles)
- **argparse**: For command-line argument parsing
- **os**: For file path operations

### **2. Function: `detect_yellow_cells(file_path)`**
This function identifies cells with yellow highlighting:
```python
def detect_yellow_cells(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=False)
    worksheet = workbook.active
```
- Opens the Excel file while preserving formatting
- Gets the active worksheet

```python
    yellow_cells = {}
    for row_num, row in enumerate(worksheet.iter_rows(min_row=1), 1):
        for col_num, cell in enumerate(row, 1):
```
- Creates a dictionary to store yellow cell locations
- Iterates through each row and column (starting from row 1, column 1)

```python
            if cell.fill and cell.fill.start_color:
                color_index = str(cell.fill.start_color.index)
                if 'FFFF' in color_index.upper():
```
- Checks if the cell has a fill color
- Looks for 'FFFF' in the color code (indicating yellow/bright colors)
- Stores the cell's position and value if it's yellow

### **3. Function: `prepare_for_memoq(input_file, output_file)`**
This is the main processing function:

#### **3a: Get Yellow Cells**
```python
yellow_cells = detect_yellow_cells(input_file)
if not yellow_cells:
    print("No yellow cells found")
    return None
```

#### **3b: Read Excel Data**
```python
df = pd.read_excel(input_file)
```

#### **3c: Find Source Column (German)**
```python
source_col = None
for col in df.columns:
    if str(col).upper() in ['DE', 'DEU']:
        source_col = col
        break
```
- Looks for German source column (labeled 'DE' or 'DEU')

#### **3d: Map Target Language Columns**
```python
lang_mapping = {}
lang_patterns = {
    'FR': ['FR', 'FRA'],
    'IT': ['IT', 'ITA'],
    'EN': ['EN', 'ENG']
}
```
- Creates mappings for French, Italian, and English columns
- Looks for various column name patterns (FR/FRA, IT/ITA, EN/ENG)

#### **3e: Process Rows with Yellow Cells**
```python
for row_num, yellow_cols in yellow_cells.items():
    if row_num <= 1:
        continue  # Skip header row
```
For each row with yellow cells:
- Adds "Komponente" field if it exists and isn't empty
- Always includes the German source text
- Only includes target language columns that have yellow highlighting
- Stores empty string for non-highlighted target languages

#### **3f: Create Output Excel File**
```python
result_df = pd.DataFrame(results)
writer = pd.ExcelWriter(output_file, engine='openpyxl')
result_df.to_excel(writer, index=False)
```
- Creates a new DataFrame from the processed results
- Writes to Excel file

#### **3g: Format Output File**
- Copies column widths from the original file
- Enables text wrapping for all cells
- Sets vertical alignment to top

### **4. Main Execution Block**
```python
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Extract yellow cells for memoQ')
    parser.add_argument('input_file', help='Input Excel file')
    parser.add_argument('--output', help='Output file name')
```
**Command-line usage:**
```bash
python script.py input_file.xlsx --output output_file.xlsx
```
**Auto-naming:** If no output file is specified, it creates one by adding "_memoQ" to the input filename.

## **Steps to run the script excel_yellow_to_memoq.py**

### **1. Save the script**
Save the final version as `excel_yellow_to_memoq.py` in the same location as your Excel file.

### **2. Open Command Prompt**
- Press `Windows + R`, type `cmd`, and press Enter.

### **3. Navigate to the folder, where you saved the file and the script**
```cmd
cd "C:\Users\user_name\folder_name"
```

### **4. Install required packages** (if not already installed)
```cmd
pip install pandas openpyxl xlrd
```

### **5. Run the script**
```cmd
python excel_yellow_to_memoq.py "file_name.xlsx"
```
**Or with the full path**
```cmd
python excel_yellow_to_memoq.py "C:\Users\user_name\folder_name\file_name.xlsx"
```

## **Expected Output**
The script will create a new file called `file_name_memoQ.xlsx` in the same folder, ready for manual import into memoQ.
The output file will have columns: **Komponente | Source | FR | IT | EN** with only the rows that had yellow highlighting in the original file.

## **Compatibility**
The script is designed to be quite flexible and should work for most users and documents with similar formats, but there are a few things to consider.

### **What works universally**
1. **Column name detection** - The script automatically detects columns by name patterns (DE/DEU, FR/FRA, IT/ITA, EN/ENG) regardless of column order.
2. **Yellow highlighting detection** - Works with standard Excel yellow highlighting.
3. **File path handling** - Works with any valid file path through command-line arguments.
4. **Different Excel versions** - Handles both older (.xls) and newer (.xlsx).

### **What might need adaptation for other users**
1. **Language combinations** - Set up for German→French/Italian/English. If someone needs different languages (e.g., Spanish, Portuguese), they'd need to:
   - Modify the `target_languages` default parameter
   - Add new language detection patterns in the column mapping
2. **Column naming conventions** - If documents use completely different column names (e.g., "Source_Language" instead of "DE"), the detection patterns would need updating
3. **Python environment** - Other users need to:
   - Have Python installed
   - Install required packages `pandas`, `openpyxl`)
   - Know how to use command line
4. **File structure expectations** - The script expects a specific structure (header row, data rows, component column optional)

### **To make it more user-friendly for others**
You could create a simple configuration file or add command-line options for:
- Custom language codes
- Custom column name patterns
- Different highlighting colors
