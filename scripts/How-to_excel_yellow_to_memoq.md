# **Excel Yellow Script**

## Table of Contents
- [What the Script Does](#what-the-script-does)
- [Use Case](#use-case)
- [How to Run the Script](#how-to-run-the-script)
  - [Command Line Mode](#command-line-mode)
- [Step-by-Step Breakdown](#step-by-step-breakdown)
  - [1. Import Libraries](#1-import-libraries)
  - [2. Function: find_yellow_cells(file_name)](#2-function-find_yellow_cellsfile_name)
  - [3. Function: create_memoq_file(input_file, output_file)](#3-function-create_memoq_fileinput_file-output_file)
    - [3a: Find Yellow Cells](#3a-find-yellow-cells)
    - [3b: Read Excel Data](#3b-read-excel-data)
    - [3c: Find Source Column (German)](#3c-find-source-column-german)
    - [3d: Map Target Language Columns](#3d-map-target-language-columns)
    - [3e: Process Rows with Yellow Cells](#3e-process-rows-with-yellow-cells)
    - [3f: Create Output Excel File with Formatting](#3f-create-output-excel-file-with-formatting)
  - [4. Function: save_excel_with_formatting()](#4-function-save_excel_with_formatting)
  - [5. Mode Detection](#5-mode-detection)
- [Installation](#installation)
- [Expected Output](#expected-output)
- [Advanced Features](#advanced-features)
- [Compatibility](#compatibility)

## **What the Script Does**
1. **Input**: Takes an Excel file with multilingual content
2. **Detection**: Finds all cells highlighted in yellow
3. **Extraction**: Extracts only rows that contain yellow-highlighted cells
4. **Processing**:
   - Always includes German source text
   - Only includes source or target language columns that are highlighted yellow
   - Preserves component information if available
   - Copies column widths from original file
   - Enables text wrapping
5. **Output**: Creates a new Excel file formatted for memoQ import

## **Use Case**
- Translators have to update/revise only the cells in yellow
- Project managers need to extract only the highlighted segments
- The extracted content needs to be imported into memoQ for translation

## **How to Run the Script**
Just click on the .py file. OR:

### **Command Line Mode**

#### **Command Line Steps:**
1. **Open Command Prompt**: Press `Windows + R`, type `cmd`, press Enter
2. **Navigate to folder**: 
   ```cmd
   cd "C:\Users\%username%\folder_name"
   ```

#### **Basic Usage:**
```bash
python excel_yellow_to_memoq.py "file_name.xlsx"
```

#### **With Full Paths:**
```bash
python excel_yellow_to_memoq.py "C:\Users\%username%\folder_name\file_name.xlsx"
```

#### **Custom Output Name:**
```bash
python excel_yellow_to_memoq.py "input.xlsx" "custom_output.xlsx"
```

3. **Run the script**

## **Step-by-Step Breakdown**

### **1. Import Libraries**
```python
import pandas as pd
import openpyxl
import sys
import os
```
- **pandas**: For data manipulation and Excel file
- **openpyxl**: For Excel formatting (colors, styles, column widths)
- **sys**: For command-line argument
- **os**: For file path operations and file existence checking

### **2. Function: `find_yellow_cells(file_name)`**
This function identifies cells with yellow highlighting using loops:
```python
def find_yellow_cells(file_name):
    workbook = openpyxl.load_workbook(file_name, data_only=False)
    sheet = workbook.active
```
- Opens the Excel file while preserving formatting
- Gets the active worksheet

```python
    row_number = 1
    for row in sheet.iter_rows():
        col_number = 1
        for cell in row:
            if cell.fill and cell.fill.start_color:
                color = str(cell.fill.start_color.index)
                if 'FFFF' in color.upper():  # Yellow color
```
- Uses simple counter variables
- Creates a dictionary to store yellow cell locations
- Looks for 'FFFF' in the color code (yellow/bright colors)

### **3. Function: `create_memoq_file(input_file, output_file)`**
This is the main processing function:

#### **3a: Find Yellow Cells**
```python
yellow_cells = find_yellow_cells(input_file)
if len(yellow_cells) == 0:
    print("No yellow cells found!")
    return False
```

#### **3b: Read Excel Data**
```python
data = pd.read_excel(input_file)
```

#### **3c: Find Source Column (German)**
```python
german_column = None
for column_name in data.columns:
    if column_name.upper() == 'DE' or column_name.upper() == 'DEU':
        german_column = column_name
        break
```
- Uses simple if-statements
- Looks for German source column (labeled 'DE' or 'DEU')

#### **3d: Map Target Language Columns**
```python
# Look for other language columns
for column_name in data.columns:
    column_upper = column_name.upper()
    if column_upper == 'FR' or column_upper == 'FRA':
        french_column = column_name
    elif column_upper == 'IT' or column_upper == 'ITA':
        italian_column = column_name
    elif column_upper == 'EN' or column_upper == 'ENG':
        english_column = column_name
```
- Creates mappings for French, Italian, and English columns
- Uses if-elif structure

#### **3e: Process Rows with Yellow Cells**
```python
for row_num in yellow_cells:
    if row_num <= 1:  # Skip header row
        continue
```
For each row with yellow cells:
- Adds "Komponente" field if it exists and isn't empty
- Always includes the German source text
- Only includes target language columns that have yellow highlighting
- Stores empty string for non-highlighted target languages

#### **3f: Create Output Excel File with Formatting**
```python
save_excel_with_formatting(results, input_file, output_file)
```
- Calls specialised function to handle formatting
- Creates a new DataFrame from the processed results
- Applies professional formatting

### **4. Function: `save_excel_with_formatting()`**

```python
def save_excel_with_formatting(results, input_file, output_file):
    # Create DataFrame and Excel writer
    final_data = pd.DataFrame(results)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    # Copy column widths from original file
    original_width = input_sheet.column_dimensions[column_letter].width
    output_sheet.column_dimensions[column_letter].width = original_width
    
    # Set text wrapping for all cells
    cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
```
**Features:**
- Copies column widths from original file
- Enables text wrapping for all cells
- Sets vertical alignment to top

### **5. Mode Detection**

```python
if len(sys.argv) > 1:
    # Command line mode - arguments were provided
    run_command_line_mode()
else:
    # Interactive mode - no arguments provided  
    run_interactive_mode()
```
- Automatically detects if user provided command-line arguments
- Switches between interactive mode and command-line mode

## **Installation**

### **Install Required Packages:**
```cmd
pip install pandas openpyxl
```

### **Save the Script:**
Save as `excel_yellow_to_memoq.py` in desired location.

## **Expected Output**
The script will create a new file called `file_name_memoQ.xlsx` with:
- **Columns**: `Komponente | Source | FR | IT | EN` 
- **Content**: Only rows that had yellow highlighting in the original file
- **Formatting**: 
  - Same column widths as original file
  - Text wrapping enabled
  - Alignment

## **Advanced Features**

### **Automatic File Naming**
- Input: `myfile.xlsx` → Output: `myfile_memoQ.xlsx`
- Input: `data.xls` → Output: `data_memoQ.xlsx`

### **Error Handling**
- File existence checking
- Missing column detection  
- Formatting failure recovery

### **Dual Mode Operation**
- **Interactive**: For occasional use
- **Command Line**: For automation and batch processing

## **Compatibility**

### **What works universally**
1. **Professional features**: Command-line support for advanced users
2. **Column name detection**: Automatically detects columns by name patterns (DE/DEU, FR/FRA, IT/ITA, EN/ENG)
3. **Yellow highlighting detection**: Works with standard Excel yellow highlighting
4. **File path handling**: Works with any valid file path and handles spaces in names
5. **Excel version support**: Handles both older (.xls) and newer (.xlsx)
6. **Formatting preservation**: Maintains proper column widths and text wrapping

### **What might need adaptation for other users**
1. **Language combinations**: Currently set up for German→French/Italian/English
2. **Column naming conventions**: Expects specific column name patterns
3. **Python environment**: Requires Python and package installation
4. **File structure**: Expects header row and specific data structure
