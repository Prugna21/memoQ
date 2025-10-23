## **Steps to run the script**

#### **1. Save the script**
Save the final version as `excel_yellow_to_memoq.py` in the same location as your Excel file.

#### **2. Open Command Prompt**
- Press `Windows + R`, type `cmd`, and press Enter.

#### **3. Navigate to the folder, where you saved the file and the script**
```cmd
cd "C:\Users\user_name\folder_name"
```

#### **4. Install required packages** (if not already installed)
```cmd
pip install pandas openpyxl xlrd
```

#### **5. Run the script**
```cmd
python excel_yellow_to_memoq.py "file_name.xlsx"
```

**Or with the full path**
```cmd
python excel_yellow_to_memoq.py "C:\Users\user_name\folder_name\file_name.xlsx"
```

#### **6. If no yellow cells are found, use the --all flag**
```cmd
python excel_yellow_to_memoq.py "file_name.xlsx" --all
```

### **Expected output**
The script will create a new file called `file_name_memoQ.xlsx` in the same folder, ready for manual import into memoQ.

The output file will have columns: **Komponente | Source | FR | IT | EN** with only the rows that had yellow highlighting in the original file.

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
   - Install required packages (`pandas`, `openpyxl`)
   - Know how to use command line

4. **File structure expectations** - The script expects a specific structure (header row, data rows, component column optional)

### **To make it more user-friendly for others**

You could create a simple configuration file or add command-line options for:
- Custom language codes
- Custom column name patterns
- Different highlighting colors
