## **Steps to run the script:**

### **1. Save the script**
Save the final version as `excel_yellow_to_memoq.py` in the same location as your Excel file.

### **2. Open Command Prompt**
- Press `Windows + R`, type `cmd`, and press Enter

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

**Or with the full path:**
```cmd
python excel_yellow_to_memoq.py "C:\Users\user_name\folder_name\file_name.xlsx"
```

### **6. If no yellow cells are found, use the --all flag:**
```cmd
python excel_yellow_to_memoq.py "file_name.xlsx" --all
```

## **Expected output:**
The script will create a new file called `file_name_memoQ.xlsx` in the same folder, ready for manual import into memoQ.

The output file will have columns: **Komponente | Source | FR | IT | EN** with only the rows that had yellow highlighting in the original file.
