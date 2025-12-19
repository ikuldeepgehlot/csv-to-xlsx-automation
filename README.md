📊 CSV to Excel Automation Tool   

📌 Business Problem

While working with Amazon Seller Central reports, I frequently downloaded reports in CSV format and immediately started working on them in Excel.
However, I faced a recurring and costly issue:

- CSV files do not retain original formatting and structure
- After editing and pressing Ctrl + Save, the original raw data was permanently altered
- Reopening the same CSV often resulted in loss of original data
- Recovery of the original data was extremely difficult or impossible

This created a risk of data corruption, especially when working with critical sales and settlement reports.

🎯 Objective

To automatically convert all downloaded CSV files into XLSX format so that:

- Original raw data remains preserved
- Work can safely continue on Excel files
- Manual conversion effort is eliminated
- Data integrity is maintained at all times

🔁 Earlier Manual Process (Pain Points)

Initially, I followed this approach:

- Download CSV file from Amazon Seller Central
- Immediately save it as .xlsx
- Start working on the Excel file

Over time:
- CSV files started piling up
- Manually converting multiple files became time-consuming
- There was a risk of forgetting to convert a file before editing

This led to the idea of building an automation solution.

🔧 Tools & Technologies Used

- Python
- csv module
- os module
- openpyxl
- Visual Studio Code
- ChatGPT-assisted development

🔄 Automation Workflow

- Monitor the Downloads folder
- Detect all .csv files
- Check whether a converted .xlsx file already exists
- Convert CSV → XLSX automatically
- Preserve all rows and columns exactly as-is
- Skip already converted files to avoid duplication

🧩 Key Automation Logic
- Reads CSV files row-by-row using Python’s csv module
- Writes data into an Excel workbook using openpyxl
- Ensures no assumptions about where the data starts
- Handles files where data may begin from:
  - 1st row
  - 4th row
  - 6th row (or any row)
- Prevents overwriting existing Excel files

⚠️ Challenges Faced & Solution  

❌ Initial Issue

Some Amazon CSV reports had:

- Metadata or blank rows
- Actual data starting from the 4th or 6th row

The initial script:

- Detected only visible rows before the actual dataset
- Resulted in incorrect or partial conversion

✅ Solution Applied

- Modified the logic to copy every row exactly as present in the CSV
- Removed assumptions about header or data start position
- Ensured 100% raw data preservation

After modification, the script worked accurately for all CSV structures.

🚀 Usability Enhancement (EXE Version)
Problem
- Script required Python & VS Code
- Not all colleagues had Python installed

Enhancement

- Converted the Python script into a standalone .exe file

- The EXE:

    - Works without Python installation
    - Can be shared with colleagues
    - Requires only a double-click to run

This significantly improved usability and adoption across the team.

📈 Business Impact

- Eliminated manual CSV → XLSX conversion
- Prevented accidental data loss
- Improved data accuracy and reliability
- Saved daily operational time
- Enabled non-technical users to use the tool via EXE

📂 Folder Structure
```
csv-to-xlsx-automation/
│
├── auto_csv_to_excel.py
├── README.md
└── dist/
    └── CSV_to_Excel_Converter.exe
```
🧪 Python Automation Script
```import csv
import os
from openpyxl import Workbook

folder = r"C:\Users\Abcom\Downloads"

def convert_csv_to_xlsx(csv_path):
    xlsx_path = csv_path.replace(".csv", ".xlsx")

    wb = Workbook()
    ws = wb.active

    with open(csv_path, "r", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        for row in reader:
            ws.append(row)

    wb.save(xlsx_path)

# Process existing CSVs
for file in os.listdir(folder):
    if file.lower().endswith(".csv"):
        csv_path = os.path.join(folder, file)
        xlsx_path = csv_path.replace(".csv", ".xlsx")

        if os.path.exists(xlsx_path):
            print(f"[SKIPPED] Already converted → {file}")
            continue

        convert_csv_to_xlsx(csv_path)
        print(f"[DONE] Converted → {file}")
```

 📝 **Notes**
- Designed specifically for Amazon Seller Central CSV reports
- Can be extended to other CSV-based workflows
- Built with a business-first automation mindset

>Note:
This project was developed with the assistance of ChatGPT for ideation, debugging, and optimization.
Final logic, testing, and business understanding were performed by me.

⭐ **Why This Project Matters**

This is not a practice project -
It is a real-world automation solution used to solve a genuine operational problem.
