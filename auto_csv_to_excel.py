import csv
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
