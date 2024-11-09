import openpyxl
from docxtpl import DocxTemplate
from datetime import datetime
from pathlib import Path

# Constants
EXCEL_PATH = Path("F:/Startup/RnD/excel-word/marks_data.xlsx")
TEMPLATE_PATH = Path("mark-sheet.docx")
OUTPUT_DIR = Path("generated_docs")

# Ensure output directory exists
OUTPUT_DIR.mkdir(exist_ok=True)

# Load data from Excel
try:
    workbook = openpyxl.load_workbook(EXCEL_PATH)
    sheet = workbook.active
    list_values = list(sheet.values)
except Exception as e:
    print(f"Error loading Excel file: {e}")
    exit(1)

# Skip header row and process data
for value_tuple in list_values[1:]:
    doc = DocxTemplate(TEMPLATE_PATH)
    context = {
        "role":    value_tuple[0],
        "cse_501": value_tuple[1],
        "cse_502": value_tuple[2],
        "cse_503": value_tuple[3],
        "cse_504": value_tuple[4],
        "cse_505": value_tuple[5],
        "cse_506": value_tuple[6],
        "cse_507": value_tuple[7],
    }
    doc.render(context)

    # Generate unique filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    doc_name = OUTPUT_DIR / f"mark-sheet_{value_tuple[0]}_{timestamp}.docx"
    doc.save(doc_name)

print("Documents generated successfully!")
