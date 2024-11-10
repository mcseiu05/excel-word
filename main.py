import openpyxl
from docxtpl import DocxTemplate
from datetime import datetime
from pathlib import Path

# Constants
EXCEL_PATH = Path("performance-summary.xlsx")
TEMPLATE_PATH = Path("performance-review.docx")
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
        "employee_id": value_tuple[0],
        "employee_name": value_tuple[1],
        "communication": value_tuple[2],
        "problem_solving": value_tuple[3],
        "teamwork": value_tuple[4],
        "punctuality": value_tuple[5],
        "total_score": value_tuple[6],
        "remarks": value_tuple[7],
    }
    doc.render(context)

    # Generate unique filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    doc_name = OUTPUT_DIR / f"performance-review_{value_tuple[0]}_{timestamp}.docx"
    doc.save(doc_name)

print("Documents generated successfully!")
