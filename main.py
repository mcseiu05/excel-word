

import openpyxl
from docxtpl import DocxTemplate
import datetime

# Load data from Excel
path = "F:\Work\excel-word\marks_data.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active

list_values = list(sheet.values)
print(list_values)

# Generate docs
doc = DocxTemplate("mark-sheet.docx")

for value_tuple in list_values[1:11]:
    doc.render({"role": value_tuple[0],
                "cse_501": value_tuple[1],
                "cse_502": value_tuple[2],
                "cse_503": value_tuple[3],
                "cse_504": value_tuple[4],
                "cse_505": value_tuple[5],
                "cse_506": value_tuple[6],
                "cse_507": value_tuple[7],
                })

    doc_name = "mark-sheet" + value_tuple[0] + ".docx"
    doc.save(doc_name)

