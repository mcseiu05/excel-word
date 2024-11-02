# Marks Data to Word Document Generator

This project automates the generation of mark sheets from an Excel file. It reads student data from an Excel spreadsheet and creates individual Word documents for each student using a template.

## Features

- Reads data from an Excel file containing student marks.
- Generates personalized mark sheets in Word format for each student.
- Uses `openpyxl` to handle Excel files and `docxtpl` to create Word documents.

## Requirements

- Python 3.x
- `openpyxl` library
- `docxtpl` library
- An Excel file named `marks_data.xlsx` with student marks
- A Word document template named `mark-sheet.docx`

## Installation

1. Clone this repository:
 
   git clone https://github.com/yourusername/yourrepository.git
   
2. Navigate to the project directory:
   cd yourrepository

3. Install the required libraries:
   pip install openpyxl docxtpl

Usage
Prepare your Excel file (marks_data.xlsx) with the following structure:

First row: Headers (e.g., Role, CSE_501, CSE_502, etc.)
Subsequent rows: Student data
Create a Word template (mark-sheet.docx) with placeholders for the data:

Use placeholders like {{ role }}, {{ cse_501 }}, {{ cse_502 }}, etc.

Run the script:

python your_script.py

The generated Word documents will be saved in the same directory with names formatted as mark-sheet[Role].docx.

Example
Here is a brief example of the data format in the Excel file:

Role	CSE_501	CSE_502	CSE_503	CSE_504	CSE_505	CSE_506	CSE_507
John	85	    90	    78	    92	    88	    75	    80
Alice	78	    85	    80	    88	    84	    90	    87

Contributing
Feel free to submit issues or pull requests if you have suggestions or improvements.
