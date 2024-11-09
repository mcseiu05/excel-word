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

`git clone https://github.com/mcseiu05/excel-word.git`
   

2. Navigate to the project directory:

`cd excel-word`


3. Install the required libraries:

`pip install openpyxl docxtpl`


## Usage

- Prepare your Excel file (marks_data.xlsx) with the following structure:

   First row: Headers (e.g., Role, CSE_501, CSE_502, etc.)
   Subsequent rows: Student data

- Create a Word template (mark-sheet.docx) with placeholders for the data:

- Use placeholders like {{ role }}, {{ cse_501 }}, {{ cse_502 }}, etc.
- Please update the path of marks_data.xlsx in the main.py file.


- Run the script:

   py main.py

The generated Word documents will be saved in the same directory with names formatted as mark-sheet[Role].docx.


## Example

Here is a brief example of the data format in the Excel file:

- CS0514019 55 65 62 70 57 72 60
- CS0514020 57 75 42 60 67 52 70
- CS0514021 65 75 72 50 67 62 50


## Contributing

Feel free to submit issues or pull requests if you have suggestions or improvements.
