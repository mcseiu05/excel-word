# Performance Review Document Generator

This script generates personalized performance review documents in `.docx` format based on data from an Excel file. It uses a pre-defined Word document template and fills in employee-specific information, creating a unique file for each entry.

## Table of Contents
- [Requirements](#requirements)
- [How It Works](#how-it-works)
- [File Structure](#file-structure)
- [Usage](#usage)
- [Example Output](#example-output)

## Requirements
This script requires the following Python packages:
- `openpyxl`: For loading data from Excel files.
- `docxtpl`: For rendering data into Word templates.
- `datetime`: For timestamping generated files.

Install these packages with:

 `pip install openpyxl docxtpl`

## How It Works
- **Load Employee Data**: Reads employee data from an Excel file (`performance_summary.xlsx`).
- **Load Template**: Loads a Word document template (`performance-review.docx`) with placeholders for data fields like employee name, score, and remarks.
- **Generate Documents**: Iterates through each row in the Excel file (excluding headers), fills the template with employee data, and saves a personalized `.docx` file.
- **Save to Output Directory**: Saves each document with a unique name (based on employee ID and timestamp) in the `generated_docs` folder.

## File Structure
- **`performance_summary.xlsx`**: Excel file containing employee performance data.
- **`performance-review.docx`**: Word template file with placeholders for employee data.
- **`generated_docs/`**: Directory where generated documents are saved.

## Usage
1. Place the Excel file (`performance_summary.xlsx`) and template (`performance-review.docx`) in the same directory as the script.
2. Run the script:
   
   `python main.py`
   
3. Generated documents will be saved in the `generated_docs` directory with unique filenames.

## Example Output
The generated documents will be saved with a filename structure: performance-review_<employee_id>_<timestamp>.docx


## Error Handling
If there’s an error loading the Excel file, an error message will be printed, and the program will terminate.

