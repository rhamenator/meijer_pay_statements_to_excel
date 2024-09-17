# Meijer Pay Statements to Excel Converter

This Python script automates the extraction of data from Meijer hourly worker pay statements (in PDF format) and organizes it into a structured Excel workbook, along with JSON, XML, and CSV files.

## Important Note

**This script is specifically designed to work with Meijer hourly worker pay statements.** It may not function correctly with pay statements from other employers or even other types of Meijer pay statements (e.g., salaried employees).

## Features

**Data Extraction:** Accurately extracts key information from Meijer pay statements, including:
    - Employee Identification
    - Pay Period Details
    - Earnings Breakdown
    - Taxes and Deductions
    - Employer Benefits
    - Payment Information
**Multiple Output Formats:** Generates:
    - Excel Workbook (xlsx) with well-formatted tables and styling
    - JSON file (json) for easy data exchange
    - XML file (xml) for structured data representation
    - CSV file (csv) for basic spreadsheet compatibility
**User-Friendly Interface:** Provides a simple graphical user interface (GUI) for selecting input and output files.
**Command-Line Support:** Can also be run from the command line for automation or scripting.

## Requirements

- Python 3.x
- The following Python libraries (install using `pip install -r requirements.txt`):

    ``` Text
    pandas
    numpy
    dicttoxml
    PyMuPDF
    keyboard
    tkinterdnd2
    openpyxl
    tqdm
    ```

## Usage

1. Open your terminal or command prompt.
2. Navigate to the directory containing the script.
3. Next, type (you only have to do this the first time you run the program)

    ``` Terminal
    pip install -r requirements.txt
    ```

4. Run the script using:

    ``` Terminal
    python pay_statements_to_excel.py <input_pdf_path> <output_excel_path>
    ```

5. If you want to be prompted to enter the name of the input PDF file and the name of the output Excel file, just leave the parameters blank:

    ``` Terminal
    python pay_statements_to_excel.py
    ```
