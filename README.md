# Meijer Pay Statements to Excel Converter

This Python program is designed to streamline the process of converting Meijer pay statements from various formats into a structured Excel spreadsheet. It automates the extraction, processing, and organization of payroll data, making it easier for users to analyze and manage their financial records.

## Key Features

- **Automated Data Extraction**: The program efficiently extracts relevant data from pay statements using advanced pattern recognition and regular expressions.
  
- **Flexible Input Formats**: It supports a variety of input formats, thanks to libraries like `fitz` for PDF processing and `json` for handling structured data.

- **Excel Integration**: Using `pandas` and `openpyxl`, the program converts the extracted data into well-organized Excel spreadsheets, facilitating easy data analysis and reporting.

- **Customizable and Extensible**: The modular design allows for easy customization and extension to support additional features or new input formats.

- **User Interaction**: The program includes a simple GUI using `tkinter` for ease of use, and supports command-line operation for advanced users.

- **Robust Error Handling**: Comprehensive error handling ensures smooth operation and provides informative feedback to users, making it resilient to various input anomalies.

## Requirements

The program relies on several Python libraries, including:

- `pandas` and `numpy` for data manipulation
- `openpyxl` for Excel file handling
- `fitz` (PyMuPDF) for PDF processing
- `dicttoxml` for converting dictionaries to XML
- `keyboard`, `tkinter`, and `tqdm` for user interface and experience
- Standard libraries such as `subprocess`, `os`, and `sys` for system-level operations

Please refer to the `requirements.txt` file for the complete list of dependencies.

## Usage

To use this program, clone the repository and install the required dependencies:

```bash
pip install -r requirements.txt
```

Then, run the script:

```bash
python pay_statements_to_excel.py
```

The program will guide you through the process of selecting input files and generating the output Excel file.
