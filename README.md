# Excel Data Processing and Formatting Script

This script is designed to process and format an Excel file and output that data to another Excel file. It performs a series of operations including column deletion, column insertion, sheet creation, data copying, and applying various formatting styles.

## Dependencies

- pandas
- openpyxl

## Installation

Ensure you have Python installed on your system. Then, install the required dependencies using pip:

bash pip install pandas openpyxl


## Usage

1. Place the original Excel file in the same directory as the script.
2. Run the script using Python:

bash python script_name.py


Replace `script_name.py` with the actual name of your script file.

## Script Overview

### Loading and Preprocessing the Data

The script starts by loading the Excel file using `pandas.read_excel()`.

- It then deletes specified columns from the DataFrame.
- New columns are inserted at specific positions in the DataFrame.
- The modified DataFrame is saved to a new Excel file.

### Formatting and Organizing the Excel File

The script loads the newly created Excel file using `openpyxl.load_workbook()`.

- The first sheet is renamed to "Skills".
- New sheets are created and positioned after the first sheet.
- Data is copied from the "Skills" sheet to the newly created sheets based on specified column mappings.

### Applying Formulas

Formulas are applied to the "Skills" sheet to check for the presence of data in specific columns and mark them with "X" if data is present.

### Styling

- The script applies a slightly darker blue fill pattern to all cells in row 1 of each sheet.
- It also applies a very light shade of blue fill pattern to every other row starting from row 3, moving downwards.
- All data in the cells is centered.
- Column widths are adjusted to fit the content.

### Finalizing

The workbook is saved.

## Output

The output is an Excel file with the following characteristics:

- The first sheet is named "Skills" and contains the processed data.
- Additional sheets are created for different categories of skills, each containing relevant data copied from the "Skills" sheet.
- Various formatting styles are applied to enhance readability and organization.

## Note

This script is designed to work with a specific structure of the input Excel file. Adjustments may be necessary if the structure of the input file changes.
