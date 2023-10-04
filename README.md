# Problem

When an Excel user has many Excel sheets (20+), it can become challenging to search for a specific sheet, especially if the sheet names are long. This difficulty is exacerbated when trying to locate a reference inside a cell.

# Purpose

This script aims to address this issue by allowing Excel users to switch quickly between sheets (thanks to an autocomplete feature) and locate specific references with ease.

# Requirements

* `Windows`: This script works only on Windows (constraint of Excel)
* `Excel`: This script primarily uses Excel installed locally on your computer to assist you.

# Installation

To get started, follow these steps:

1. Install Python 3 (Tested on version 3.8.10).
2. During installation, ensure that you set up the Python environment variables correctly.
3. Execute `run.bat`, which will install the necessary dependencies for the script.

# Execution

To use the script, follow these instructions:

1. Run `app.py`.
2. Select the path to your Excel file.
3. Define the column that contains the reference values in the Excel file. This is the column that the script will search for reference values.
4. Choose a sheet name. Initially, it will open Excel, which may take some time to display hints for autocomplete.
5. Optionally, you can also specify the reference value if you want to locate a specific reference within the selected sheet.

