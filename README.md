# Data Entry Form (Tkinter + Excel)

This is a simple desktop application built with **Python Tkinter** that allows users to enter personal information (Name, Email, Phone, Address, Notes) and save it into an **Excel file** (`submissions.xlsx`).

## Features
- GUI form built using **Tkinter**.
- Data validation for:
  - Name (letters only, min 2 characters)
  - Email (valid format, no duplicates)
  - Phone (10–15 digits)
  - Address (min 10 characters)
- Saves submissions into an Excel file automatically.
- Creates the Excel file with headers if it doesn’t exist.
- Option to clear form fields.
- Button to show where the Excel file is saved.

## Requirements
Install dependencies before running:

```bash
pip install pandas openpyxl
