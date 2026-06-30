# LSPR Data Analysis Software

A simple desktop data analysis prototype. It can load CSV and Excel files, then generate a basic automatic analysis report.

## Run

```powershell
pip install -r requirements.txt
python main.py
```

## Current Features

- Load `.csv`, `.xlsx`, and `.xls` files
- Validate CSV row counts and use a raw-line fallback when parsed rows are missing
- Automatically split one-column data when values contain common delimiters, preserving the widest row
- Drag column headers to select one or more columns
- Drag the `#` row header to select one or more rows
- Highlight selected rows and columns in blue
- Highlight generated `Average` columns in yellow
- Insert an average column after the selected column range, rounded to 3 decimal places
- Keep the first cell of each generated `Average` column editable
- Undo the last average calculation with the `Undo` button or `Ctrl+Z`
- Zoom the table and report with toolbar buttons or `Ctrl + mouse wheel`
- Display all rows directly for normal-sized datasets, with virtual scrolling only for very large files
- Analyze missing values, duplicate rows, numeric statistics, top text/category values, and numeric correlations
