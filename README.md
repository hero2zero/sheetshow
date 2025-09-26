# SheetShow

A Python CLI tool that searches Excel and text files while capturing complete row context for Excel matches.

## Features

- **Excel File Support**: Search across multiple sheets in .xlsx/.xls files with full row context
- **Text File Support**: Search traditional text files (.txt, .py, .js, .html, .css, .md, .json, .xml, .csv)
- **Complete Row Context**: For Excel matches, capture and export all column data from the matching row
- **Interactive CLI**: User-friendly command-line interface with flexible options
- **Rich Export**: Save results to formatted Excel workbooks with detailed metadata
- **Multi-Sheet Search**: Automatically searches all sheets within Excel workbooks
- **Flexible Filtering**: Customizable file extension filtering for targeted searches

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

**Note:** Either `--path` or `--file` is required for all searches.

### Search in Excel file (with full row context):
```bash
python sheetshow.py "factor" --file "Hosting Services IP Networks.xlsx" --export
```

### Search in specific directory:
```bash
python sheetshow.py "function" --path ./src
```

### Search in specific text file:
```bash
python sheetshow.py "TODO" --file script.py
```

### Auto-export to Excel:
```bash
python sheetshow.py "error" --path . --export
```

### Specify file extensions (directory search):
```bash
python sheetshow.py "error" --path ./src --extensions .py .js .ts .xlsx
```

### Custom output file:
```bash
python sheetshow.py "bug" --path /home/user/project --output my_results.xlsx
```

## Command Line Options

- `search_term`: The text to search for (required)
- `--path, -p`: Directory to search in (mutually exclusive with --file, one required)
- `--file, -f`: Specific file to search in (mutually exclusive with --path, one required)
- `--extensions, -e`: File extensions to include in search (applies to directory searches)
- `--export, -x`: Automatically export results to Excel
- `--output, -o`: Specify output Excel filename
- `--max-display, -m`: Maximum results to display on screen (default: 20)

## Output Format

The Excel output includes:
- **Search Results** sheet: Detailed list of all matches with:
  - Core search info: file path, line number, matching content, sheet name, column name
  - Complete row context: All columns from the source row (prefixed with `source_`)
- **Summary** sheet: Search statistics and metadata
- Formatted headers and auto-adjusted column widths

### Excel Search Results

When searching Excel files, SheetShow provides complete row context:
```
File: Hosting Services IP Networks.xlsx
Sheet: Main, Row 18, Column: Unnamed: 2
Found: Factor Hosting
Full Row: IP=10.50.0.0, Mask=16.0, Description=Factor Hosting
```

## GitHub Repository Description

**SheetShow - Excel & Text File Search with Full Context**

A Python CLI tool that searches Excel and text files while capturing complete row context for Excel matches.