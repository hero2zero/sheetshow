# SheetShow

A Python CLI tool that searches Excel and text files while capturing complete row context for Excel matches.

**Version: 1.1**

## Features

- **Multi-Term Search**: Search for multiple terms simultaneously in a single pass
- **Progress Bar**: Real-time visual feedback showing search progress and speed
- **Excel File Support**: Search across multiple sheets in .xlsx/.xls files with full row context
- **Text File Support**: Search traditional text files (.txt, .py, .js, .html, .css, .md, .json, .xml, .csv)
- **Complete Row Context**: For Excel matches, capture and export all column data from the matching row
- **Match Tracking**: Each result identifies which search term triggered the match
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

### Search for single term in Excel file (with full row context):
```bash
python sheetshow.py "factor" --file "Hosting Services IP Networks.xlsx" --export
```

### Search for multiple terms simultaneously:
```bash
python sheetshow.py "TODO" "FIXME" "HACK" --path ./src
```

### Search in specific directory:
```bash
python sheetshow.py "function" --path ./src
```

### Search in specific text file:
```bash
python sheetshow.py "TODO" --file script.py
```

### Search multiple terms with auto-export:
```bash
python sheetshow.py "error" "warning" "critical" --path . --export
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

- `search_terms`: One or more text terms to search for (required, space-separated)
- `--path, -p`: Directory to search in (mutually exclusive with --file, one required)
- `--file, -f`: Specific file to search in (mutually exclusive with --path, one required)
- `--extensions, -e`: File extensions to include in search (applies to directory searches)
- `--export, -x`: Automatically export results to Excel
- `--output, -o`: Specify output Excel filename
- `--max-display, -m`: Maximum results to display on screen (default: 20)

## Output Format

The Excel output includes:
- **Search Results** sheet: Detailed list of all matches with:
  - Core search info: file path, line number, matching content, matched term, sheet name, column name
  - Complete row context: All columns from the source row (prefixed with `source_`)
- **Summary** sheet: Search statistics and metadata including all search terms
- Formatted headers and auto-adjusted column widths

### Progress Display

During search, a progress bar shows:
```
Searching files: 45%|████████▌         | 123/275 [00:08<00:10, 15.2 file/s]
```

### Excel Search Results

When searching Excel files, SheetShow provides complete row context:
```
File: Hosting Services IP Networks.xlsx
Matched term: 'factor'
Sheet: Main, Row 18, Column: Unnamed: 2
Value: Factor Hosting
----------------------------------------
```

### Multi-Term Search Results

When searching for multiple terms, each result shows which term matched:
```
Search Results for: 'TODO', 'FIXME', 'HACK'
============================================================
File: src/utils.py
Matched term: 'TODO'
Line 42: # TODO: Refactor this function
----------------------------------------
File: src/main.py
Matched term: 'FIXME'
Line 156: # FIXME: Handle edge case
----------------------------------------
```

## GitHub Repository Description

**SheetShow - Excel & Text File Search with Full Context**

A Python CLI tool that searches Excel and text files while capturing complete row context for Excel matches.