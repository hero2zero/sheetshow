#!/usr/bin/env python3
"""
SheetShow - Search and Export Excel/Text Files
A Python script that searches for a given search term across Excel and text files
and provides the option to save the results to a new spreadsheet/workbook.

Version: 1.1
"""

import os
import sys
import argparse
from typing import List, Dict, Any
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from tqdm import tqdm


class SearchResults:
    """Container for search results with Excel export functionality."""

    def __init__(self):
        self.results = []
        self.search_terms = []
        self.search_location = ""

    def add_result(self, file_path: str, line_number: int, line_content: str, matched_term: str = None, sheet_name: str = None, column_name: str = None, full_row_data: dict = None):
        """Add a search result."""
        result = {
            'file_path': file_path,
            'line_number': line_number,
            'line_content': line_content.strip()
        }
        if matched_term:
            result['matched_term'] = matched_term
        if sheet_name:
            result['sheet_name'] = sheet_name
        if column_name:
            result['column_name'] = column_name
        if full_row_data:
            result['full_row_data'] = full_row_data
        self.results.append(result)

    def save_to_excel(self, output_file: str = None):
        """Save search results to Excel workbook."""
        if not self.results:
            print("No results to save.")
            return

        if output_file is None:
            safe_terms = "_".join("".join(c for c in term if c.isalnum() or c in (' ', '_')).strip().replace(' ', '_') for term in self.search_terms)
            output_file = f"search_results_{safe_terms[:50]}.xlsx"

        # Prepare data for DataFrame
        export_data = []
        for result in self.results:
            row_data = {
                'file_path': result['file_path'],
                'line_number': result['line_number'],
                'line_content': result['line_content']
            }

            if 'matched_term' in result:
                row_data['matched_term'] = result['matched_term']
            if 'sheet_name' in result:
                row_data['sheet_name'] = result['sheet_name']
            if 'column_name' in result:
                row_data['column_name'] = result['column_name']

            # Add full row data if available
            if 'full_row_data' in result:
                for col_name, col_value in result['full_row_data'].items():
                    # Prefix with 'source_' to distinguish from search result columns
                    clean_col_name = f"source_{col_name}" if col_name not in ['file_path', 'line_number', 'line_content', 'sheet_name', 'column_name'] else f"orig_{col_name}"
                    row_data[clean_col_name] = col_value

            export_data.append(row_data)

        # Create DataFrame
        df = pd.DataFrame(export_data)

        # Create Excel workbook with formatting
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Search Results', index=False)

            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Search Results']

            # Style headers
            header_font = Font(bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')

            for cell in worksheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')

            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

            # Add summary sheet
            summary_data = {
                'Search Terms': [', '.join(self.search_terms)],
                'Search Location': [self.search_location],
                'Total Results': [len(self.results)],
                'Unique Files': [len(set(r['file_path'] for r in self.results))]
            }

            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Style summary sheet
            summary_ws = writer.sheets['Summary']
            for cell in summary_ws[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center')

            for column in summary_ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                summary_ws.column_dimensions[column_letter].width = max_length + 2

        print(f"Results saved to: {output_file}")
        return output_file


def search_excel_file(search_terms: List[str], file_path: Path, results: SearchResults):
    """
    Search for terms in an Excel file across all sheets.

    Args:
        search_terms: List of terms to search for
        file_path: Path to the Excel file
        results: SearchResults object to add matches to
    """
    try:
        # Read all sheets from the Excel file
        xl = pd.ExcelFile(file_path)

        # Search each sheet for the terms
        for sheet_name in xl.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                # Convert all columns to string and search for the terms
                for col in df.columns:
                    if df[col].dtype == 'object' or pd.api.types.is_string_dtype(df[col]):
                        # Search for each term in this column
                        for search_term in search_terms:
                            matches = df[df[col].astype(str).str.contains(search_term, case=False, na=False)]
                            if not matches.empty:
                                for idx, row in matches.iterrows():
                                    # Convert entire row to dictionary for full row data
                                    full_row_data = {}
                                    for column in df.columns:
                                        full_row_data[str(column)] = str(row[column]) if pd.notna(row[column]) else ""

                                    results.add_result(
                                        str(file_path.name),
                                        idx + 2,  # +2 because pandas is 0-indexed and Excel starts at 1, plus header
                                        str(row[col]),
                                        search_term,
                                        sheet_name,
                                        str(col),
                                        full_row_data
                                    )
            except Exception as e:
                print(f"Warning: Could not read sheet '{sheet_name}' in {file_path}: {e}")

    except Exception as e:
        print(f"Warning: Could not read Excel file {file_path}: {e}")


def search_files(search_terms: List[str], search_path: str = ".", file_extensions: List[str] = None) -> SearchResults:
    """
    Search for terms in files within the specified path.

    Args:
        search_terms: List of terms to search for
        search_path: Path to search in (default: current directory)
        file_extensions: List of file extensions to search (default: common text files and Excel)

    Returns:
        SearchResults object containing all matches
    """
    if file_extensions is None:
        file_extensions = ['.txt', '.py', '.js', '.html', '.css', '.md', '.json', '.xml', '.csv', '.xlsx', '.xls']

    results = SearchResults()
    results.search_terms = search_terms
    results.search_location = os.path.abspath(search_path)

    search_path = Path(search_path)

    if not search_path.exists():
        print(f"Error: Path '{search_path}' does not exist.")
        return results

    print(f"Searching for {len(search_terms)} term(s) in {search_path}...")

    total_files = 0
    matching_files = 0

    # Collect all files first for progress bar
    if search_path.is_file():
        files_to_search = [search_path] if search_path.suffix.lower() in file_extensions else []
    else:
        files_to_search = [f for f in search_path.rglob('*') if f.is_file() and f.suffix.lower() in file_extensions]

    total_files = len(files_to_search)

    # Search through files with progress bar
    for file_path in tqdm(files_to_search, desc="Searching files", unit="file"):
        initial_results_count = len(results.results)

        if file_path.suffix.lower() in ['.xlsx', '.xls']:
            # Handle Excel file
            search_excel_file(search_terms, file_path, results)
        else:
            # Handle text file
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line_number, line in enumerate(f, 1):
                        for search_term in search_terms:
                            if search_term.lower() in line.lower():
                                results.add_result(
                                    str(file_path.name) if search_path.is_file() else str(file_path.relative_to(search_path)),
                                    line_number,
                                    line,
                                    search_term
                                )
            except Exception as e:
                tqdm.write(f"Warning: Could not read file {file_path}: {e}")

        if len(results.results) > initial_results_count:
            matching_files += 1

    print(f"Search complete. Found {len(results.results)} matches in {matching_files} files out of {total_files} files searched.")
    return results


def display_results(results: SearchResults, max_display: int = 20):
    """Display search results in a formatted way."""
    if not results.results:
        print("No results found.")
        return

    print(f"\nSearch Results for: {', '.join(repr(t) for t in results.search_terms)}")
    print("=" * 60)

    displayed = 0
    for result in results.results[:max_display]:
        print(f"File: {result['file_path']}")
        if 'matched_term' in result:
            print(f"Matched term: '{result['matched_term']}'")
        if 'sheet_name' in result and 'column_name' in result:
            print(f"Sheet: {result['sheet_name']}, Row {result['line_number']}, Column: {result['column_name']}")
            print(f"Value: {result['line_content']}")
        else:
            print(f"Line {result['line_number']}: {result['line_content']}")
        print("-" * 40)
        displayed += 1

    if len(results.results) > max_display:
        print(f"... and {len(results.results) - max_display} more results")

    print(f"\nTotal: {len(results.results)} matches found")


def main():
    """Main function to handle command line interface."""
    parser = argparse.ArgumentParser(
        description='Search for text in files and optionally export results to Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sheetshow.py "function" --path ./src
  python sheetshow.py "TODO" "FIXME" --file script.py
  python sheetshow.py "factor" --file "Hosting Services IP Networks.xlsx"
  python sheetshow.py "error" "warning" --path /home/user/project --output my_results.xlsx
        """
    )

    parser.add_argument('search_terms', nargs='+', help='One or more text terms to search for')

    # Create mutually exclusive group for file or path (at least one required)
    target_group = parser.add_mutually_exclusive_group(required=True)
    target_group.add_argument('--path', '-p', help='Path/directory to search in')
    target_group.add_argument('--file', '-f', help='Specific file to search in')

    parser.add_argument('--extensions', '-e', nargs='+',
                       help='File extensions to search (default: .txt .py .js .html .css .md .json .xml .csv)')
    parser.add_argument('--export', '-x', action='store_true', help='Automatically export to Excel')
    parser.add_argument('--output', '-o', help='Output Excel file name (optional)')
    parser.add_argument('--max-display', '-m', type=int, default=20,
                       help='Maximum number of results to display (default: 20)')

    args = parser.parse_args()

    # Determine search target
    search_target = args.path if args.path else args.file

    # Perform search
    results = search_files(args.search_terms, search_target, args.extensions)

    # Display results
    display_results(results, args.max_display)

    # Handle Excel export
    if args.export or args.output:
        results.save_to_excel(args.output)
    elif results.results:
        # Ask user if they want to save to Excel
        response = input("\nWould you like to save these results to Excel? (y/n): ").strip().lower()
        if response in ['y', 'yes']:
            output_file = input("Enter output filename (press Enter for auto-generated name): ").strip()
            if not output_file:
                output_file = None
            results.save_to_excel(output_file)


if __name__ == "__main__":
    main()