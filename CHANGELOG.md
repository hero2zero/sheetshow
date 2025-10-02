# Changelog

All notable changes to SheetShow will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1] - 2025-10-02

### Added
- **Multi-term search**: Search for multiple terms simultaneously by providing multiple arguments
  - Example: `python sheetshow.py "TODO" "FIXME" "HACK" --path ./src`
  - Each result now tracks which term was matched via `matched_term` field
  - Excel export includes `matched_term` column to identify which search term triggered each match
- **Progress bar**: Visual feedback during file search using tqdm library
  - Shows real-time progress with file count and processing speed
  - Progress bar displays: "Searching files" with completion percentage and file/sec metrics
  - Uses `tqdm.write()` for warnings to avoid interfering with progress display

### Changed
- CLI argument `search_term` renamed to `search_terms` and now accepts one or more values (`nargs='+'`)
- `SearchResults.search_term` changed to `SearchResults.search_terms` (now stores a list)
- `search_files()` function signature updated to accept `List[str]` instead of `str`
- `search_excel_file()` function signature updated to accept `List[str]` instead of `str`
- Display output now shows all search terms in results header
- Auto-generated output filenames now incorporate all search terms (truncated to 50 chars)
- Summary sheet now displays "Search Terms" (plural) with comma-separated list

### Dependencies
- Added `tqdm` for progress bar functionality

## [1.0] - Initial Release

### Added
- Excel file search support (.xlsx, .xls) across multiple sheets
- Text file search support (.txt, .py, .js, .html, .css, .md, .json, .xml, .csv)
- Complete row context capture for Excel matches
- Interactive CLI with flexible options
- Rich Excel export with formatted workbooks
- Summary sheet with search statistics
- Customizable file extension filtering
- Option to search specific file or entire directory
- Auto-generated or custom output filenames
- Configurable display limit for results
