# excelDB - Simple Excel File Query and Export

excelDB is a lightweight web-based tool that allows you to easily query and export data from Excel files. With excelDB, you can upload an Excel file, search through the data, select specific columns and rows, and export the filtered data in either tab-separated, CSV, or Excel format.

## Features

- Upload and display Excel files
- Select and filter columns to display
- Search through the data
- Sort data by column
- Select a primary key field to highlight the column
- Handle blank sheets with appropriate message
- Copy displayed rows or selected rows to clipboard
- Export data in tab-separated, CSV, or Excel format
- Show only selected rows

## Getting Started

### Prerequisites

To use excelDB, you'll need a web browser with JavaScript enabled.

### Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/yourusername/excelDB.git
    ```
2. Navigate to the project directory:
    ```bash
    cd excelDB
    ```
3. Open `index.html` in your web browser.

### Usage

1. Open `index.html` in your preferred web browser.
2. Click the "Choose File" button to upload an Excel file.
3. Use the "Header Row" input to specify the header row (1-based index).
4. Use the "Primary Key" dropdown to select a primary key field.
5. Use the "Select columns to display" checkboxes to choose which columns to display.
6. Use the search bar to filter rows based on the search term.
7. Click the column headers to sort the data.
8. Use the "Copy Displayed Rows" button to copy the filtered and displayed rows to the clipboard.
9. Use the "Copy Selected Rows" button to copy the selected rows to the clipboard.
10. Use the "Export to Excel" button to export the displayed rows to an Excel file.
11. Use the "Show Only Selected" button to display only the selected rows.
12. Use the "Clear Search and Selection" button to reset the search and clear all selections.
13. Select the export format (Tab-separated or CSV) before copying.

## Versioning

excelDB uses [Semantic Versioning](https://semver.org/) for versioning.

| Version | Description                                                                                          | Date       |
|---------|------------------------------------------------------------------------------------------------------|------------|
| v0.1.0  | Initial commit: Added core functionality for uploading, displaying, and exporting Excel file data.   | 2024-06-21 |
| v0.1.1  | Added Excel export, merged clear buttons, fixed date display.                                        | 2024-06-25 |
| v0.1.2  | Fix for identifying sheets that are blank, added Primary Key feature.                                | 2024-06-25 |


## Acknowledgments

- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) for providing the library to read Excel files.
