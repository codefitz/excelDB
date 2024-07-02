# ExcelDB - Simple Excel File Query and Export

ExcelDB is a web application that allows users to upload Excel files, query data, and export the results. It supports sorting, searching, and selecting specific columns for display. Additionally, users can export the displayed data to Excel or copy it to the clipboard.

## Features

- **File Upload**: Upload Excel files directly from your computer.
- **Sheet Selection**: Choose which sheet to load and display.
- **Column Selection**: Select specific columns to display in the table.
- **Primary Key Selection**: Designate a primary key for more efficient data handling.
- **Search Functionality**: Search through the data with an input field.
- **Sorting**: Sort data by clicking on the column headers.
- **Export Options**: Export displayed or selected data to Excel or copy it to the clipboard.
- **Display Options**: Choose to display all records, single records, or paginate through 10 records at a time.
- **Multi-Sheet Support**: Load and display data from multiple sheets within the same workbook.

## Getting Started

To get started with ExcelDB, follow these simple steps:

1. **Clone the Repository**:
    ```sh
    git clone https://github.com/yourusername/excelDB.git
    ```
2. **Navigate to the Project Directory**:
    ```sh
    cd excelDB
    ```
3. **Open `index.html` in Your Web Browser**:
    Simply open the `index.html` file in your favorite web browser to start using the application.

## Usage

### Uploading an Excel File

1. Click on the "Choose File" button to upload an Excel file.
2. Select the sheet you want to load from the dropdown menu.
3. Specify the header row (1-based index) if it differs from the default value (1).
4. Click the "Load Data" button to display the data.

### Selecting Columns

- Use the checkboxes under "Select columns to display" to choose which columns to display.
- Use the "Select All" and "Deselect All" buttons to quickly select or deselect all columns.

### Sorting Data

- Click on the column headers to sort the data in ascending or descending order.

### Searching Data

- Use the search input field to filter the displayed data based on your search query.

### Exporting Data

- Click "Export to Excel" to export the displayed data to an Excel file.
- Use the "Copy Displayed Rows" or "Copy Selected Rows" buttons to copy the data to your clipboard.

### Display Options

- Choose "Show All Records" to display all records.
- Choose "Show Single Record" to display one record at a time and use navigation buttons to browse through records.
- Choose "Show 10 Records at a Time" to paginate through records 10 at a time.

## Contributing

We welcome contributions! Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/YourFeature`).
3. Make your changes.
4. Commit your changes (`git commit -m 'Add some feature'`).
5. Push to the branch (`git push origin feature/YourFeature`).
6. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.


## Versioning

excelDB uses [Semantic Versioning](https://semver.org/) for versioning.

| Version | Description                                                                                          | Date       |
|---------|------------------------------------------------------------------------------------------------------|------------|
| v0.1.0  | Initial commit: Added core functionality for uploading, displaying, and exporting Excel file data.   | 2024-06-21 |
| v0.1.1  | Added Excel export, merged clear buttons, fixed date display.                                        | 2024-06-25 |
| v0.1.2  | Fix for identifying sheets that are blank, added Primary Key feature.                                | 2024-06-25 |
| v0.1.3  | Added ability to load second sheet for primary key lookup.<br>Added single record view and 10 rows view by default.<br>Bug fixes and optimisation. | 2024-07-02 |


## Acknowledgments

- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) for providing the library to read Excel files.
