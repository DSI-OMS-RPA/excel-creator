
# ExcelCreator

**ExcelCreator** is a Python class that helps you create and manipulate Excel workbooks using the `openpyxl` library. It simplifies common tasks like adding headers, writing data rows, applying formatting, creating charts, adding data validation, and more.

## Features

- **Create and Save Excel Files**: Easily create and save Excel files (`.xlsx`).
- **Add Headers and Rows**: Insert custom headers and rows at any position.
- **Auto-Numbering**: Automatically number rows in a specified column.
- **Conditional Formatting**: Apply built-in or custom conditional formatting.
- **Charts**: Create bar, line, pie, and scatter charts with various customization options.
- **Data Validation**: Add drop-down lists and other data validation types to cells.
- **Merge Cells**: Merge a range of cells.
- **Freeze Panes**: Freeze headers or specific columns/rows.
- **Data Import**: Import data from CSV or JSON files.
- **Password Protection**: Protect sheets with a password.
- **Insert Images and Hyperlinks**: Add images and hyperlinks to cells.
- **Zebra Striping**: Apply alternating row colors for readability.

## Requirements

- Python 3.x
- `openpyxl` library

You can install the required library using:

```bash
pip install openpyxl
```

## Usage

### Example Code

```python
from excel_creator import ExcelCreator
from openpyxl.styles import Font

# Create an ExcelCreator instance
file_name = "example.xlsx"
excel_creator = ExcelCreator(file_name, header_font=Font(bold=True, color="000000"))

# Set sheet name
excel_creator.set_sheet_name("Report")

# Define and add headers
headers = ["ID", "Name", "Date", "Amount"]
excel_creator.add_headers(headers, start_row=1)

# Add rows of data
data = [
    [1, "Alice", "2024-01-01", 100],
    [2, "Bob", "2024-01-02", 200],
]
for i, row in enumerate(data, start=2):
    excel_creator.add_row(row, start_row=i)

# Auto-size columns and save
excel_creator.set_column_widths(auto_size=True)
excel_creator.save()
```

## Features in Detail

### Adding Headers
You can add headers starting from a specific row:

```python
headers = ["ID", "Name", "Date", "Amount"]
excel_creator.add_headers(headers, start_row=1)
```

### Adding Rows
Add rows of data with optional font styling:

```python
row_data = [1, "Alice", "2024-01-01", 100]
excel_creator.add_row(row_data, start_row=2)
```

### Creating Charts
Create a bar chart from data in your sheet:

```python
excel_creator.create_chart(
    min_col=2, min_row=2, max_col=5, max_row=10,
    chart_type="bar", title="Sales Data",
    x_axis_title="Products", y_axis_title="Sales",
    position="H10", include_legend=True, show_data_labels=True
)
```

### Importing Data from CSV or JSON
Import data directly from CSV or JSON files:

```python
excel_creator.import_from_csv("data.csv", start_row=2)
```

### Data Validation
Add a drop-down list (data validation) to a range of cells:

```python
payment_methods = '"Card,Cash"'
excel_creator.add_data_validation("D2:D10", validation_type="list", formula1=payment_methods)
```

## Saving the File

After performing your operations, save the Excel file:

```python
excel_creator.save()
```

## License

This script is free to use and modify.
