# size-chart-converter

## Description
The `size-chart-converter` script is a Python tool that can convert an Excel file containing size chart data, using start and end markers, into an HTML format. It can process multiple size charts within a single Excel file and generate the corresponding HTML output for each chart.

## Features
- Reads an Excel file (`.xlsx` or `.xls`) with size chart data
- Identifies individual size chart blocks using "start" and "end" markers
- Extracts the SKU, logo, and chart title from the SKU row
- Converts the table data between the "start" and "end" markers into an HTML table
- Preserves bold formatting from the original Excel file by rendering bold cells as `<th>` elements
- Generates a consolidated output file in Excel format, containing the HTML output for all processed size charts
- Provides a Streamlit-based user interface for easy file upload and conversion

## Prerequisites/Dependencies
- Python 3.x
- The following Python libraries:
  - `streamlit`
  - `pandas`
  - `openpyxl` (for reading cell styling from `.xlsx` files)

## How to Use/Run
1. Ensure you have Python 3.x installed on your system.
2. Install the required dependencies by running the following command in your terminal or command prompt:
   ```
   pip install streamlit pandas openpyxl
   ```
3. Save the Python script (e.g., `size-chart-converter.py`) to your desired location.
4. Run the script using the following command:
   ```
   streamlit run size-chart-converter.py
   ```
5. In the Streamlit web application, click the "Upload an Excel file with size chart data (using start/end markers)" button and select the appropriate Excel file.
6. The script will process the file and generate the HTML output for each size chart. The output will be displayed in the web application and can also be downloaded as an Excel file.

## Input Format
The input Excel file should be structured as follows:

1. **SKU Row**:
   - **Column A**: The SKU number (e.g., `SKU12`)
   - **Column B**: The full URL for the brand logo image (e.g., `https://.../logo.png`)
   - **Column C**: The title of the size chart (e.g., `XS Scuba Phenom Fins Size Chart`)
2. **'start' Marker Row**:
   - Some rows below the SKU row
   - **Column A**: Must contain the word `start` (case-insensitive)
3. **Table Data Rows**:
   - These are the rows strictly *between* the 'start' marker row and the 'end' marker row
   - Data for the table begins from **Column B** onwards in these rows
   - The **first data row** (immediately after 'start' row) is treated as the table header (`<thead>`). Its cells (from Col B onwards) are the size headers (e.g., "Small", "Medium")
   - **Subsequent data rows** (up to the row before 'end') are treated as table body rows (`<tbody>`)
   - **Cell Styling**:
     - For `.xlsx` files, if a cell within this table data (headers or body, from Col B onwards) is **bold** in Excel, it will be rendered as `<th>`
     - Otherwise, it will be rendered as `<td>`
     - Empty cells will be rendered as `<td> </td>` (or `<th> </th>` if an empty cell is bold)
4. **'end' Marker Row**:
   - Some rows below the last data row for the chart
   - **Column A**: Must contain the word `end` (case-insensitive)

## License
License to be determined.