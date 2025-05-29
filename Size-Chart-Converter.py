import streamlit as st
import pandas as pd
import io
import openpyxl # For reading cell styling

st.set_page_config(layout="wide")
st.title("Excel Size Chart to HTML Converter (Start/End Markers - v5)")

# --- FUNCTION DEFINITION ---
def generate_html_for_chart_with_start_end(
    df_full_sheet, # The complete DataFrame of the Excel sheet
    sku_row_index, # 0-based index of the SKU row in the full sheet
    uploaded_file_obj_for_styling
):
    # --- Extract SKU, Logo, Title from the SKU row ---
    try:
        sku_value = str(df_full_sheet.iloc[sku_row_index, 0]).strip() # Should match calling SKU
        logo_src = str(df_full_sheet.iloc[sku_row_index, 1])
        chart_title = str(df_full_sheet.iloc[sku_row_index, 2])
        if pd.isna(logo_src) or str(logo_src).strip().lower() == 'nan': logo_src = ""
        if pd.isna(chart_title) or str(chart_title).strip().lower() == 'nan': chart_title = "Size Chart"
    except IndexError:
        return f"<p>Error for SKU at sheet row {sku_row_index + 1}: Could not read SKU/logo/title. Ensure Columns A, B, C exist for the SKU row.</p>"

    # --- Find "start" and "end" markers for this SKU ---
    start_marker_row_idx = -1
    end_marker_row_idx = -1

    # Scan for "start" from the row after the SKU row
    for i in range(sku_row_index + 1, len(df_full_sheet)):
        marker_val = df_full_sheet.iloc[i, 0] # Column A for markers
        if pd.notna(marker_val) and str(marker_val).strip().lower() == "start":
            start_marker_row_idx = i
            break
    
    if start_marker_row_idx == -1:
        return f"<p>Error for chart '{chart_title}' (SKU: {sku_value}): 'start' marker not found below SKU row.</p>"

    # Scan for "end" from the row after the "start" marker
    for i in range(start_marker_row_idx + 1, len(df_full_sheet)):
        marker_val = df_full_sheet.iloc[i, 0] # Column A for markers
        if pd.notna(marker_val) and str(marker_val).strip().lower() == "end":
            end_marker_row_idx = i
            break
            
    if end_marker_row_idx == -1:
        return f"<p>Error for chart '{chart_title}' (SKU: {sku_value}): 'end' marker not found after 'start' marker.</p>"

    # --- Define the actual table data boundaries (exclusive of start/end marker rows) ---
    table_data_first_row_abs = start_marker_row_idx + 1
    table_data_last_row_abs = end_marker_row_idx - 1

    if table_data_first_row_abs > table_data_last_row_abs:
        return f"<p>Error for chart '{chart_title}' (SKU: {sku_value}): No data rows found between 'start' and 'end' markers (or 'end' is before/on 'start' data row).</p>"

    # --- Slice the table data block from the full sheet ---
    # Table data starts from Column B (pandas index 1)
    try:
        # +1 because iloc is exclusive for the end row index
        df_table_block = df_full_sheet.iloc[table_data_first_row_abs : table_data_last_row_abs + 1, 1:].reset_index(drop=True)
    except IndexError:
        return f"<p>Error for chart '{chart_title}' (SKU: {sku_value}): Problem slicing the table block. Check sheet structure.</p>"

    if df_table_block.empty:
        return f"<p>Error for chart '{chart_title}' (SKU: {sku_value}): Table block between 'start' and 'end' is empty or not structured correctly starting Column B.</p>"
    
    # --- Prepare openpyxl workbook for styling if it's an .xlsx file ---
    workbook_for_styling = None
    sheet_for_styling = None
    is_xlsx = False
    if uploaded_file_obj_for_styling and hasattr(uploaded_file_obj_for_styling, 'name'):
        file_name_lower = uploaded_file_obj_for_styling.name.lower()
        if file_name_lower.endswith('.xlsx'):
            is_xlsx = True
            try:
                uploaded_file_obj_for_styling.seek(0)
                workbook_for_styling = openpyxl.load_workbook(uploaded_file_obj_for_styling, read_only=True, data_only=False)
                sheet_for_styling = workbook_for_styling.active
            except Exception as e:
                st.warning(f"Could not load .xlsx for styling for chart '{chart_title}': {e}. Bold detection disabled.")
                workbook_for_styling = None # Disable styling if load fails

    # --- Generate HTML for the table block ---
    html_table_rows = []
    num_table_block_rows = len(df_table_block)
    num_table_block_cols = len(df_table_block.columns)

    for r_idx_in_block in range(num_table_block_rows): # Iterate rows of the sliced table block
        row_cells_html = []
        for c_idx_in_block in range(num_table_block_cols): # Iterate columns of the sliced table block
            cell_value_raw = df_table_block.iloc[r_idx_in_block, c_idx_in_block]
            
            cell_value_str = " " # Default for empty
            if pd.notna(cell_value_raw):
                temp_str = str(cell_value_raw).strip()
                if temp_str != "" and temp_str.lower() != "nan":
                    cell_value_str = temp_str.replace('\n', ' ') # Clean up newlines

            is_bold = False # Default, will be <th> if bold
            if is_xlsx and sheet_for_styling:
                # Calculate absolute row/col in the original sheet for openpyxl
                # df_table_block starts from col B of sheet, so its col 0 is sheet col 1 (B)
                abs_sheet_col_for_styling = (1 + c_idx_in_block) + 1 # 0-based sheet_col + 1 for 1-based openpyxl
                abs_sheet_row_for_styling = table_data_first_row_abs + r_idx_in_block + 1 # 0-based sheet_row + 1 for 1-based openpyxl
                
                try:
                    cell_obj = sheet_for_styling.cell(row=abs_sheet_row_for_styling, column=abs_sheet_col_for_styling)
                    if cell_obj.font and cell_obj.font.bold:
                        is_bold = True
                except Exception as e_cell: # Catch potential errors if cell doesn't exist (e.g. ragged table)
                    # st.warning(f"Styling check error for cell ({abs_sheet_row_for_styling},{abs_sheet_col_for_styling}): {e_cell}")
                    pass # Keep is_bold as False

            tag = "th" if is_bold else "td"
            row_cells_html.append(f'<{tag}>{cell_value_str}</{tag}>')
        
        html_table_rows.append(f'<tr>{"".join(row_cells_html)}</tr>')

    # Construct the full table HTML
    # First row of table_block is thead, rest is tbody
    thead_html = ""
    tbody_html = ""

    if html_table_rows:
        thead_html = f'<thead>{html_table_rows[0]}</thead>'
        if len(html_table_rows) > 1:
            tbody_html = f'<tbody>{"".join(html_table_rows[1:])}</tbody>'
        else: # Only one row means it's all header, no body
            tbody_html = '<tbody></tbody>' # Or omit tbody entirely if preferred
    
    full_html = f'''<div id="sizeChartContainer">
<div class="sizeChartHeader">
<img class="sizeChartBrandLogo" src="{logo_src}"/>
<strong>{chart_title}</strong>
<table class="sizeChart" cellspacing="0" cellpadding="0" width="100%">
{thead_html}
{tbody_html}
</table>
</div>
</div>'''
    return full_html

# --- Main Streamlit App ---
uploaded_file = st.file_uploader("Upload an Excel file with size chart data (using start/end markers)", type=["xlsx", "xls"])

if uploaded_file is not None:
    all_results = []
    try:
        df_full_sheet = pd.read_excel(uploaded_file, header=None, sheet_name=0)
        st.write(f"Processing file: {uploaded_file.name}. Found {df_full_sheet.shape[0]} rows.")

        # Find all rows that contain SKUs in Column A
        sku_row_indices = []
        for idx, row_val in df_full_sheet.iloc[:, 0].items():
            # Check if it's not 'start' or 'end' and not NaN, to identify SKU rows
            if pd.notna(row_val):
                val_str = str(row_val).strip().lower()
                if val_str not in ["start", "end"] and val_str != "":
                    sku_row_indices.append(idx)
        
        if not sku_row_indices:
            st.warning("No SKU rows found in Column A (excluding 'start'/'end' markers).")
        else:
            st.write(f"Found {len(sku_row_indices)} potential SKU blocks.")
            for sku_idx in sku_row_indices:
                current_sku_val = str(df_full_sheet.iloc[sku_idx, 0]).strip()
                st.markdown(f"--- \n**Processing SKU: {current_sku_val} (Original sheet row {sku_idx + 1})**")

                html_output_for_sku = generate_html_for_chart_with_start_end(
                    df_full_sheet,
                    sku_idx,
                    uploaded_file # Pass the file object for openpyxl styling
                )
                all_results.append({'SKU': current_sku_val, 'HTML_Output': html_output_for_sku})

                st.subheader(f"HTML Preview for {current_sku_val}:")
                if "<p>Error:" not in html_output_for_sku:
                    st.markdown(html_output_for_sku, unsafe_allow_html=True)
                else:
                    st.error(html_output_for_sku)

        if not all_results:
            st.info("No chart data could be successfully processed.")
        else:
            st.markdown("--- \n## All Processed Charts Output")
            df_output = pd.DataFrame(all_results)
            st.dataframe(df_output)

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_output.to_excel(writer, index=False, sheet_name="SizeChartHTML")
            excel_buffer.seek(0)

            output_filename = "size_charts_output.xlsx"
            if hasattr(uploaded_file, 'name') and uploaded_file.name:
                base_name = uploaded_file.name.rsplit('.', 1)[0]
                output_filename = f"{base_name}_output.xlsx"

            st.download_button(
                label="Download All Output as Excel",
                data=excel_buffer.getvalue(),
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"A critical error occurred during file processing: {e}")
        st.exception(e)
else:
    st.info("Please upload an Excel file.")

st.markdown("---")
st.markdown("""
### Instructions for Input Excel File:
The Excel file can contain multiple size charts. Each chart block starts with a SKU row, followed by a 'start' marker, the table data, and an 'end' marker.

**For each chart block:**
1.  **SKU Row**:
    *   **Column A**: The SKU number (e.g., `SKU12`).
    *   **Column B**: The full URL for the brand logo image (e.g., `https://.../logo.png`).
    *   **Column C**: The title of the size chart (e.g., `XS Scuba Phenom Fins Size Chart`).

2.  **'start' Marker Row**:
    *   Some rows below the SKU row.
    *   **Column A**: Must contain the word `start` (case-insensitive).

3.  **Table Data Rows**:
    *   These are the rows strictly *between* the 'start' marker row and the 'end' marker row.
    *   Data for the table begins from **Column B** onwards in these rows.
    *   The **first data row** (immediately after 'start' row) is treated as the table header (`<thead>`). Its cells (from Col B onwards) are the size headers (e.g., "Small", "Medium").
    *   **Subsequent data rows** (up to the row before 'end') are treated as table body rows (`<tbody>`).
    *   **Cell Styling**:
        *   For `.xlsx` files, if a cell within this table data (headers or body, from Col B onwards) is **bold** in Excel, it will be rendered as `<th>`.
        *   Otherwise, it will be rendered as `<td>`.
        *   Empty cells will be rendered as `<td> </td>` (or `<th> </th>` if an empty cell is bold).

4.  **'end' Marker Row**:
    *   Some rows below the last data row for the chart.
    *   **Column A**: Must contain the word `end` (case-insensitive).

**Sheet Structure Example (using start/end markers):**""")