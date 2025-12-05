import streamlit as st
import pandas as pd
import io
import random
from datetime import datetime
import pytz

st.title("Shipment Grouping Tool")
st.write(
    "Upload an Excel file. This tool will group rows based on the first 15 characters "
    "of Column C and separate shipments (A, B, C...) into alphabetical order, exporting "
    "one sheet per group."
)

uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded:
    # Read Excel with all columns as strings to preserve leading zeros and avoid scientific notation
    df = pd.read_excel(uploaded, dtype=str)

    # Access the third column (Excel column C) by position
    if len(df.columns) < 3:
        st.error("File needs at least 3 columns. Please check your file.")
        st.stop()
    
    # Save original data for the first tab
    df_original = df.copy()
    
    # Get the third column (index 2 = column C)
    third_column = df.iloc[:, 2]

    # --- Extract grouping keys ---
    df["group_15"] = third_column.astype(str).str[:15]      # First 15 characters
    df["shipment"] = third_column.astype(str).str[:16]      # First 16 characters (with A/B/C)

    # Sort alphabetically by shipment key
    df = df.sort_values(by=["group_15", "shipment"])

    # --- Create Excel output with multiple sheets ---
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book

    # Define cell formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'fg_color': '#4472C4',
        'font_color': 'white',
        'border': 1
    })
    
    # Red header format for Issues column
    red_header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#FF0000',  # Red background
        'font_color': 'white',
        'border': 1
    })
    
    # Text format for all data - preserves leading zeros and shows full numbers
    text_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '@',  # Format as text to preserve leading zeros and avoid scientific notation
        'locked': False  # Allow editing
    })
    
    # Locked text format for protected formula cells
    locked_text_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '@',
        'locked': True  # Protect from editing
    })
    
    # Bold text format for grand total
    bold_text_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bold': True,
        'num_format': '@'
    })
    
    # Color formats for team member assignments (unlocked for editing)
    orville_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFFFE0',  # Light yellow
        'num_format': '@',
        'locked': False
    })
    
    sunshine_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#ADD8E6',  # Light blue
        'num_format': '@',
        'locked': False
    })
    
    stephanie_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFDAB9',  # Light orange/peach
        'num_format': '@',
        'locked': False
    })
    
    paulo_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFB6C1',  # Light pink
        'num_format': '@',
        'locked': False
    })
    
    jb_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#90EE90',  # Light green
        'num_format': '@',
        'locked': False
    })
    
    # Red format for Issues column with white text
    red_issues_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FF6B6B',  # Red background
        'font_color': 'white',  # White text for readability
        'num_format': '@'
    })
    
    # Yellow format for "UPLOADED" status
    uploaded_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFFF00',  # Yellow background
        'num_format': '@',
        'locked': False
    })
    
    # Red format for "WITH ISSUE" status
    with_issue_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FF0000',  # Red background
        'font_color': 'white',  # White text for readability
        'num_format': '@',
        'locked': False
    })
    
    # --- Write original data to first sheet ---
    # Keep original formatting as close to input file as possible
    original_sheet_name = "Original Data"
    df_original.to_excel(writer, sheet_name=original_sheet_name, index=False)
    
    # Get the original data worksheet
    original_worksheet = writer.sheets[original_sheet_name]
    
    # Set black tab color for Original Data tab
    original_worksheet.set_tab_color('#000000')
    
    # Auto-fit columns for readability
    for col_num, col_name in enumerate(df_original.columns):
        max_width = len(str(col_name)) + 2
        for value in df_original.iloc[:, col_num]:
            try:
                cell_width = len(str(value)) + 2
                if cell_width > max_width:
                    max_width = cell_width
            except:
                pass
        original_worksheet.set_column(col_num, col_num, min(max_width, 50))
    
    # --- Create PO Summary tab (second tab) ---
    po_summary_sheet_name = "PO Summary"
    
    # Get unique PO numbers (15 characters from group_15)
    unique_pos = sorted(df["group_15"].unique())
    total_pos = len(unique_pos)
    
    # Team members - Orville gets lower priority for remainders
    team_members = ["Paulo", "JB", "Sunshine", "Stephanie", "Orville"]
    
    # Calculate base assignment per person
    base_per_person = total_pos // len(team_members)
    remainder = total_pos % len(team_members)
    
    # Create assignment list
    assignments = []
    for i, member in enumerate(team_members):
        # Give extra POs to Paulo, JB, Sunshine, Stephanie (not Orville) if there's remainder
        if i < remainder and member != "Orville":
            count = base_per_person + 1
        elif i < remainder and member == "Orville":
            # Give Orville's extra to the first members instead
            count = base_per_person
        else:
            count = base_per_person
        assignments.extend([member] * count)
    
    # Handle any remaining POs (if Orville didn't get extras)
    while len(assignments) < total_pos:
        # Give to Paulo, JB, Sunshine, or Stephanie (not Orville)
        assignments.append(random.choice(["Paulo", "JB", "Sunshine", "Stephanie"]))
    
    # Shuffle assignments for randomness
    random.shuffle(assignments)
    
    # Create dataframe for PO Summary
    po_summary_df = pd.DataFrame({
        'PO Number': unique_pos,
        'Assigned to': assignments[:total_pos]
    })
    
    # Write PO Summary to Excel
    po_summary_df.to_excel(writer, sheet_name=po_summary_sheet_name, index=False, startrow=1, header=False)
    po_summary_worksheet = writer.sheets[po_summary_sheet_name]
    
    # Set black tab color for PO Summary tab
    po_summary_worksheet.set_tab_color('#000000')
    
    # Write headers with formatting
    po_summary_worksheet.write(0, 0, 'PO Number', header_format)
    po_summary_worksheet.write(0, 1, 'Assigned to', header_format)
    po_summary_worksheet.write(0, 2, 'Workflow Link', header_format)
    po_summary_worksheet.write(0, 3, 'Shipment ID', header_format)
    po_summary_worksheet.write(0, 4, 'Issues', red_header_format)  # Red header for Issues
    po_summary_worksheet.write(0, 5, 'Status', header_format)
    
    # Create mapping of PO to assigned person for tab coloring later
    po_to_person = {}
    for row_num in range(len(po_summary_df)):
        po_num = str(po_summary_df.iloc[row_num, 0])
        assigned_person = str(po_summary_df.iloc[row_num, 1])
        po_to_person[po_num] = assigned_person
    
    # Apply formatting to PO Summary cells with colors based on assignment
    for row_num in range(len(po_summary_df)):
        po_num = str(po_summary_df.iloc[row_num, 0])
        assigned_person = str(po_summary_df.iloc[row_num, 1])
        
        # Get assigned person and choose appropriate color format
        if assigned_person == "Orville":
            cell_format = orville_format
        elif assigned_person == "Sunshine":
            cell_format = sunshine_format
        elif assigned_person == "Stephanie":
            cell_format = stephanie_format
        elif assigned_person == "Paulo":
            cell_format = paulo_format
        elif assigned_person == "JB":
            cell_format = jb_format
        else:
            cell_format = text_format
        
        # Color both column A (PO Number) and column B (Assigned to)
        po_summary_worksheet.write(row_num + 1, 0, po_num, cell_format)
        po_summary_worksheet.write(row_num + 1, 1, assigned_person, cell_format)
        
        # Column C (Workflow Link) - blank with borders
        po_summary_worksheet.write(row_num + 1, 2, "", text_format)
        
        # Column D (Shipment ID) - blank with borders
        po_summary_worksheet.write(row_num + 1, 3, "", text_format)
        
        # Column E (Issues) - blank with borders (only header is red)
        po_summary_worksheet.write(row_num + 1, 4, "", text_format)
        
        # Column F (Status) - formula that checks Workflow Link and Issues
        # Excel row number = row_num + 2 (row 1 is header, row 2 is first data row)
        excel_row = row_num + 2
        # Column C is Workflow Link (index 2), Column E is Issues (index 4)
        # Formula: IF both C and E have content → "WITH ISSUE", else IF C has content → "UPLOADED", else blank
        status_formula = f'=IF(AND(C{excel_row}<>"", E{excel_row}<>""), "WITH ISSUE", IF(C{excel_row}<>"", "UPLOADED", ""))'
        po_summary_worksheet.write_formula(row_num + 1, 5, status_formula, text_format)
    
    # Add conditional formatting for Status column (column F, index 5)
    # Apply yellow format when cell contains "UPLOADED"
    # Apply red format when cell contains "WITH ISSUE"
    if len(po_summary_df) > 0:
        # Status column is column F (index 5)
        # Data rows start at row 2 (Excel row 2, xlsxwriter row 1) and go to row len(po_summary_df) + 1
        first_data_row = 1  # xlsxwriter row 1 = Excel row 2
        last_data_row = len(po_summary_df)  # xlsxwriter row len = Excel row len + 1
        
        # Conditional format: if cell contains "UPLOADED", apply yellow format
        po_summary_worksheet.conditional_format(
            first_data_row, 5, last_data_row, 5,  # Column F, rows with data
            {
                'type': 'text',
                'criteria': 'containing',
                'value': 'UPLOADED',
                'format': uploaded_format
            }
        )
        
        # Conditional format: if cell contains "WITH ISSUE", apply red format
        po_summary_worksheet.conditional_format(
            first_data_row, 5, last_data_row, 5,  # Column F, rows with data
            {
                'type': 'text',
                'criteria': 'containing',
                'value': 'WITH ISSUE',
                'format': with_issue_format
            }
        )
    
    # Set column widths for PO Summary
    po_summary_worksheet.set_column(0, 0, 20)  # PO Number column
    po_summary_worksheet.set_column(1, 1, 15)  # Assigned to column
    po_summary_worksheet.set_column(2, 2, 120)  # Workflow Link column (903px ≈ 120 chars)
    po_summary_worksheet.set_column(3, 3, 20)  # Shipment ID column
    po_summary_worksheet.set_column(4, 4, 25)  # Issues column
    po_summary_worksheet.set_column(5, 5, 20)  # Status column

    # Get unique groups based on first 15 characters
    for g in df["group_15"].unique():
        group_df = df[df["group_15"] == g].copy()

        # Within each sheet, sort by shipment letter (16th char)
        group_df["shipment_letter"] = third_column[group_df.index].astype(str).str[15:16]
        group_df = group_df.sort_values(by="shipment_letter")

        # Remove helper columns before writing
        group_df = group_df.drop(columns=["shipment_letter", "group_15", "shipment"])
        
        # Create Box# column based on unique values in the first column (Carton Num)
        # Get the first column (column A - Carton Num)
        carton_col = group_df.iloc[:, 0]
        
        # Create a mapping of unique carton numbers to box numbers
        unique_cartons = carton_col.unique()
        carton_to_box = {carton: idx + 1 for idx, carton in enumerate(unique_cartons)}
        
        # Map each carton to its box number
        box_numbers = carton_col.map(carton_to_box).astype(str)
        
        # Insert Box# column at position 1 (column B)
        group_df.insert(1, 'Box#', box_numbers)
        
        # Write to Excel sheet (without default formatting)
        sheet_name = g[:31]
        group_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
        
        # Get the worksheet object
        worksheet = writer.sheets[sheet_name]
        
        # Color the tab based on assigned person
        if g in po_to_person:
            assigned_person = po_to_person[g]
            if assigned_person == "Orville":
                worksheet.set_tab_color('#FFFFE0')  # Light yellow
            elif assigned_person == "Sunshine":
                worksheet.set_tab_color('#ADD8E6')  # Light blue
            elif assigned_person == "Stephanie":
                worksheet.set_tab_color('#FFDAB9')  # Light orange
            elif assigned_person == "Paulo":
                worksheet.set_tab_color('#FFB6C1')  # Light pink
            elif assigned_person == "JB":
                worksheet.set_tab_color('#90EE90')  # Light green
        
        # Write the header row with formatting
        for col_num, value in enumerate(group_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Apply formatting to data cells as text to preserve leading zeros
        for row_num in range(len(group_df)):
            for col_num, value in enumerate(group_df.iloc[row_num]):
                # Since we read as strings, just handle NaN values
                if pd.isna(value) or value == 'nan':
                    str_value = ""
                else:
                    str_value = str(value)
                
                worksheet.write(row_num + 1, col_num, str_value, text_format)
        
        # Auto-fit column widths
        for col_num, col_name in enumerate(group_df.columns):
            # Calculate max width needed for column
            max_width = len(str(col_name)) + 2
            for value in group_df.iloc[:, col_num]:
                try:
                    cell_width = len(str(value)) + 2
                    if cell_width > max_width:
                        max_width = cell_width
                except:
                    pass
            # Set column width (max 50 to avoid extremely wide columns)
            worksheet.set_column(col_num, col_num, min(max_width, 50))
        
        # --- Add Summary: Total Boxes and Total Quantity ---
        summary_start_row = len(group_df) + 3  # Leave a blank row after data
        
        # Calculate total number of unique boxes
        if 'Box#' in group_df.columns:
            total_boxes = group_df['Box#'].nunique()
        else:
            total_boxes = 0
        
        # Calculate total quantity
        qty_col_name = None
        for col in group_df.columns:
            if 'quantity' in str(col).lower() or 'qty' in str(col).lower():
                qty_col_name = col
                break
        
        if qty_col_name:
            total_qty = pd.to_numeric(group_df[qty_col_name], errors='coerce').fillna(0).sum()
        else:
            total_qty = 0
        
        # Write summary headers and values
        worksheet.write(summary_start_row, 0, 'Total Number of Boxes:', header_format)
        worksheet.write(summary_start_row, 1, str(int(total_boxes)), bold_text_format)
        
        worksheet.write(summary_start_row + 1, 0, 'Total Quantity:', header_format)
        worksheet.write(summary_start_row + 1, 1, str(int(total_qty)), bold_text_format)
        
        # --- Create Pivot Table Summary starting at column L ---
        # Find the columns we need for pivot (UPC and Quantity)
        # Assuming columns A-J contain the data
        pivot_data = group_df.iloc[:, :10].copy()  # First 10 columns (A to J)
        
        # Identify UPC and Quantity columns - look for column names containing these terms
        upc_col = None
        qty_col = None
        
        for col in pivot_data.columns:
            col_lower = str(col).lower()
            if 'upc' in col_lower:
                upc_col = col
            if 'quantity' in col_lower or 'qty' in col_lower:
                qty_col = col
        
        # Create pivot table if we found the necessary columns
        if upc_col and qty_col and 'Box#' in pivot_data.columns:
            # Convert quantity to numeric for counting
            pivot_data[qty_col] = pd.to_numeric(pivot_data[qty_col], errors='coerce').fillna(0).astype(int)
            
            # Create pivot table: UPC as rows, Box# as columns, sum of Quantity as values
            pivot_table = pd.pivot_table(
                pivot_data,
                values=qty_col,
                index=upc_col,
                columns='Box#',
                aggfunc='sum',
                fill_value=0
            )
            
            # Write pivot table starting at column L (column index 11)
            start_col = 11
            start_row = 0
            
            # Calculate row totals (total per UPC)
            row_totals = pivot_table.sum(axis=1)
            
            # Calculate column totals (total per Box)
            col_totals = pivot_table.sum(axis=0)
            
            # Calculate grand total
            grand_total = pivot_table.sum().sum()
            
            # Write "UPC" header at L1
            worksheet.write(start_row, start_col, 'UPC', header_format)
            
            # Write Box# headers (Box 1, Box 2, etc.)
            for i, box_num in enumerate(pivot_table.columns):
                worksheet.write(start_row, start_col + 1 + i, f'Box {box_num}', header_format)
            
            # Write "Total" header for row totals column
            worksheet.write(start_row, start_col + 1 + len(pivot_table.columns), 'Total', header_format)
            
            # Write the pivot table data with row totals
            for row_idx, upc in enumerate(pivot_table.index):
                # Write UPC value
                worksheet.write(start_row + 1 + row_idx, start_col, str(upc), text_format)
                
                # Write quantities for each box (leave blank if zero)
                for col_idx, box_num in enumerate(pivot_table.columns):
                    qty = int(pivot_table.loc[upc, box_num])
                    qty_value = "" if qty == 0 else str(qty)
                    worksheet.write(start_row + 1 + row_idx, start_col + 1 + col_idx, qty_value, text_format)
                
                # Write row total (bold, leave blank if zero)
                row_total = int(row_totals[upc])
                row_total_value = "" if row_total == 0 else str(row_total)
                worksheet.write(start_row + 1 + row_idx, start_col + 1 + len(pivot_table.columns), row_total_value, bold_text_format)
            
            # Write "Total" row at the bottom
            total_row_idx = start_row + 1 + len(pivot_table.index)
            worksheet.write(total_row_idx, start_col, 'Total', header_format)
            
            # Write column totals (bold, leave blank if zero)
            for col_idx, box_num in enumerate(pivot_table.columns):
                col_total = int(col_totals[box_num])
                col_total_value = "" if col_total == 0 else str(col_total)
                worksheet.write(total_row_idx, start_col + 1 + col_idx, col_total_value, bold_text_format)
            
            # Write grand total (bold)
            worksheet.write(total_row_idx, start_col + 1 + len(pivot_table.columns), str(int(grand_total)), bold_text_format)
            
            # Auto-fit columns for pivot table
            worksheet.set_column(start_col, start_col, 25)  # UPC column
            for i in range(len(pivot_table.columns) + 1):  # +1 for Total column
                worksheet.set_column(start_col + 1 + i, start_col + 1 + i, 12)  # Box columns and Total column

    writer.close()
    st.success("Processing complete!")
    
    # Generate filename with Central time
    central_tz = pytz.timezone('America/Chicago')
    current_time = datetime.now(central_tz)
    formatted_time = current_time.strftime('%Y-%m-%d %I-%M-%S %p')
    output_filename = f"SMW Bulk Shipments {formatted_time}.xlsx"

    st.download_button(
        label="Download Organized Excel File",
        data=output.getvalue(),
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
