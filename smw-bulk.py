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

def shuffle_no_consecutive(items):
    """Shuffle items ensuring no two consecutive items are the same when possible"""
    if len(items) <= 1:
        return items

    from collections import Counter
    counts = Counter(items)
    result = []
    remaining = list(items)
    last_item = None

    # Try to avoid consecutive duplicates
    while remaining:
        # Get available items (different from last)
        available = [x for x in remaining if x != last_item]

        if available:
            # Choose randomly from available items
            chosen = random.choice(available)
        else:
            # No choice - must use the same as last (unavoidable)
            chosen = remaining[0]

        result.append(chosen)
        remaining.remove(chosen)
        last_item = chosen

    # Final pass: try to fix any remaining consecutive duplicates by swapping
    for i in range(len(result) - 1):
        if result[i] == result[i + 1]:
            # Try to find a different item to swap with
            for j in range(i + 2, len(result)):
                if result[j] != result[i] and (j == len(result) - 1 or result[j] != result[i + 1]):
                    result[i + 1], result[j] = result[j], result[i + 1]
                    break

    return result

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
    
    # Create a mapping of group_15 to the full PO number (first occurrence)
    group_to_full_po = {}
    for idx in df.index:
        group_key = df.loc[idx, "group_15"]
        if group_key not in group_to_full_po:
            # Store the full value from column C
            group_to_full_po[group_key] = str(third_column.loc[idx])

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
        'bg_color': '#FF0000',
        'font_color': 'white',
        'border': 1
    })
    
    # Text format for all data - preserves leading zeros and shows full numbers
    text_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '@',
        'text_wrap': False,
        'locked': False
    })
    
    # Locked text format for protected formula cells
    locked_text_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '@',
        'locked': True
    })
    
    # Bold text format for grand total
    bold_text_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bold': True,
        'num_format': '@'
    })
    
    # Dark orange format for pivot table headers and totals
    dark_orange_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'fg_color': '#CC6600',
        'font_color': 'white',
        'border': 1
    })
    
    # Number format for numeric column formatting
    number_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'num_format': '0',
        'locked': False
    })
    
    # Color formats for team member assignments (unlocked for editing)
    orville_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFFFE0',  # Light yellow
        'num_format': '@',
        'text_wrap': False,
        'locked': False
    })
    
    sunshine_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#ADD8E6',  # Light blue
        'num_format': '@',
        'text_wrap': False,
        'locked': False
    })
    
    stephanie_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFDAB9',
        'num_format': '@',
        'text_wrap': False,
        'locked': False
    })
    
    paulo_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFB6C1',
        'num_format': '@',
        'text_wrap': False,
        'locked': False
    })
    
    jb_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#90EE90',
        'num_format': '@',
        'text_wrap': False,
        'locked': False
    })
    
    # Red format for Issues column with white text
    red_issues_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FF6B6B',
        'font_color': 'white',
        'num_format': '@'
    })
    
    # Red highlight format for PO Number column when missing letters
    red_highlight_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FF0000',
        'font_color': 'white',
        'num_format': '@',
        'text_wrap': False
    })
    
    # Red format for missing PO number warning message
    red_warning_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FF0000',
        'font_color': 'white',
        'num_format': '@',
        'bold': True
    })
    
    # Yellow format for "UPLOADED" status
    uploaded_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFFF00',
        'num_format': '@',
        'locked': False
    })
    
    # Red format for "WITH ISSUE" status
    with_issue_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FF0000',
        'font_color': 'white',
        'num_format': '@',
        'locked': False
    })
    
    # Orange format for "AWAITING UPLOAD" status
    awaiting_upload_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FFA500',
        'num_format': '@',
        'locked': False
    })
    
    # --- Write original data to first sheet ---
    original_sheet_name = "Original Data"
    df_original.to_excel(writer, sheet_name=original_sheet_name, index=False, startrow=1, header=False)
    
    # Get the original data worksheet
    original_worksheet = writer.sheets[original_sheet_name]
    
    # Set black tab color for Original Data tab
    original_worksheet.set_tab_color('#000000')
    
    # Write headers with blue header format
    for col_num, col_name in enumerate(df_original.columns):
        original_worksheet.write(0, col_num, col_name, header_format)
    
    # Apply text format to all data cells
    for row_num in range(len(df_original)):
        for col_num, value in enumerate(df_original.iloc[row_num]):
            # Skip writing empty cells in column H (index 7) to avoid borders
            if col_num == 7:
                if pd.isna(value) or value == 'nan' or str(value).strip() == '':
                    continue  # Skip empty cells in column H
            
            # Handle NaN values
            if pd.isna(value) or value == 'nan':
                str_value = ""
            else:
                str_value = str(value)
            original_worksheet.write(row_num + 1, col_num, str_value, text_format)
    
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
    
    # Format column H (index 7) as numbers in Original Data tab
    if len(df_original.columns) > 7:
        last_data_row = -1
        for row_num in range(len(df_original) - 1, -1, -1):
            value = df_original.iloc[row_num, 7]
            if not (pd.isna(value) or value == 'nan' or str(value).strip() == ''):
                try:
                    num_value = pd.to_numeric(value, errors='coerce')
                    if not pd.isna(num_value):
                        last_data_row = row_num
                        break
                except:
                    pass
        
        for row_num in range(len(df_original)):
            if row_num > last_data_row:
                break
            value = df_original.iloc[row_num, 7]
            if pd.isna(value) or value == 'nan' or str(value).strip() == '':
                continue
            else:
                try:
                    num_value = pd.to_numeric(value, errors='coerce')
                    if pd.isna(num_value):
                        continue
                    else:
                        original_worksheet.write(row_num + 1, 7, float(num_value), number_format)
                except:
                    continue
    
    # --- Create PO Summary tab (second tab) ---
    po_summary_sheet_name = "PO Summary"
    
    # Get unique PO numbers - use full PO numbers for grouping
    unique_groups = sorted(df["group_15"].unique())
    unique_pos_full = [group_to_full_po[g] for g in unique_groups]  # Use full PO numbers
    
    # Process PO numbers: remove last character if it's a letter, keep if it's a number
    def process_po_number(po_num):
        """Remove last character if it's a letter, keep if it's a number"""
        po_str = str(po_num)
        if len(po_str) > 0:
            last_char = po_str[-1]
            if last_char.isalpha():
                return po_str[:-1]
            return po_str
        return po_str
    
    # Process PO numbers and handle duplicates
    full_to_processed = {}
    processed_pos = []
    seen_pos = set()
    for po_full in unique_pos_full:
        po_processed = process_po_number(po_full)
        full_to_processed[po_full] = po_processed
        if po_processed not in seen_pos:
            processed_pos.append(po_processed)
            seen_pos.add(po_processed)
    
    unique_pos = processed_pos
    total_pos = len(unique_pos)
    
    # Team members - Orville gets lower priority for remainders
    team_members = ["Paulo", "JB", "Stephanie", "Sunshine", "Orville"]
    preferred_for_remainder = ["Paulo", "JB", "Stephanie", "Sunshine"]  # Orville excluded
    
    # Calculate base assignment per person
    base_per_person = total_pos // len(team_members)
    remainder = total_pos % len(team_members)
    
    # Create assignment list
    assignments = []
    # Give base_per_person to everyone
    for m in team_members:
        assignments.extend([m] * base_per_person)
    
    # Distribute remainder to preferred members (Paulo, JB, Stephanie, Sunshine)
    for i in range(remainder):
        assignments.append(preferred_for_remainder[i % len(preferred_for_remainder)])
    
    # If still short (edge cases), fill with preferred members
    while len(assignments) < total_pos:
        assignments.append(random.choice(preferred_for_remainder))
    
    # Ensure Sunshine appears at least once when there is at least one PO
    if total_pos > 0 and "Sunshine" not in assignments:
        for idx, a in enumerate(assignments):
            if a != "Orville":
                assignments[idx] = "Sunshine"
                break
    
    # Shuffle to reduce consecutive duplicates, preserving counts
    assignments = shuffle_no_consecutive(assignments)
    
    # Trim to exact length just in case
    assignments = assignments[:total_pos]
    
    # Create dataframe for PO Summary
    po_summary_df = pd.DataFrame({
        'PO Number': unique_pos,
        'Assigned to': assignments[:total_pos]
    })
    
    # Show assignment preview in Streamlit so you can confirm Sunshine appears
    st.write("Assignment preview:", po_summary_df)
    
    # Write PO Summary to Excel
    po_summary_df.to_excel(writer, sheet_name=po_summary_sheet_name, index=False, startrow=1, header=False)
    po_summary_worksheet = writer.sheets[po_summary_sheet_name]
    
    # Set black tab color for PO Summary tab
    po_summary_worksheet.set_tab_color('#000000')
    
    # Write headers with formatting
    po_summary_worksheet.write(0, 0, 'PO Number', header_format)
    po_summary_worksheet.write(0, 1, 'Assigned to', header_format)
    po_summary_worksheet.write(0, 2, 'Workflow Link', header_format)
    po_summary_worksheet.write(0, 3, 'Issues', red_header_format)  # Red header for Issues
    po_summary_worksheet.write(0, 4, 'Status', header_format)
    
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
        elif assigned_person == "Stephanie":
            cell_format = stephanie_format
        elif assigned_person == "Paulo":
            cell_format = paulo_format
        elif assigned_person == "JB":
            cell_format = jb_format
        elif assigned_person == "Sunshine":
            cell_format = sunshine_format
        else:
            cell_format = text_format
        
        # Color both column A (PO Number) and column B (Assigned to)
        po_summary_worksheet.write(row_num + 1, 0, po_num, cell_format)
        po_summary_worksheet.write(row_num + 1, 1, assigned_person, cell_format)
        
        # Column C (Workflow Link) - blank with borders
        po_summary_worksheet.write(row_num + 1, 2, "", text_format)
        
        # Column D (Issues) - blank with borders (only header is red)
        po_summary_worksheet.write(row_num + 1, 3, "", text_format)
        
        # Column E (Status) - formula that checks Workflow Link and Issues
        excel_row = row_num + 2
        status_formula = f'=IF(AND(C{excel_row}="", D{excel_row}=""), "AWAITING UPLOAD", IF(AND(C{excel_row}="", D{excel_row}<>""), "WITH ISSUE", IF(AND(C{excel_row}<>"", D{excel_row}<>""), "WITH ISSUE", "UPLOADED")))' 
        po_summary_worksheet.write_formula(row_num + 1, 4, status_formula, text_format)
    
    # Add conditional formatting for Status column (column E, index 4)
    if len(po_summary_df) > 0:
        first_data_row = 1
        last_data_row = len(po_summary_df)
        
        # Status conditional formats
        po_summary_worksheet.conditional_format(
            first_data_row, 4, last_data_row, 4,
            {'type': 'text', 'criteria': 'containing', 'value': 'UPLOADED', 'format': uploaded_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 4, last_data_row, 4,
            {'type': 'text', 'criteria': 'containing', 'value': 'WITH ISSUE', 'format': with_issue_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 4, last_data_row, 4,
            {'type': 'text', 'criteria': 'containing', 'value': 'AWAITING UPLOAD', 'format': awaiting_upload_format}
        )
        
        # Assigned-to conditional formats (column B)
        po_summary_worksheet.conditional_format(
            first_data_row, 1, last_data_row, 1,
            {'type': 'text', 'criteria': 'containing', 'value': 'Orville', 'format': orville_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 1, last_data_row, 1,
            {'type': 'text', 'criteria': 'containing', 'value': 'Stephanie', 'format': stephanie_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 1, last_data_row, 1,
            {'type': 'text', 'criteria': 'containing', 'value': 'Paulo', 'format': paulo_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 1, last_data_row, 1,
            {'type': 'text', 'criteria': 'containing', 'value': 'JB', 'format': jb_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 1, last_data_row, 1,
            {'type': 'text', 'criteria': 'containing', 'value': 'Sunshine', 'format': sunshine_format}
        )
        
        # Column A (PO Number) formula-based conditional formats referencing column B
        po_summary_worksheet.conditional_format(
            first_data_row, 0, last_data_row, 0,
            {'type': 'formula', 'criteria': '=$B2="Orville"', 'format': orville_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 0, last_data_row, 0,
            {'type': 'formula', 'criteria': '=$B2="Stephanie"', 'format': stephanie_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 0, last_data_row, 0,
            {'type': 'formula', 'criteria': '=$B2="Paulo"', 'format': paulo_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 0, last_data_row, 0,
            {'type': 'formula', 'criteria': '=$B2="JB"', 'format': jb_format}
        )
        po_summary_worksheet.conditional_format(
            first_data_row, 0, last_data_row, 0,
            {'type': 'formula', 'criteria': '=$B2="Sunshine"', 'format': sunshine_format}
        )
    
    # Set column widths for PO Summary - wider columns for better visibility
    po_summary_worksheet.set_column(0, 0, 30)
    po_summary_worksheet.set_column(1, 1, 18)
    po_summary_worksheet.set_column(2, 2, 120)
    po_summary_worksheet.set_column(3, 3, 30)
    po_summary_worksheet.set_column(4, 4, 25)

    # Get unique groups based on first 15 characters
    unique_groups_list = list(df["group_15"].unique())
    
    # Create list of (group, processed_po) tuples and sort by processed_po
    groups_with_processed_pos = []
    for g in unique_groups_list:
        full_po = group_to_full_po[g]
        processed_po = process_po_number(full_po)
        groups_with_processed_pos.append((g, processed_po))
    
    # Sort by processed PO number alphabetically
    groups_with_processed_pos.sort(key=lambda x: x[1])
    
    # Create sheets in alphabetical order
    for g, processed_po in groups_with_processed_pos:
        group_df = df[df["group_15"] == g].copy()

        # Remove helper columns before processing
        group_df = group_df.drop(columns=["group_15", "shipment"])
        
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
        
        # Sort by Box# first (Column B, index 1) numerically, then by PO Number (Column D, index 3) alphabetically
        box_col_name = group_df.columns[1]  # Column B (index 1) is the Box# column
        po_col_name = group_df.columns[3]  # Column D (index 3) is the PO Number column
        
        # Convert Box# to numeric for proper numerical sorting, PO Number to string for alphabetical sorting
        group_df[box_col_name] = pd.to_numeric(group_df[box_col_name], errors='coerce').fillna(0)
        group_df[po_col_name] = group_df[po_col_name].astype(str)
        
        # Sort by Box# first (ascending), then by PO Number (alphabetically)
        group_df = group_df.sort_values(by=[box_col_name, po_col_name])
        
        # Convert Box# back to string for display
        group_df[box_col_name] = group_df[box_col_name].astype(int).astype(str)
        
        # Write to Excel sheet (without default formatting)
        full_po = group_to_full_po[g]
        sheet_name = process_po_number(full_po)[:31]  # Use processed PO, truncated to 31 chars
        group_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
        
        # Get the worksheet object
        worksheet = writer.sheets[sheet_name]
        
        # Color the tab based on assigned person
        processed_po = process_po_number(full_po)
        if processed_po in po_to_person:
            assigned_person = po_to_person[processed_po]
            if assigned_person == "Orville":
                worksheet.set_tab_color('#FFFFE0')  # Light yellow
            elif assigned_person == "Stephanie":
                worksheet.set_tab_color('#FFDAB9')  # Light orange
            elif assigned_person == "Paulo":
                worksheet.set_tab_color('#FFB6C1')  # Light pink
            elif assigned_person == "JB":
                worksheet.set_tab_color('#90EE90')  # Light green
            elif assigned_person == "Sunshine":
                worksheet.set_tab_color('#ADD8E6')  # Light blue
        
        # Write the header row with formatting
        for col_num, value in enumerate(group_df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Apply formatting to data cells as text to preserve leading zeros
        for row_num in range(len(group_df)):
            for col_num, value in enumerate(group_df.iloc[row_num]):
                # Skip writing empty cells in column I (index 8) to avoid borders
                if col_num == 8:
                    if pd.isna(value) or value == 'nan' or str(value).strip() == '':
                        continue  # Skip empty cells in column I
                
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
            worksheet.set_column(col_num, col_num, min(max_width, 50))
        
        # Format column I (index 8) as numbers in all group sheets
        if len(group_df.columns) > 8:
            last_data_row = -1
            for row_num in range(len(group_df) - 1, -1, -1):
                value = group_df.iloc[row_num, 8]
                if not (pd.isna(value) or value == 'nan' or str(value).strip() == ''):
                    try:
                        num_value = pd.to_numeric(value, errors='coerce')
                        if not pd.isna(num_value):
                            last_data_row = row_num
                            break
                    except:
                        pass
            
            for row_num in range(len(group_df)):
                if row_num > last_data_row:
                    break
                value = group_df.iloc[row_num, 8]
                if pd.isna(value) or value == 'nan' or str(value).strip() == '':
                    continue
                else:
                    try:
                        num_value = pd.to_numeric(value, errors='coerce')
                        if pd.isna(num_value):
                            continue
                        else:
                            worksheet.write(row_num + 1, 8, float(num_value), number_format)
                    except:
                        continue
        
        # --- Add Summary: Total Boxes and Total Quantity ---
        summary_start_row = len(group_df) + 3  # Leave a blank row after data
        
        if 'Box#' in group_df.columns:
            total_boxes = group_df['Box#'].nunique()
        else:
            total_boxes = 0
        
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
        
        # --- Check for missing PO Number letters in Column D ---
        po_col_name = group_df.columns[3]  # Column D (index 3) is the PO Number column
        po_numbers = group_df[po_col_name].astype(str)
        
        last_letters = []
        for po in po_numbers:
            if len(po) > 0:
                last_char = po[-1].upper()
                if last_char.isalpha():
                    last_letters.append(last_char)
        
        has_missing = False
        if len(last_letters) > 0:
            unique_letters = sorted(set(last_letters))
            if unique_letters and unique_letters[0] == 'A':
                if len(unique_letters) > 1:
                    for i in range(len(unique_letters) - 1):
                        current_letter = unique_letters[i]
                        next_letter = unique_letters[i + 1]
                        if ord(next_letter) - ord(current_letter) > 1:
                            has_missing = True
                            break
        
        if has_missing:
            for row_num in range(len(group_df)):
                po_value = str(group_df.iloc[row_num, 3])
                worksheet.write(row_num + 1, 3, po_value, red_highlight_format)
            
            worksheet.write(summary_start_row + 2, 0, 'With Missing PO Number', red_warning_format)
            worksheet.write(summary_start_row + 2, 1, '', red_warning_format)
        
        # --- Create Pivot Table Summary starting at column Q ---
        # Fixed version: ensure Box# treated numeric and columns sorted numerically
        pivot_data = group_df.iloc[:, :10].copy()  # First 10 columns (A to J)
        
        upc_col = None
        qty_col = None
        for col in pivot_data.columns:
            col_lower = str(col).lower()
            if 'upc' in col_lower:
                upc_col = col
            if 'quantity' in col_lower or 'qty' in col_lower:
                qty_col = col
        
        if upc_col and qty_col and 'Box#' in pivot_data.columns:
            # Ensure Box# numeric and quantities numeric
            pivot_data['Box#'] = pd.to_numeric(pivot_data['Box#'], errors='coerce').fillna(0).astype(int)
            pivot_data[qty_col] = pd.to_numeric(pivot_data[qty_col], errors='coerce').fillna(0).astype(int)
            
            pivot_table = pd.pivot_table(
                pivot_data,
                values=qty_col,
                index=upc_col,
                columns='Box#',
                aggfunc='sum',
                fill_value=0
            )
            
            # Ensure columns sorted numerically
            pivot_table = pivot_table.reindex(sorted(pivot_table.columns), axis=1)
            
            # Calculate totals
            row_totals = pivot_table.sum(axis=1)
            col_totals = pivot_table.sum(axis=0)
            grand_total = pivot_table.sum().sum()
            
            # Write pivot table starting at column Q (index 16)
            start_col = 16
            start_row = 0
            
            worksheet.write(start_row, start_col, 'UPC', dark_orange_format)
            
            # Write Box headers in numeric order
            for i, box_num in enumerate(pivot_table.columns):
                worksheet.write(start_row, start_col + 1 + i, f'Box {box_num}', dark_orange_format)
            
            # Write "Total" header for row totals column
            worksheet.write(start_row, start_col + 1 + len(pivot_table.columns), 'Total', dark_orange_format)
            
            # Write pivot rows
            for row_idx, upc in enumerate(pivot_table.index):
                worksheet.write(start_row + 1 + row_idx, start_col, str(upc), text_format)
                
                for col_idx, box_num in enumerate(pivot_table.columns):
                    qty = int(pivot_table.loc[upc, box_num])
                    qty_value = "" if qty == 0 else str(qty)
                    worksheet.write(start_row + 1 + row_idx, start_col + 1 + col_idx, qty_value, text_format)
                
                row_total = int(row_totals[upc])
                row_total_value = "" if row_total == 0 else str(row_total)
                worksheet.write(start_row + 1 + row_idx, start_col + 1 + len(pivot_table.columns), row_total_value, bold_text_format)
            
            # Write totals
            total_row_idx = start_row + 1 + len(pivot_table.index)
            worksheet.write(total_row_idx, start_col, 'Total', dark_orange_format)
            
            for col_idx, box_num in enumerate(pivot_table.columns):
                col_total = int(col_totals[box_num])
                col_total_value = "" if col_total == 0 else str(col_total)
                worksheet.write(total_row_idx, start_col + 1 + col_idx, col_total_value, dark_orange_format)
            
            worksheet.write(total_row_idx, start_col + 1 + len(pivot_table.columns), str(int(grand_total)), dark_orange_format)
            
            # Auto-fit pivot table columns
            worksheet.set_column(start_col, start_col, 25)
            for i in range(len(pivot_table.columns) + 1):
                worksheet.set_column(start_col + 1 + i, start_col + 1 + i, 12)

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
