import streamlit as st
import pandas as pd
import io
import random
from datetime import datetime
import pytz

st.title("Shipment Grouping Tool")
st.write(
    "Upload an Excel file. This tool will group rows based on the first 15 characters "
    "of Column C and separate shipments (A, B, C...). In each group sheet paste the raw URL "
    "into Column F (editable). Column G will auto-create a clickable hyperlink (HYPERLINK) showing the full URL. "
    "PO Summary will mirror the clickable URL from Column F (creating its own HYPERLINK)."
)

uploaded = st.file_uploader("Upload Excel File", type=["xlsx"])


def shuffle_no_consecutive(items):
    if len(items) <= 1:
        return items
    result = []
    remaining = list(items)
    last_item = None
    while remaining:
        available = [x for x in remaining if x != last_item]
        chosen = random.choice(available) if available else remaining[0]
        result.append(chosen)
        remaining.remove(chosen)
        last_item = chosen
    for i in range(len(result) - 1):
        if result[i] == result[i + 1]:
            for j in range(i + 2, len(result)):
                if result[j] != result[i] and (j == len(result) - 1 or result[j] != result[i + 1]):
                    result[i + 1], result[j] = result[j], result[i + 1]
                    break
    return result


if uploaded:
    df = pd.read_excel(uploaded, dtype=str)
    if len(df.columns) < 3:
        st.error("File needs at least 3 columns (A, B, C). Please check your file.")
        st.stop()

    df_original = df.copy()
    third_column = df.iloc[:, 2]
    df["group_15"] = third_column.astype(str).str[:15]
    df["shipment"] = third_column.astype(str).str[:16]

    group_to_full_po = {}
    for idx in df.index:
        g = df.loc[idx, "group_15"]
        if g not in group_to_full_po:
            group_to_full_po[g] = str(third_column.loc[idx])

    df = df.sort_values(by=["group_15", "shipment"])

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    workbook = writer.book

    # --- FORMATS ---
    header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
    })
    red_header_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'bg_color': '#FF0000', 'font_color': 'white', 'border': 1
    })
    text_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '@',
        'locked': False
    })
    locked_text_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '@',
        'locked': True
    })
    bold_text_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True, 'num_format': '@'
    })
    dark_orange_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'fg_color': '#CC6600', 'font_color': 'white', 'border': 1
    })

    number_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1,
        'num_format': '0', 'locked': False
    })

    # TEAM COLORS
    orville_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FFFFE0','num_format': '@','locked': False})
    sunshine_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#ADD8E6','num_format': '@','locked': False})
    stephanie_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FFDAB9','num_format': '@','locked': False})
    paulo_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FFB6C1','num_format': '@','locked': False})
    jb_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#90EE90','num_format': '@','locked': False})

    red_highlight_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FF0000','font_color': 'white'})
    red_warning_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FF0000','font_color': 'white','bold': True})

    uploaded_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FFFF00'})
    with_issue_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FF0000','font_color': 'white'})
    awaiting_upload_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 1,'bg_color': '#FFA500'})

    maroon_no_border = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'fg_color': '#800000', 'font_color': 'white', 'border': 0
    })

    # --- ORIGINAL DATA SHEET ---
    original_sheet_name = "Original Data"
    ws_original = workbook.add_worksheet(original_sheet_name)
    writer.sheets[original_sheet_name] = ws_original
    ws_original.set_tab_color('#000000')

    # Write headers
    for c, name in enumerate(df_original.columns):
        ws_original.write(0, c, name, header_format)

    # Write data WITH BORDER
    for r in range(len(df_original)):
        for c, val in enumerate(df_original.iloc[r]):
            val = "" if pd.isna(val) else str(val)
            ws_original.write(r + 1, c, val, text_format)

    # Autofit
    for col in range(len(df_original.columns)):
        max_width = len(str(df_original.columns[col])) + 2
        for val in df_original.iloc[:, col]:
            w = len(str(val)) + 2
            max_width = max(max_width, w)
        ws_original.set_column(col, col, min(max_width, 50))

    # --- PO SUMMARY PREP ---
    po_summary_sheet_name = "PO Summary"

    unique_groups = sorted(df["group_15"].unique())
    unique_pos_full = [group_to_full_po[g] for g in unique_groups]

    def process_po_number(po):
        s = str(po)
        if s and s[-1].isalpha():
            return s[:-1]
        return s

    full_to_processed = {}
    processed_pos = []
    seen = set()

    for full in unique_pos_full:
        proc = process_po_number(full)
        full_to_processed[full] = proc
        if proc not in seen:
            processed_pos.append(proc)
            seen.add(proc)

    processed_pos = sorted(processed_pos)
    unique_pos = processed_pos

    team_members = ["Paulo", "JB", "Stephanie", "Sunshine", "Orville"]
    total_pos = len(unique_pos)

    base = total_pos // len(team_members)
    remainder = total_pos % len(team_members)

    assignments = []
    for i, t in enumerate(team_members):
        n = base + (1 if i < remainder else 0)
        assignments.extend([t] * n)

    assignments = assignments[:total_pos]

    po_summary_df = pd.DataFrame({
        "PO Number": unique_pos,
        "Assigned to": assignments,
        "Workflow Link": ["" for _ in range(total_pos)],
    })

    ws_po = workbook.add_worksheet(po_summary_sheet_name)
    writer.sheets[po_summary_sheet_name] = ws_po
    ws_po.set_tab_color('#000000')

    ws_po.write(0, 0, "PO Number", header_format)
    ws_po.write(0, 1, "Assigned to", header_format)
    ws_po.write(0, 2, "Workflow Link", header_format)
    ws_po.write(0, 3, "Issues", red_header_format)
    ws_po.write(0, 4, "Status", header_format)

    po_to_person = {str(po_summary_df.iloc[i, 0]): str(po_summary_df.iloc[i, 1])
                    for i in range(len(po_summary_df))}

    group_sheet_link_locations = {}

    # -------------------------------------------------------------
    #              PROCESS EACH GROUP SHEET
    # -------------------------------------------------------------
    groups_sorted = []
    for g in df["group_15"].unique():
        full_po = group_to_full_po[g]
        proc_po = process_po_number(full_po)
        groups_sorted.append((g, proc_po))
    groups_sorted.sort(key=lambda x: x[1])

    for g, proc_po in groups_sorted:

        group_df = df[df["group_15"] == g].copy()
        group_df = group_df.drop(columns=["group_15", "shipment"])

        carton_col = group_df.iloc[:, 0]
        unique_cartons = carton_col.unique()
        carton_to_box = {carton: i + 1 for i, carton in enumerate(unique_cartons)}
        box_numbers = carton_col.map(carton_to_box).astype(str)
        group_df.insert(1, "Box#", box_numbers)

        box_col = group_df.columns[1]
        po_col = group_df.columns[3]

        group_df[box_col] = pd.to_numeric(group_df[box_col], errors="coerce").fillna(0)
        group_df = group_df.sort_values(by=[box_col, po_col])
        group_df[box_col] = group_df[box_col].astype(int).astype(str)

        full_po = group_to_full_po[g]
        sheet_name = process_po_number(full_po)[:31]

        ws = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = ws

        if proc_po in po_to_person:
            person = po_to_person[proc_po]
            if person == "Orville": ws.set_tab_color("#FFFFE0")
            elif person == "Stephanie": ws.set_tab_color("#FFDAB9")
            elif person == "Paulo": ws.set_tab_color("#FFB6C1")
            elif person == "JB": ws.set_tab_color("#90EE90")
            elif person == "Sunshine": ws.set_tab_color("#ADD8E6")

        # Write headers
        for c, col in enumerate(group_df.columns.values):
            ws.write(0, c, col, header_format)

        # Write values
        for r in range(len(group_df)):
            for c, val in enumerate(group_df.iloc[r]):
                val = "" if pd.isna(val) or val == "nan" else str(val)
                ws.write(r + 1, c, val, text_format)

        # Autofit
        for col in range(len(group_df.columns)):
            max_width = len(str(group_df.columns[col])) + 2
            for v in group_df.iloc[:, col]:
                w = len(str(v)) + 2
                max_width = max(max_width, w)
            ws.set_column(col, col, min(max_width, 50))

        # -------------------------------------------------------------
        #   HIGHLIGHT COLUMNS L, M, N, O WITH RED IF BLANK OR ZERO
        # -------------------------------------------------------------
        red_dim_fill = workbook.add_format({
            'bg_color': '#FF0000',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        # Excel columns L=11, M=12, N=13, O=14
        target_dim_cols = [11, 12, 13, 14]

        for rr in range(len(group_df)):
            for cc in target_dim_cols:
                if cc < len(group_df.columns):
                    val = group_df.iat[rr, cc]
                    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "0":
                        ws.write(rr + 1, cc, "" if pd.isna(val) else str(val), red_dim_fill)

        # -------------------------------------------------------------
        # Totals section
        # -------------------------------------------------------------
        summary_start_row = len(group_df) + 3
        total_boxes = group_df["Box#"].nunique()

        qty_col = None
        for col in group_df.columns:
            if "qty" in col.lower() or "quantity" in col.lower():
                qty_col = col
                break

        total_qty = (
            pd.to_numeric(group_df[qty_col], errors="coerce").fillna(0).sum()
            if qty_col else 0
        )

        ws.write(summary_start_row, 0, "Total Number of Boxes:", header_format)
        ws.write(summary_start_row, 1, str(int(total_boxes)), bold_text_format)
        ws.write(summary_start_row + 1, 0, "Total Quantity:", header_format)
        ws.write(summary_start_row + 1, 1, str(int(total_qty)), bold_text_format)

        # Workflow link
        link_row = summary_start_row + 3
        ws.write(link_row, 4, "Workflow Link:", maroon_no_border)
        ws.write(link_row, 5, "", text_format)
        excel_row = link_row + 1

        ws.write_formula(
            link_row, 6,
            f'=IF(TRIM(F{excel_row})="","",HYPERLINK(F{excel_row},F{excel_row}))',
            text_format,
        )

        ws.set_column(5, 5, 80)
        ws.set_column(6, 6, 120)

        group_sheet_link_locations[proc_po] = (sheet_name, excel_row)

        # Missing PO detection
        po_letters = []
        for po in group_df[po_col]:
            if po and str(po)[-1].isalpha():
                po_letters.append(str(po)[-1].upper())

        missing = False
        if po_letters:
            uniq = sorted(set(po_letters))
            if uniq and uniq[0] == "A":
                for i in range(len(uniq) - 1):
                    if ord(uniq[i + 1]) - ord(uniq[i]) > 1:
                        missing = True
                        break

        if missing:
            for r in range(len(group_df)):
                ws.write(r + 1, 3, group_df.iloc[r, 3], red_highlight_format)
            ws.write(summary_start_row + 2, 0, "With Missing PO Number", red_warning_format)

        # Pivot data
        pivot_data = group_df.iloc[:, :10].copy()

        upc_col = None
        qty_col = None
        for col in pivot_data.columns:
            lc = col.lower()
            if "upc" in lc: upc_col = col
            if "qty" in lc or "quantity" in lc: qty_col = col

        if upc_col and qty_col and "Box#" in pivot_data.columns:

            pivot_data["Box#"] = pd.to_numeric(pivot_data["Box#"], errors="coerce").fillna(0).astype(int)
            pivot_data[qty_col] = pd.to_numeric(pivot_data[qty_col], errors="coerce").fillna(0).astype(int)

            pivot = pd.pivot_table(
                pivot_data,
                values=qty_col,
                index=upc_col,
                columns="Box#",
                aggfunc="sum",
                fill_value=0
            )
            pivot = pivot.reindex(sorted(pivot.columns), axis=1)

            row_totals = pivot.sum(axis=1)
            col_totals = pivot.sum(axis=0)
            grand_total = pivot.sum().sum()

            start_col = 16
            start_row = 0

            ws.write(start_row, start_col, "UPC", dark_orange_format)
            for i, box in enumerate(pivot.columns):
                ws.write(start_row, start_col + 1 + i, f"Box {box}", dark_orange_format)
            ws.write(start_row, start_col + 1 + len(pivot.columns), "Total", dark_orange_format)

            for r, upc in enumerate(pivot.index):
                ws.write(start_row + 1 + r, start_col, str(upc), text_format)
                for c, box in enumerate(pivot.columns):
                    qty = pivot.loc[upc, box]
                    ws.write(start_row + 1 + r, start_col + 1 + c,
                             "" if qty == 0 else qty, text_format)
                ws.write(start_row + 1 + r,
                         start_col + 1 + len(pivot.columns),
                         row_totals[upc],
                         bold_text_format)

            total_row = start_row + 1 + len(pivot.index)
            ws.write(total_row, start_col, "Total", dark_orange_format)
            for c, box in enumerate(pivot.columns):
                ws.write(total_row, start_col + 1 + c,
                         col_totals[box], dark_orange_format)
            ws.write(total_row, start_col + 1 + len(pivot.columns),
                     grand_total, dark_orange_format)

            ws.set_column(start_col, start_col, 25)
            for i in range(len(pivot.columns) + 1):
                ws.set_column(start_col + 1 + i, start_col + 1 + i, 12)

            blank_col = start_col + 1 + len(pivot.columns) + 1
            ws.set_column(blank_col, blank_col, 3)

            # Dimensions summary
            dim_indices = [11, 12, 13, 14] if group_df.shape[1] >= 15 else []

            teal_header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter',
                'align': 'center', 'fg_color': '#008080',
                'font_color': 'white', 'border': 1
            })

            if dim_indices:
                box_idx = 1
                selected = [box_idx] + dim_indices
                dim_df = group_df.iloc[:, selected].copy()
                dim_df.columns = [
                    "Box#", "Pkg Wt (Lbs)", "Pkg Length (in)",
                    "Pkg Width (in)", "Pkg Height (in)"
                ]

                # Remove duplicates
                dim_df = dim_df.drop_duplicates(subset=["Box#"], keep="first")

                for c in dim_df.columns[1:]:
                    dim_df[c] = dim_df[c].fillna("").astype(str)

                dim_df["Box#_sort"] = pd.to_numeric(dim_df["Box#"], errors="coerce") \
                                      .fillna(0).astype(int)
                dim_df = dim_df.sort_values(by="Box#_sort").drop(columns=["Box#_sort"])

                summary_start_col = blank_col + 1
                summary_start_row = start_row

                for i, col in enumerate(dim_df.columns):
                    ws.write(summary_start_row, summary_start_col + i,
                             col, teal_header_format)

                for r in range(len(dim_df)):
                    for c in range(len(dim_df.columns)):
                        ws.write(summary_start_row + 1 + r,
                                 summary_start_col + c,
                                 dim_df.iat[r, c],
                                 text_format)

                for c in range(len(dim_df.columns)):
                    ws.set_column(summary_start_col + c,
                                  summary_start_col + c,
                                  18)

    # -------------------------------------------------------------
    #                  PO SUMMARY FINALIZATION
    # -------------------------------------------------------------
    for r in range(len(po_summary_df)):
        po_num = str(po_summary_df.iloc[r, 0])
        assigned = str(po_summary_df.iloc[r, 1])

        excel_row = r + 2
        row = r + 1

        if assigned == "Orville": fmt = orville_format
        elif assigned == "Stephanie": fmt = stephanie_format
        elif assigned == "Paulo": fmt = paulo_format
        elif assigned == "JB": fmt = jb_format
        elif assigned == "Sunshine": fmt = sunshine_format
        else: fmt = text_format

        ws_po.write(row, 0, po_num, fmt)
        ws_po.write(row, 1, assigned, fmt)

        if po_num in group_sheet_link_locations:
            sheet, glink_row = group_sheet_link_locations[po_num]
            esc = sheet.replace("'", "''")
            ws_po.write_formula(
                row, 2,
                f'=IF(TRIM(\'{esc}\'!F{glink_row})="","",'
                f'HYPERLINK(\'{esc}\'!F{glink_row},\'{esc}\'!F{glink_row}))',
                text_format
            )
        else:
            ws_po.write(row, 2, "", text_format)

        ws_po.write(row, 3, "", text_format)

        status_formula = (
            f'=IF(AND(CELL("contents",C{excel_row})="",D{excel_row}=""),'
            f'"AWAITING UPLOAD",'
            f'IF(AND(CELL("contents",C{excel_row})="",D{excel_row}<>""),'
            f'"WITH ISSUE",'
            f'IF(AND(CELL("contents",C{excel_row})<>"",D{excel_row}<>""),'
            f'"WITH ISSUE","UPLOADED")))'
        )

        ws_po.write_formula(row, 4, status_formula, text_format)

    ws_po.conditional_format(
        1, 4, len(po_summary_df), 4,
        {"type": "text", "criteria": "containing",
         "value": "UPLOADED", "format": uploaded_format}
    )
    ws_po.conditional_format(
        1, 4, len(po_summary_df), 4,
        {"type": "text", "criteria": "containing",
         "value": "WITH ISSUE", "format": with_issue_format}
    )
    ws_po.conditional_format(
        1, 4, len(po_summary_df), 4,
        {"type": "text", "criteria": "containing",
         "value": "AWAITING UPLOAD", "format": awaiting_upload_format}
    )

    ws_po.set_column(0, 0, 30)
    ws_po.set_column(1, 1, 18)
    ws_po.set_column(2, 2, 120)
    ws_po.set_column(3, 3, 30)
    ws_po.set_column(4, 4, 25)

    writer.close()

    st.success("Processing complete!")

    central_tz = pytz.timezone("America/Chicago")
    now = datetime.now(central_tz)
    fname = "SMW Bulk Shipments " + now.strftime("%Y-%m-%d %I-%M-%S %p") + ".xlsx"

    st.download_button(
        label="Download Organized Excel File",
        data=output.getvalue(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
