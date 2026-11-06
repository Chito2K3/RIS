import os
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

SKIP_FILES = {
    "INVENTORY UTILIZATION REPORT_MASTERLIST_2026_V30.xlsx",
    "RSMI.xlsx",
    "RSMI.xls",
}

def is_rsmi_or_master_file(fname, master_filename=None):
    u = fname.upper()
    if fname in SKIP_FILES:
        return True
    if u.startswith("RSMI"):
        return True
    if (master_filename and fname == master_filename) or ("POPULATED" in u):
        return True
    return False

def extract_ris_no(raw_val):
    if not raw_val:
        return ""
    s = str(raw_val).strip()
    s = re.sub(r'^(?:RIS|RSI)\s*No\.?\s*:?\s*', '', s, flags=re.IGNORECASE).strip()
    return s

def _read_ris_file_items(fpath):
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb.active

    office = str(ws.cell(6, 2).value or "").strip()
    raw_ris_no = ws.cell(6, 5).value
    ris_no = extract_ris_no(raw_ris_no)

    rows = []
    item_count = 0
    total_qty = 0

    for r in range(9, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_a_str = str(col_a).strip() if col_a is not None else ""
        if col_a_str.upper() == "TOTAL" or col_a_str.upper().startswith("PURPOSE:"):
            break
        col_b = ws.cell(r, 2).value
        col_c = ws.cell(r, 3).value
        col_d = ws.cell(r, 4).value

        if not col_a_str and not col_c:
            continue

        unit_str = str(col_b).strip() if col_b is not None else ""
        desc_str = str(col_c).strip() if col_c is not None else ""
        qty_val = col_d if col_d is not None else 0
        try:
            qty_num = int(float(qty_val)) if float(qty_val).is_integer() else float(qty_val)
        except Exception:
            qty_num = qty_val

        rows.append((ris_no, col_a_str, unit_str, desc_str, qty_num))
        item_count += 1
        if isinstance(qty_num, (int, float)):
            total_qty += qty_num

    wb.close()
    return office, ris_no, rows, item_count, total_qty

def _write_rsmi_workbook(rsmi_rows, summary_rows, output_xlsx, output_xls=None):
    out_wb = openpyxl.Workbook()
    ws_rsmi = out_wb.active
    ws_rsmi.title = "RSMI"

    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    data_font = Font(name="Calibri", size=11, bold=False)
    bold_font = Font(name="Calibri", size=11, bold=True)

    thin_side = Side(style="thin", color="000000")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    headers = ["RSI No:", "Stock No.", "Unit", "Description", "Quantity"]
    ws_rsmi.row_dimensions[1].height = 24
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws_rsmi.cell(1, col_idx, h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_all

    for row_idx, (ris_no, stock_no, unit, desc, qty) in enumerate(rsmi_rows, start=2):
        c1 = ws_rsmi.cell(row_idx, 1, ris_no if ris_no else None)
        c2 = ws_rsmi.cell(row_idx, 2, stock_no)
        c3 = ws_rsmi.cell(row_idx, 3, unit)
        c4 = ws_rsmi.cell(row_idx, 4, desc)
        c5 = ws_rsmi.cell(row_idx, 5, qty)

        for c in (c1, c2, c3, c4, c5):
            c.font = data_font
            c.border = border_all

        c1.alignment = align_center
        c2.alignment = align_center
        c3.alignment = align_center
        c4.alignment = align_left
        c5.alignment = align_right
        if isinstance(qty, (int, float)):
            c5.number_format = "#,##0"

    ws_rsmi.column_dimensions["A"].width = 28
    ws_rsmi.column_dimensions["B"].width = 16
    ws_rsmi.column_dimensions["C"].width = 14
    ws_rsmi.column_dimensions["D"].width = 58
    ws_rsmi.column_dimensions["E"].width = 14
    ws_rsmi.freeze_panes = "A2"
    if rsmi_rows:
        ws_rsmi.auto_filter.ref = f"A1:E{len(rsmi_rows) + 1}"

    ws_sum = out_wb.create_sheet(title="RIS Count Summary")
    sum_headers = ["RSI No:", "Office / Ward", "Item Count", "Total Quantity", "Source File"]
    ws_sum.row_dimensions[1].height = 24
    for col_idx, h_text in enumerate(sum_headers, start=1):
        cell = ws_sum.cell(1, col_idx, h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_all

    grand_items = 0
    grand_qty = 0
    for idx, srow in enumerate(summary_rows, start=2):
        c1 = ws_sum.cell(idx, 1, srow["ris_no"] if srow["ris_no"] else None)
        c2 = ws_sum.cell(idx, 2, srow["office"])
        c3 = ws_sum.cell(idx, 3, srow["item_count"])
        c4 = ws_sum.cell(idx, 4, srow["total_qty"])
        c5 = ws_sum.cell(idx, 5, srow["filename"])

        for c in (c1, c2, c3, c4, c5):
            c.font = data_font
            c.border = border_all

        c1.alignment = align_center
        c2.alignment = align_left
        c3.alignment = align_right
        c3.number_format = "#,##0"
        c4.alignment = align_right
        c4.number_format = "#,##0"
        c5.alignment = align_left

        grand_items += srow["item_count"]
        grand_qty += srow["total_qty"]

    tot_row = len(summary_rows) + 2
    ws_sum.cell(tot_row, 1, "TOTAL").font = bold_font
    ws_sum.cell(tot_row, 1).alignment = align_center
    ws_sum.cell(tot_row, 2, f"{len(summary_rows)} RIS File(s)").font = bold_font
    ws_sum.cell(tot_row, 2).alignment = align_left
    ws_sum.cell(tot_row, 3, grand_items).font = bold_font
    ws_sum.cell(tot_row, 3).alignment = align_right
    ws_sum.cell(tot_row, 3).number_format = "#,##0"
    ws_sum.cell(tot_row, 4, grand_qty).font = bold_font
    ws_sum.cell(tot_row, 4).alignment = align_right
    ws_sum.cell(tot_row, 4).number_format = "#,##0"
    ws_sum.cell(tot_row, 5, "").font = bold_font
    for col_idx in range(1, 6):
        ws_sum.cell(tot_row, col_idx).border = border_all
        ws_sum.cell(tot_row, col_idx).fill = header_fill

    ws_sum.column_dimensions["A"].width = 28
    ws_sum.column_dimensions["B"].width = 38
    ws_sum.column_dimensions["C"].width = 15
    ws_sum.column_dimensions["D"].width = 18
    ws_sum.column_dimensions["E"].width = 48
    ws_sum.freeze_panes = "A2"

    try:
        out_wb.save(output_xlsx)
        if output_xls:
            out_wb.save(output_xls)
    except PermissionError:
        out_wb.close()
        raise PermissionError(
            f"Cannot overwrite {os.path.basename(output_xlsx)} because it is currently open in Excel or LibreOffice Calc. Please close the file first and try again."
        )
    out_wb.close()
    return grand_qty

def generate_populated_rsmi(final_dir, master_filename=None, month_upper="OCTOBER", year_str="2026"):
    """
    Generates a separate RSMI workbook specifically for `RIS - <MONTH>_<YEAR>_POPULATED.xlsx`,
    named `RSMI-<MONTH>_<YEAR>-Populated.xlsx` (and `.xls`) inside `final_dir`.
    """
    if not master_filename:
        # Auto-detect RIS - *_POPULATED.xlsx in final_dir
        for f in os.listdir(final_dir):
            if "POPULATED" in f.upper() and f.lower().endswith(".xlsx") and not f.startswith("~$") and not f.upper().startswith("RSMI"):
                master_filename = f
                break

    if not master_filename:
        return None

    master_path = os.path.join(final_dir, master_filename)
    if not os.path.exists(master_path):
        return None

    # Extract month_year from master_filename if needed (e.g. RIS - OCTOBER_2026_POPULATED.xlsx)
    m_match = re.search(r'RIS\s*-\s*([A-Z]+)_(\d{4})_POPULATED', master_filename, re.IGNORECASE)
    if m_match:
        month_upper = m_match.group(1).upper()
        year_str = m_match.group(2)

    pop_xlsx_name = f"RSMI-{month_upper}_{year_str}-Populated.xlsx"
    pop_xls_name = f"RSMI-{month_upper}_{year_str}-Populated.xls"
    pop_xlsx_path = os.path.join(final_dir, pop_xlsx_name)
    pop_xls_path = os.path.join(final_dir, pop_xls_name)

    office, ris_no, rows, item_count, total_qty = _read_ris_file_items(master_path)
    summary_rows = [{
        "ris_no": ris_no,
        "office": office,
        "item_count": item_count,
        "total_qty": total_qty,
        "filename": master_filename,
    }]

    _write_rsmi_workbook(rows, summary_rows, pop_xlsx_path, pop_xls_path)

    return {
        "filename": pop_xlsx_name,
        "xls_filename": pop_xls_name,
        "path": pop_xlsx_path,
        "xls_path": pop_xls_path,
        "source_ris_file": master_filename,
        "ris_no": ris_no,
        "office": office,
        "item_count": item_count,
        "total_qty": total_qty,
    }

def generate_rsmi_from_folder(final_dir, master_filename=None, excluded_files=None):
    """
    Scans all generated Per-Ward RIS .xlsx files inside `final_dir` (excluding Master Populated RIS,
    RSMI files, and any files in `excluded_files`), and builds `RSMI.xlsx` + `RSMI.xls`.
    """
    output_xlsx = os.path.join(final_dir, "RSMI.xlsx")
    output_xls = os.path.join(final_dir, "RSMI.xls")
    excluded_set = set(excluded_files or [])

    all_files = sorted([
        f for f in os.listdir(final_dir)
        if os.path.isfile(os.path.join(final_dir, f))
        and f.lower().endswith(".xlsx")
        and not f.startswith("~$")
        and not is_rsmi_or_master_file(f, master_filename)
        and f not in excluded_set
    ])

    rsmi_rows = []
    summary_rows = []

    for fname in all_files:
        fpath = os.path.join(final_dir, fname)
        office, ris_no, rows, item_count, total_qty = _read_ris_file_items(fpath)
        rsmi_rows.extend(rows)
        summary_rows.append({
            "ris_no": ris_no,
            "office": office,
            "item_count": item_count,
            "total_qty": total_qty,
            "filename": fname,
        })

    grand_qty = _write_rsmi_workbook(rsmi_rows, summary_rows, output_xlsx, output_xls)

    return {
        "rsmi_path": output_xlsx,
        "rsmi_xls_path": output_xls,
        "total_ris_files": len(summary_rows),
        "total_medicine_rows": len(rsmi_rows),
        "grand_total_qty": grand_qty,
        "summary_rows": summary_rows,
    }
