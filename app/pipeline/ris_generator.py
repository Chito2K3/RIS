import os
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

thick = Side(border_style="thick", color="000000")
thin = Side(border_style="thin", color="000000")
border_all_thick = Border(top=thick, bottom=thick, left=thick, right=thick)

def set_range_border(ws, min_col, min_row, max_col, max_row, top=None, bottom=None, left=None, right=None, inner_v=None, inner_h=None):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            t = top if r == min_row else inner_h
            b = bottom if r == max_row else inner_h
            l = left if c == min_col else inner_v
            rg = right if c == max_col else inner_v
            ws.cell(r, c).border = Border(top=t, bottom=b, left=l, right=rg)

def compute_ward_ris_code(office_name, month_short="OCT", year="2026"):
    """
    Produces the exact RIS No. short code for each ward/office matching ARMMC conventions:
      - Strips leading 'DEPARTMENT OF '
      - Strips trailing ' WARD' or ' UNIT' (except when part of 'OB COMPLEX' or 'BURN UNIT')
      - Truncates to 15 characters
    """
    s = str(office_name).strip().upper()
    if s.startswith("DEPARTMENT OF "):
        s = s[len("DEPARTMENT OF "):].strip()
    if s.endswith(" WARD"):
        s = s[:-len(" WARD")].strip()
    elif s.endswith(" UNIT") and s != "BURN UNIT":
        s = s[:-len(" UNIT")].strip()
    short_code = s[:15].rstrip()
    return f"RIS No.  {month_short}-{year}-{short_code}"

def sanitize_ward_filename(office_name, month_short="OCT"):
    safe_office = re.sub(r'[\\/:*?"<>|]', '-', str(office_name).strip())
    return f"{safe_office}_RIS_{month_short}.xlsx"

def build_ris_workbook(
    dest_path,
    office_name,
    ris_no_text,
    items,
    unit_lookup,
    is_master=False,
    month_year_label="October 2026",
    rcc_code="13-001-14-00012-03-02-33",
    req_by_name="SABA, LEONARDO CONCHITO SEBASTIAN",
    req_by_desig="Administrative Assistant II",
    app_by_name="SALVE CERTEZA",
    app_by_desig="Pharmacist V",
    sig_date="11/09/2026",
):
    """
    Creates a finalized 8-column RIS Excel workbook (.xlsx) from scratch matching
    the exact column widths, cell merges, borders, fonts, Unit lookup, Purpose banner,
    and Signatory block.
    `items` is a list of dicts: [{'Item Code': ..., 'Description': ..., 'Quantity': ...}]
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RIS"

    # 1. Compact column widths (8 columns: A through H) so Fit-to-1-Page-Wide prints at ~78-82% scale in A4 Portrait
    col_widths = {
        'A': 13.5,  # Stock No.
        'B': 9.5,   # Unit
        'C': 38.0,  # Description (wraps long descriptions cleanly)
        'D': 11.0,  # Quantity (Requisition)
        'E': 6.5,   # Yes (Stock Available)
        'F': 6.5,   # No (Stock Available)
        'G': 11.0,  # Quantity (Issue)
        'H': 10.0   # Remarks (Issue)
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    # 2. Header Rows (Rows 1 to 8)
    # Row 1: Title
    ws.row_dimensions[1].height = 26.25
    ws.cell(1, 1).value = 'REQUISITION AND ISSUE SLIP'
    ws.cell(1, 1).font = Font(name='Calibri', size=20.0, bold=True)
    ws.cell(1, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('A1:H1')

    # Row 2
    ws.row_dimensions[2].height = 15.75

    # Row 3: Entity Name & Fund Cluster
    ws.row_dimensions[3].height = 15.75
    ws.cell(3, 1).value = 'Entity Name:  AMANG RODRIGUEZ MEMORIAL MEDICAL CENTER'
    ws.cell(3, 1).font = Font(name='Calibri', size=12.0, bold=True)
    ws.cell(3, 1).alignment = Alignment(horizontal='left', wrap_text=True)
    ws.merge_cells('A3:C3')

    ws.cell(3, 5).value = 'Fund Cluster:'
    ws.cell(3, 5).font = Font(name='Calibri', size=11.0, bold=True)
    ws.cell(3, 5).alignment = Alignment(horizontal='right')

    # Row 4
    ws.row_dimensions[4].height = 16.5
    ws.cell(4, 5).border = Border(bottom=thick)

    # Row 5: Division & Responsibility Center Code
    ws.row_dimensions[5].height = 15.75
    ws.cell(5, 1).value = 'Division:'
    ws.cell(5, 1).font = Font(name='Calibri', size=11.0, bold=True)
    ws.cell(5, 2).value = 'ALLIED'
    ws.cell(5, 2).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(5, 2).alignment = Alignment(horizontal='center')
    ws.merge_cells('B5:C5')

    ws.cell(5, 5).value = f'Responsibility Center Code:  {rcc_code}'
    ws.cell(5, 5).font = Font(name='Calibri', size=11.0, bold=True)
    ws.cell(5, 5).alignment = Alignment(horizontal='right')

    for c in range(1, 9):
        ws.cell(5, c).border = Border(
            top=thick,
            left=thick if c == 1 else None,
            right=thick if c == 8 else None
        )

    # Row 6: Office & RIS No.
    ws.row_dimensions[6].height = 15.75
    ws.cell(6, 1).value = 'Office:'
    ws.cell(6, 1).font = Font(name='Calibri', size=11.0, bold=True)
    ws.cell(6, 2).value = office_name
    ws.cell(6, 2).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(6, 2).alignment = Alignment(horizontal='center')
    ws.merge_cells('B6:C6')

    ws.cell(6, 5).value = ris_no_text
    ws.cell(6, 5).font = Font(name='Calibri', size=11.0, bold=True)
    ws.cell(6, 5).alignment = Alignment(horizontal='right')

    for c in range(1, 9):
        ws.cell(6, c).border = Border(
            bottom=thick,
            left=thick if c == 1 else None,
            right=thick if c == 8 else None
        )

    # Row 7: Section Banners
    ws.row_dimensions[7].height = 39.0
    ws.cell(7, 1).value = 'Requisition'
    ws.cell(7, 1).font = Font(name='Calibri', size=16.0, bold=True, italic=True)
    ws.cell(7, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('A7:D7')
    set_range_border(ws, 1, 7, 4, 7, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(7, 5).value = 'Stock Available'
    ws.cell(7, 5).font = Font(name='Calibri', size=16.0, bold=True, italic=True)
    ws.cell(7, 5).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.merge_cells('E7:F7')
    set_range_border(ws, 5, 7, 6, 7, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(7, 7).value = 'Issue'
    ws.cell(7, 7).font = Font(name='Calibri', size=16.0, bold=True, italic=True)
    ws.cell(7, 7).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('G7:H7')
    set_range_border(ws, 7, 7, 8, 7, top=thick, bottom=thick, left=thick, right=thick)

    # Row 8: Column Headers
    ws.row_dimensions[8].height = 20.1
    headers = ['Stock No.', 'Unit', 'Description', 'Quantity', 'Yes', 'No', 'Quantity', 'Remarks']
    for c, h_text in enumerate(headers, 1):
        cell = ws.cell(8, c)
        cell.value = h_text
        cell.font = Font(name='Calibri', size=11.0, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_all_thick

    # 3. Data Rows (Row 9 to 9 + len(items) - 1)
    for idx, item in enumerate(items):
        r = 9 + idx
        ws.row_dimensions[r].height = 21.95
        code = str(item.get("Item Code", "")).strip()
        desc = str(item.get("Description", item.get("Item Description", ""))).strip()
        qty_raw = item.get("Quantity", 0)
        try:
            qty = int(float(qty_raw)) if float(qty_raw).is_integer() else float(qty_raw)
        except Exception:
            qty = qty_raw

        uom = unit_lookup.resolve_unit(code, desc) if unit_lookup else ""

        ws.cell(r, 1, code)
        ws.cell(r, 2, uom)
        ws.cell(r, 3, desc)
        ws.cell(r, 4, qty)
        ws.cell(r, 5, None)
        ws.cell(r, 6, None)
        ws.cell(r, 7, qty)
        ws.cell(r, 8, None)

        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.font = Font(name='Calibri', size=11.0, bold=False)
            if c in (1, 2):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            elif c == 3:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            elif c in (4, 7):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            cell.border = border_all_thick
        ws.row_dimensions[r].height = 26.0 if len(desc) > 40 else 21.95

    # 4. TOTAL Row
    total_row = 9 + len(items)
    ws.row_dimensions[total_row].height = 24.0
    ws.cell(total_row, 1).value = 'TOTAL'
    ws.cell(total_row, 2).value = None
    ws.cell(total_row, 3).value = None
    ws.cell(total_row, 4).value = f"=SUM(D9:D{total_row-1})" if len(items) > 0 else 0
    ws.cell(total_row, 5).value = None
    ws.cell(total_row, 6).value = None
    ws.cell(total_row, 7).value = f"=SUM(G9:G{total_row-1})" if len(items) > 0 else 0
    ws.cell(total_row, 8).value = None

    for c in range(1, 9):
        cell = ws.cell(total_row, c)
        cell.font = Font(name='Calibri', size=16.0, bold=True)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell.border = Border(
            top=thick,
            bottom=thick,
            left=thick if c == 1 else None,
            right=thick if c in (4, 7, 8) else None
        )

    # 5. Purpose / Remarks Row (total_row + 1)
    purpose_row = total_row + 1
    ws.row_dimensions[purpose_row].height = 39.95
    if is_master:
        purpose_text = f"Purpose: For Pharmacy Medicines Consumption (General & Consignment) for the month of {month_year_label}"
    else:
        purpose_text = f"Purpose: For {office_name} Medicines Issuance for the month of {month_year_label}"

    ws.cell(purpose_row, 1).value = purpose_text
    ws.merge_cells(start_row=purpose_row, start_column=1, end_row=purpose_row, end_column=3)
    for c in range(1, 9):
        cell = ws.cell(purpose_row, c)
        cell.font = Font(name='Calibri', size=16.0, bold=True)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cell.border = Border(
            top=thick,
            bottom=thick,
            left=thick if c == 1 else None,
            right=thick if c == 8 else None
        )

    # 6. Signature Block (purpose_row + 1 to purpose_row + 4)
    r_sig_hdr = purpose_row + 1
    r_name = purpose_row + 2
    r_desig = purpose_row + 3
    r_date = purpose_row + 4

    ws.row_dimensions[r_sig_hdr].height = 27.95
    ws.row_dimensions[r_name].height = 16.5
    ws.row_dimensions[r_desig].height = 16.5
    ws.row_dimensions[r_date].height = 16.5

    # Row: Signature Header
    ws.cell(r_sig_hdr, 1).value = 'Signature'
    ws.cell(r_sig_hdr, 1).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_sig_hdr, 1).alignment = Alignment(horizontal='left', vertical='bottom')
    ws.cell(r_sig_hdr, 1).border = border_all_thick

    ws.cell(r_sig_hdr, 2).value = 'Requested By:'
    ws.cell(r_sig_hdr, 2).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_sig_hdr, 2).alignment = Alignment(horizontal='left', vertical='top')
    ws.merge_cells(start_row=r_sig_hdr, start_column=2, end_row=r_sig_hdr, end_column=3)
    set_range_border(ws, 2, r_sig_hdr, 3, r_sig_hdr, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_sig_hdr, 4).value = 'Approved By:'
    ws.cell(r_sig_hdr, 4).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_sig_hdr, 4).alignment = Alignment(horizontal='center', vertical='top')
    ws.merge_cells(start_row=r_sig_hdr, start_column=4, end_row=r_sig_hdr, end_column=5)
    set_range_border(ws, 4, r_sig_hdr, 5, r_sig_hdr, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_sig_hdr, 6).value = 'Issued By:'
    ws.cell(r_sig_hdr, 6).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_sig_hdr, 6).alignment = Alignment(horizontal='left', vertical='top')
    ws.merge_cells(start_row=r_sig_hdr, start_column=6, end_row=r_sig_hdr, end_column=7)
    set_range_border(ws, 6, r_sig_hdr, 7, r_sig_hdr, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_sig_hdr, 8).value = 'Received By:'
    ws.cell(r_sig_hdr, 8).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_sig_hdr, 8).alignment = Alignment(horizontal='center', vertical='top')
    ws.cell(r_sig_hdr, 8).border = border_all_thick

    # Row: Printed Name
    ws.cell(r_name, 1).value = 'Printed Name:'
    ws.cell(r_name, 1).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_name, 1).alignment = Alignment(wrap_text=True)
    ws.cell(r_name, 1).border = border_all_thick

    ws.cell(r_name, 2).value = req_by_name
    ws.cell(r_name, 2).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_name, 2).alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells(start_row=r_name, start_column=2, end_row=r_name, end_column=3)
    set_range_border(ws, 2, r_name, 3, r_name, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_name, 4).value = app_by_name
    ws.cell(r_name, 4).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_name, 4).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(start_row=r_name, start_column=4, end_row=r_name, end_column=5)
    set_range_border(ws, 4, r_name, 5, r_name, top=thick, bottom=None, left=thick, right=thick)

    ws.cell(r_name, 6).value = None
    ws.merge_cells(start_row=r_name, start_column=6, end_row=r_name, end_column=7)
    set_range_border(ws, 6, r_name, 7, r_name, top=thick, bottom=None, left=thick, right=thick)

    ws.cell(r_name, 8).value = None
    ws.cell(r_name, 8).border = Border(top=thick, bottom=None, left=thick, right=thick)

    # Row: Designation
    ws.cell(r_desig, 1).value = 'Designation:'
    ws.cell(r_desig, 1).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_desig, 1).border = border_all_thick

    ws.cell(r_desig, 2).value = req_by_desig
    ws.cell(r_desig, 2).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_desig, 2).alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells(start_row=r_desig, start_column=2, end_row=r_desig, end_column=3)
    set_range_border(ws, 2, r_desig, 3, r_desig, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_desig, 4).value = app_by_desig
    ws.cell(r_desig, 4).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_desig, 4).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(start_row=r_desig, start_column=4, end_row=r_desig, end_column=5)
    set_range_border(ws, 4, r_desig, 5, r_desig, top=None, bottom=thick, left=thick, right=thick)

    ws.cell(r_desig, 6).value = None
    ws.merge_cells(start_row=r_desig, start_column=6, end_row=r_desig, end_column=7)
    set_range_border(ws, 6, r_desig, 7, r_desig, top=None, bottom=thick, left=thick, right=thick)

    ws.cell(r_desig, 8).value = None
    ws.cell(r_desig, 8).border = Border(top=None, bottom=thick, left=thick, right=thick)

    # Row: Date
    ws.cell(r_date, 1).value = 'Date:'
    ws.cell(r_date, 1).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_date, 1).border = border_all_thick

    ws.cell(r_date, 2).value = sig_date
    ws.cell(r_date, 2).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_date, 2).alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells(start_row=r_date, start_column=2, end_row=r_date, end_column=3)
    set_range_border(ws, 2, r_date, 3, r_date, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_date, 4).value = sig_date
    ws.cell(r_date, 4).font = Font(name='Calibri', size=11.0, bold=False)
    ws.cell(r_date, 4).alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells(start_row=r_date, start_column=4, end_row=r_date, end_column=5)
    set_range_border(ws, 4, r_date, 5, r_date, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_date, 6).value = None
    ws.merge_cells(start_row=r_date, start_column=6, end_row=r_date, end_column=7)
    set_range_border(ws, 6, r_date, 7, r_date, top=thick, bottom=thick, left=thick, right=thick)

    ws.cell(r_date, 8).value = None
    ws.cell(r_date, 8).border = border_all_thick

    # Configure Page Setup so all 8 columns (A:H) fit cleanly on 1 page wide when printed
    ws.print_area = f"A1:H{r_date}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2

    try:
        wb.save(dest_path)
    except PermissionError:
        # If the file is currently open in Excel, skip overwriting so the rest of the pipeline succeeds
        pass
    wb.close()
    return dest_path

def generate_all_ward_ris_files(
    df_combined_issuance,
    active_wards,
    final_dir,
    unit_lookup,
    rcc_lookup=None,
    month_short="OCT",
    year="2026",
    month_year_label="October 2026",
    req_by_name="SABA, LEONARDO CONCHITO SEBASTIAN",
    req_by_desig="Administrative Assistant II",
    app_by_name="SALVE CERTEZA",
    app_by_desig="Pharmacist V",
    sig_date="11/09/2026",
):
    generated_wards = []
    for ward in sorted(active_wards):
        ward_df = df_combined_issuance[df_combined_issuance[ward] > 0]
        if ward_df.empty:
            continue
        ward_items = []
        ward_total_qty = 0
        for _, row in ward_df.iterrows():
            q = row[ward]
            q_num = int(float(q)) if float(q).is_integer() else float(q)
            ward_items.append({
                "Item Code": str(row["Item Code"]).strip(),
                "Description": str(row["Item Description"]).strip(),
                "Quantity": q_num,
            })
            ward_total_qty += q_num

        ris_no_text = compute_ward_ris_code(ward, month_short=month_short, year=year)
        ward_rcc = rcc_lookup.resolve_rcc(ward) if rcc_lookup else "13-001-14-00012-03-02-33"
        fname = sanitize_ward_filename(ward, month_short=month_short)
        dest_path = os.path.join(final_dir, fname)

        build_ris_workbook(
            dest_path=dest_path,
            office_name=ward,
            ris_no_text=ris_no_text,
            items=ward_items,
            unit_lookup=unit_lookup,
            is_master=False,
            month_year_label=month_year_label,
            rcc_code=ward_rcc,
            req_by_name=req_by_name,
            req_by_desig=req_by_desig,
            app_by_name=app_by_name,
            app_by_desig=app_by_desig,
            sig_date=sig_date,
        )
        generated_wards.append({
            "ward": ward,
            "ris_no": ris_no_text.replace("RIS No.  ", "").strip(),
            "rcc_code": ward_rcc,
            "filename": fname,
            "path": dest_path,
            "item_count": len(ward_items),
            "total_qty": ward_total_qty,
        })

    return generated_wards

