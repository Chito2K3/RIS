import os
import re
import math
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break

thin = Side(border_style="thin", color="000000")
border_all_thin = Border(top=thin, bottom=thin, left=thin, right=thin)

RIS_ITEMS_PER_PAGE = 22

def _build_chunked_sum_formula(col_letter, row_ranges, chunk_size=20):
    """
    Builds an Excel SUM formula across multiple (start_row, end_row) tuples.
    Splits into chunks of at most `chunk_size` arguments per SUM() call
    so it remains 100% compatible with both .xlsx and legacy .xls (BIFF8 30-arg limit).
    """
    if not row_ranges:
        return 0
    range_strs = [f"{col_letter}{s}:{col_letter}{e}" for s, e in row_ranges]
    if len(range_strs) <= chunk_size:
        return f"=SUM({','.join(range_strs)})"
    parts = []
    for i in range(0, len(range_strs), chunk_size):
        chunk = range_strs[i:i + chunk_size]
        parts.append(f"SUM({','.join(chunk)})")
    return "=" + "+".join(parts)

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

    # 1. Column widths (8 columns: A through H) with balanced widths for Signatory columns (Requested By, Approved By, Issued By, Received By)
    col_widths = {
        'A': 14.0,  # Stock No. / Signatory Row Labels
        'B': 10.5,  # Unit / Requested By (left)
        'C': 34.0,  # Description / Requested By (right) -> B+C = 44.5
        'D': 12.0,  # Quantity (Requisition) / Approved By (left)
        'E': 10.0,  # Yes (Stock Available) / Approved By (right) -> D+E = 22.0
        'F': 10.0,  # No (Stock Available) / Issued By (left)
        'G': 12.0,  # Quantity (Issue) / Issued By (right) -> F+G = 22.0
        'H': 22.0   # Remarks (Issue) / Received By -> H = 22.0
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    # Split items into page chunks so every printed page has its own Header, 22-row grid, Total/Continuation, and Signatory block
    page_chunks = []
    if items:
        for i in range(0, len(items), RIS_ITEMS_PER_PAGE):
            page_chunks.append(items[i:i + RIS_ITEMS_PER_PAGE])
    else:
        page_chunks.append([])

    total_pages = len(page_chunks)
    data_row_ranges = []
    cur_row = 1

    for page_idx, page_items in enumerate(page_chunks):
        is_last_page = (page_idx == total_pages - 1)
        hdr_base = cur_row

        # 2. Header Rows (8 rows per page: hdr_base to hdr_base + 7)
        r1 = hdr_base
        r2 = hdr_base + 1
        r3 = hdr_base + 2
        r4 = hdr_base + 3
        r5 = hdr_base + 4
        r6 = hdr_base + 5
        r7 = hdr_base + 6
        r8 = hdr_base + 7

        # Row +0: Title
        ws.row_dimensions[r1].height = 26.25
        ws.cell(r1, 1).value = 'REQUISITION AND ISSUE SLIP'
        ws.cell(r1, 1).font = Font(name='Calibri', size=20.0, bold=True)
        ws.cell(r1, 1).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r1, start_column=1, end_row=r1, end_column=8)

        # Row +1: Spacer
        ws.row_dimensions[r2].height = 10.0

        # Row +2: Entity Name & Fund Cluster (merged cleanly so neither gets clipped)
        ws.row_dimensions[r3].height = 18.0
        ws.cell(r3, 1).value = 'Entity Name:  AMANG RODRIGUEZ MEMORIAL MEDICAL CENTER'
        ws.cell(r3, 1).font = Font(name='Calibri', size=11.5, bold=True)
        ws.cell(r3, 1).alignment = Alignment(horizontal='left', vertical='center')
        ws.merge_cells(start_row=r3, start_column=1, end_row=r3, end_column=5)

        ws.cell(r3, 6).value = 'Fund Cluster:'
        ws.cell(r3, 6).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r3, 6).alignment = Alignment(horizontal='right', vertical='center')
        ws.merge_cells(start_row=r3, start_column=6, end_row=r3, end_column=7)
        ws.cell(r3, 8).border = Border(bottom=thin)

        # Row +3: Spacer
        ws.row_dimensions[r4].height = 8.0

        # Row +4: Division & Responsibility Center Code (D5:H5 merged so RCC code never gets cut off)
        ws.row_dimensions[r5].height = 18.0
        ws.cell(r5, 1).value = 'Division:'
        ws.cell(r5, 1).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r5, 1).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(r5, 2).value = 'ALLIED'
        ws.cell(r5, 2).font = Font(name='Calibri', size=11.0, bold=False)
        ws.cell(r5, 2).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r5, start_column=2, end_row=r5, end_column=3)

        ws.cell(r5, 4).value = f'Responsibility Center Code:  {rcc_code}'
        ws.cell(r5, 4).font = Font(name='Calibri', size=10.5, bold=True)
        ws.cell(r5, 4).alignment = Alignment(horizontal='left', vertical='center')
        ws.merge_cells(start_row=r5, start_column=4, end_row=r5, end_column=8)

        for c in range(1, 9):
            ws.cell(r5, c).border = Border(
                top=thin,
                left=thin if c == 1 else None,
                right=thin if c == 8 else None
            )

        # Row +5: Office & RIS No. (D6:H6 merged so RIS No. never gets cut off)
        ws.row_dimensions[r6].height = 18.0
        ws.cell(r6, 1).value = 'Office:'
        ws.cell(r6, 1).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r6, 1).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(r6, 2).value = office_name
        ws.cell(r6, 2).font = Font(name='Calibri', size=11.0, bold=False)
        ws.cell(r6, 2).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r6, start_column=2, end_row=r6, end_column=3)

        ws.cell(r6, 4).value = ris_no_text
        ws.cell(r6, 4).font = Font(name='Calibri', size=10.5, bold=True)
        ws.cell(r6, 4).alignment = Alignment(horizontal='left', vertical='center')
        ws.merge_cells(start_row=r6, start_column=4, end_row=r6, end_column=8)

        for c in range(1, 9):
            ws.cell(r6, c).border = Border(
                bottom=thin,
                left=thin if c == 1 else None,
                right=thin if c == 8 else None
            )

        # Row +6: Section Banners
        ws.row_dimensions[r7].height = 30.0
        ws.cell(r7, 1).value = 'Requisition'
        ws.cell(r7, 1).font = Font(name='Calibri', size=15.0, bold=True, italic=True)
        ws.cell(r7, 1).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r7, start_column=1, end_row=r7, end_column=4)
        set_range_border(ws, 1, r7, 4, r7, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r7, 5).value = 'Stock Available'
        ws.cell(r7, 5).font = Font(name='Calibri', size=13.5, bold=True, italic=True)
        ws.cell(r7, 5).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r7, start_column=5, end_row=r7, end_column=6)
        set_range_border(ws, 5, r7, 6, r7, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r7, 7).value = 'Issue'
        ws.cell(r7, 7).font = Font(name='Calibri', size=15.0, bold=True, italic=True)
        ws.cell(r7, 7).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r7, start_column=7, end_row=r7, end_column=8)
        set_range_border(ws, 7, r7, 8, r7, top=thin, bottom=thin, left=thin, right=thin)

        # Row +7: Column Headers
        ws.row_dimensions[r8].height = 20.0
        headers = ['Stock No.', 'Unit', 'Description', 'Quantity', 'Yes', 'No', 'Quantity', 'Remarks']
        for c, h_text in enumerate(headers, 1):
            cell = ws.cell(r8, c)
            cell.value = h_text
            cell.font = Font(name='Calibri', size=11.0, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border_all_thin

        # 3. Data Rows for this page (Always standardized to RIS_ITEMS_PER_PAGE = 22 rows per page)
        data_start_r = r8 + 1
        for slot_idx in range(RIS_ITEMS_PER_PAGE):
            r = data_start_r + slot_idx
            if slot_idx < len(page_items):
                item = page_items[slot_idx]
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

                # Dynamic row height with generous vertical padding so wrapped lines never look cramped
                est_lines = max(1, math.ceil(len(desc) / 32.0)) if desc else 1
                if est_lines == 1:
                    ws.row_dimensions[r].height = 22.0
                elif est_lines == 2:
                    ws.row_dimensions[r].height = 30.0
                else:
                    ws.row_dimensions[r].height = 40.0
            else:
                # Blank bordered filler row to standardize page height up to the bottom of the page
                for c in range(1, 9):
                    ws.cell(r, c, None)
                ws.row_dimensions[r].height = 22.0

            for c in range(1, 9):
                cell = ws.cell(r, c)
                cell.font = Font(name='Calibri', size=10.5, bold=False)
                if c in (1, 2):
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                elif c == 3:
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                elif c in (4, 7):
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                else:
                    cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.border = border_all_thin

        data_end_r = data_start_r + RIS_ITEMS_PER_PAGE - 1
        data_row_ranges.append((data_start_r, data_end_r))
        total_row = data_end_r + 1

        # 4. TOTAL / Continuation Row
        ws.row_dimensions[total_row].height = 24.0
        ws.cell(total_row, 1).value = 'TOTAL'
        ws.cell(total_row, 2).value = None
        if is_last_page:
            ws.cell(total_row, 3).value = None
            ws.cell(total_row, 4).value = _build_chunked_sum_formula("D", data_row_ranges) if data_row_ranges else 0
            ws.cell(total_row, 5).value = None
            ws.cell(total_row, 6).value = None
            ws.cell(total_row, 7).value = _build_chunked_sum_formula("G", data_row_ranges) if data_row_ranges else 0
            ws.cell(total_row, 8).value = None
        else:
            ws.cell(total_row, 3).value = 'Continue to the next page...'
            ws.cell(total_row, 4).value = None
            ws.cell(total_row, 5).value = None
            ws.cell(total_row, 6).value = None
            ws.cell(total_row, 7).value = None
            ws.cell(total_row, 8).value = None

        for c in range(1, 9):
            cell = ws.cell(total_row, c)
            if not is_last_page and c == 3:
                cell.font = Font(name='Calibri', size=11.0, bold=True, italic=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            elif c in (4, 7):
                cell.font = Font(name='Calibri', size=14.0, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.font = Font(name='Calibri', size=14.0, bold=True)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cell.border = Border(
                top=thin,
                bottom=thin,
                left=thin if c in (1, 4, 7) else None,
                right=thin if c in (3, 4, 6, 7, 8) else None
            )

        # 5. Purpose / Remarks Row (total_row + 1)
        purpose_row = total_row + 1
        ws.row_dimensions[purpose_row].height = 32.0
        if is_master:
            purpose_text = f"Purpose: For Pharmacy Medicines Consumption (General & Consignment) for the month of {month_year_label}"
        else:
            purpose_text = f"Purpose: For {office_name} Medicines Issuance for the month of {month_year_label}"

        ws.cell(purpose_row, 1).value = purpose_text
        ws.merge_cells(start_row=purpose_row, start_column=1, end_row=purpose_row, end_column=8)
        for c in range(1, 9):
            cell = ws.cell(purpose_row, c)
            cell.font = Font(name='Calibri', size=12.0, bold=True)
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            cell.border = Border(
                top=thin,
                bottom=thin,
                left=thin if c == 1 else None,
                right=thin if c == 8 else None
            )

        # 6. Signature Block (purpose_row + 1 to purpose_row + 4)
        r_sig_hdr = purpose_row + 1
        r_name = purpose_row + 2
        r_desig = purpose_row + 3
        r_date = purpose_row + 4

        # Generous space in r_sig_hdr (45pt) to sign over the printed name, and equal spacious height (22pt) for Name, Designation, and Date
        ws.row_dimensions[r_sig_hdr].height = 45.0
        ws.row_dimensions[r_name].height = 22.0
        ws.row_dimensions[r_desig].height = 22.0
        ws.row_dimensions[r_date].height = 22.0

        # Row: Signature Header (with ample blank space below "Requested By:", "Approved By:", etc. for wet signature)
        ws.cell(r_sig_hdr, 1).value = 'Signature:'
        ws.cell(r_sig_hdr, 1).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_sig_hdr, 1).alignment = Alignment(horizontal='left', vertical='bottom')
        ws.cell(r_sig_hdr, 1).border = border_all_thin

        ws.cell(r_sig_hdr, 2).value = 'Requested By:'
        ws.cell(r_sig_hdr, 2).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_sig_hdr, 2).alignment = Alignment(horizontal='left', vertical='top')
        ws.merge_cells(start_row=r_sig_hdr, start_column=2, end_row=r_sig_hdr, end_column=3)
        set_range_border(ws, 2, r_sig_hdr, 3, r_sig_hdr, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_sig_hdr, 4).value = 'Approved By:'
        ws.cell(r_sig_hdr, 4).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_sig_hdr, 4).alignment = Alignment(horizontal='left', vertical='top')
        ws.merge_cells(start_row=r_sig_hdr, start_column=4, end_row=r_sig_hdr, end_column=5)
        set_range_border(ws, 4, r_sig_hdr, 5, r_sig_hdr, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_sig_hdr, 6).value = 'Issued By:'
        ws.cell(r_sig_hdr, 6).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_sig_hdr, 6).alignment = Alignment(horizontal='left', vertical='top')
        ws.merge_cells(start_row=r_sig_hdr, start_column=6, end_row=r_sig_hdr, end_column=7)
        set_range_border(ws, 6, r_sig_hdr, 7, r_sig_hdr, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_sig_hdr, 8).value = 'Received By:'
        ws.cell(r_sig_hdr, 8).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_sig_hdr, 8).alignment = Alignment(horizontal='left', vertical='top')
        ws.cell(r_sig_hdr, 8).border = border_all_thin

        # Row: Printed Name
        ws.cell(r_name, 1).value = 'Printed Name:'
        ws.cell(r_name, 1).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_name, 1).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.cell(r_name, 1).border = border_all_thin

        ws.cell(r_name, 2).value = req_by_name
        ws.cell(r_name, 2).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_name, 2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r_name, start_column=2, end_row=r_name, end_column=3)
        set_range_border(ws, 2, r_name, 3, r_name, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_name, 4).value = app_by_name
        ws.cell(r_name, 4).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_name, 4).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r_name, start_column=4, end_row=r_name, end_column=5)
        set_range_border(ws, 4, r_name, 5, r_name, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_name, 6).value = None
        ws.merge_cells(start_row=r_name, start_column=6, end_row=r_name, end_column=7)
        set_range_border(ws, 6, r_name, 7, r_name, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_name, 8).value = None
        ws.cell(r_name, 8).border = border_all_thin

        # Row: Designation
        ws.cell(r_desig, 1).value = 'Designation:'
        ws.cell(r_desig, 1).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_desig, 1).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(r_desig, 1).border = border_all_thin

        ws.cell(r_desig, 2).value = req_by_desig
        ws.cell(r_desig, 2).font = Font(name='Calibri', size=11.0, bold=False)
        ws.cell(r_desig, 2).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r_desig, start_column=2, end_row=r_desig, end_column=3)
        set_range_border(ws, 2, r_desig, 3, r_desig, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_desig, 4).value = app_by_desig
        ws.cell(r_desig, 4).font = Font(name='Calibri', size=11.0, bold=False)
        ws.cell(r_desig, 4).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r_desig, start_column=4, end_row=r_desig, end_column=5)
        set_range_border(ws, 4, r_desig, 5, r_desig, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_desig, 6).value = None
        ws.merge_cells(start_row=r_desig, start_column=6, end_row=r_desig, end_column=7)
        set_range_border(ws, 6, r_desig, 7, r_desig, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_desig, 8).value = None
        ws.cell(r_desig, 8).border = border_all_thin

        # Row: Date
        ws.cell(r_date, 1).value = 'Date:'
        ws.cell(r_date, 1).font = Font(name='Calibri', size=11.0, bold=True)
        ws.cell(r_date, 1).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(r_date, 1).border = border_all_thin

        ws.cell(r_date, 2).value = sig_date
        ws.cell(r_date, 2).font = Font(name='Calibri', size=11.0, bold=False)
        ws.cell(r_date, 2).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r_date, start_column=2, end_row=r_date, end_column=3)
        set_range_border(ws, 2, r_date, 3, r_date, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_date, 4).value = sig_date
        ws.cell(r_date, 4).font = Font(name='Calibri', size=11.0, bold=False)
        ws.cell(r_date, 4).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r_date, start_column=4, end_row=r_date, end_column=5)
        set_range_border(ws, 4, r_date, 5, r_date, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_date, 6).value = None
        ws.merge_cells(start_row=r_date, start_column=6, end_row=r_date, end_column=7)
        set_range_border(ws, 6, r_date, 7, r_date, top=thin, bottom=thin, left=thin, right=thin)

        ws.cell(r_date, 8).value = None
        ws.cell(r_date, 8).border = border_all_thin

        if not is_last_page:
            ws.row_breaks.append(Break(id=r_date))

        cur_row = r_date + 1

    last_sheet_row = cur_row - 1

    # Configure Page Setup so all 8 columns (A:H) fit cleanly on 1 page wide when printed
    ws.print_area = f"A1:H{last_sheet_row}"
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

