import os
import re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from .rcc_lookup import RccLookup

SKIP_FILES = {
    "INVENTORY UTILIZATION REPORT_MASTERLIST_2026_V30.xlsx",
    "RSMI.xlsx",
    "RSMI.xls",
}

DEFAULT_RCC = "13-001-14-00012-03-02-33"
DEFAULT_CUSTODIAN = "SALVE CERTEZA"
_RCC_LOOKUP = RccLookup()

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

def extract_rcc(raw_val):
    if not raw_val:
        return DEFAULT_RCC
    s = str(raw_val).strip()
    s = re.sub(r'^Responsibility\s*Center\s*Code\s*:?\s*', '', s, flags=re.IGNORECASE).strip()
    return s if s else DEFAULT_RCC

def _read_ris_file_items(fpath, rcc_lookup=None):
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb.active

    office = str(ws.cell(6, 2).value or "").strip()
    raw_ris_no = ws.cell(6, 5).value
    ris_no = extract_ris_no(raw_ris_no)
    lookup = rcc_lookup or _RCC_LOOKUP
    if office and lookup:
        rcc_code = lookup.resolve_rcc(office)
    else:
        raw_rcc = ws.cell(5, 5).value
        rcc_code = extract_rcc(raw_rcc)

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

        rows.append((ris_no, rcc_code, col_a_str, desc_str, unit_str, qty_num))
        item_count += 1
        if isinstance(qty_num, (int, float)):
            total_qty += qty_num

    wb.close()
    return office, ris_no, rcc_code, rows, item_count, total_qty

def _build_recapitulation(rsmi_rows):
    """
    Aggregates rsmi_rows by Stock No. preserving first-appearance order.
    Returns list of tuples: (stock_no, total_qty, desc, unit)
    """
    ordered_codes = []
    recap_map = {}
    for (ris_no, rcc_code, stock_no, desc, unit, qty) in rsmi_rows:
        key = str(stock_no).strip()
        if not key:
            continue
        q_val = qty if isinstance(qty, (int, float)) else 0
        if key not in recap_map:
            ordered_codes.append(key)
            recap_map[key] = {
                "stock_no": key,
                "qty": q_val,
                "desc": desc,
                "unit": unit,
            }
        else:
            recap_map[key]["qty"] += q_val
            if not recap_map[key]["unit"] and unit:
                recap_map[key]["unit"] = unit
            if not recap_map[key]["desc"] and desc:
                recap_map[key]["desc"] = desc

    recap_rows = []
    for code in ordered_codes:
        info = recap_map[code]
        q = info["qty"]
        q_int = int(q) if isinstance(q, float) and q.is_integer() else q
        recap_rows.append((info["stock_no"], q_int, info["desc"], info["unit"]))
    return recap_rows

def _write_rsmi_workbook(rsmi_rows, summary_rows, output_xlsx, output_xls=None, custodian_name=DEFAULT_CUSTODIAN):
    """
    Generates the official ARMMC `RSMI FORMAT.xlsx` layout:
      - Sheet 1 ("RSMI"):
        * Rows 1–11: Official ARMMC Report of Supplies and Materials Issued header
        * Rows 12..M: Supply/Property Division Main Table (Cols B..I populated: RIS No., RCC, Stock No., Item, Unit, Quantity Issued; Col J Unit Cost blank; Col K Amount formula `=+J{r}*I{r}`)
        * Row M+1: Main Table Total Row (`=SUM(I12:I{M})` and `=SUM(K12:K{M})`)
        * Rows M+2..M+3: Recapitulation Headers
        * Rows M+4..R: Recapitulation Rows (Cols C..H populated: Stock No., Quantity, Item, Unit; Col I Unit Cost blank; Col J Total Cost formula `=+I{r}*E{r}`; Col K UACS Object Code blank)
        * Row R+1: Recapitulation Total Row (`=SUM(E...)`, variance `=+I{main_tot}-E{recap_tot}`, `=SUM(J...)`, variance `=+K{main_tot}-J{recap_tot}`)
        * Rows R+2..R+9: Certification & Signatory Block with Supply/Property Custodian Name in C{R+5}:G{R+5}
      - Sheet 2 ("RIS Count Summary"):
        * Per-RIS summary table for quick checking
    """
    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "RSMI"

    # Compact column widths so Fit-to-1-Page-Wide prints at ~75-80% scale in A4 Portrait (much larger & easier to read)
    col_widths = {
        "A": 1.5,
        "B": 13.5,   # RIS No. (wraps if long)
        "C": 9.5,    # Responsibility Center Code (left half) / Recap Stock No. (left half)
        "D": 9.5,    # Responsibility Center Code (right half) / Recap Stock No. (right half)
        "E": 12.5,   # Stock No. / Recap Quantity
        "F": 15.0,   # Item (left half)
        "G": 19.0,   # Item (right half) -> F+G = 34 chars wide, wraps long descriptions cleanly
        "H": 9.0,    # Unit
        "I": 9.5,    # Quantity Issued / Recap Unit Cost
        "J": 11.0,   # Unit Cost / Recap Total Cost
        "K": 11.5,   # Amount / Recap UACS Object Code
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    thin = Side(style="thin", color="000000")
    medium = Side(style="medium", color="000000")

    acct_fmt = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'
    int_fmt = '#,##0'
    var_int_fmt = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'

    # --- TOP HEADER (Rows 1 to 11) ---
    ws.row_dimensions[2].height = 21.0
    ws.merge_cells("B2:K2")
    ws.cell(2, 2, "REPORT OF SUPPLIES AND MATERIALS ISSUED")
    ws.cell(2, 2).font = Font(name="Calibri", size=16.0, bold=True)
    ws.cell(2, 2).alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[3].height = 15.5
    ws.merge_cells("E3:F3")
    ws.cell(3, 5, "Kind of Supplies: ")
    ws.cell(3, 5).font = Font(name="Calibri", size=12.0, bold=False)
    ws.cell(3, 5).alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells("G3:I3")
    for c in range(7, 10):
        ws.cell(3, c).border = Border(bottom=thin)

    ws.merge_cells("B4:K4")
    ws.merge_cells("B5:K5")

    # Row 6: Entity Name & Serial No.
    ws.row_dimensions[6].height = 15.5
    ws.merge_cells("B6:C6")
    ws.cell(6, 2, "Entity Name:")
    ws.cell(6, 2).font = Font(name="Calibri", size=12.0, bold=True)
    ws.cell(6, 2).alignment = Alignment(horizontal="left", vertical="center")

    ws.cell(6, 4, "'AMANG' RODRIGUEZ MEMORIAL MEDICAL CENTER")
    ws.cell(6, 4).font = Font(name="Calibri", size=12.0, bold=True)
    ws.cell(6, 4).alignment = Alignment(horizontal="left", vertical="center")

    ws.cell(6, 10, "Serial No.:")
    ws.cell(6, 10).font = Font(name="Calibri", size=12.0, bold=True)
    ws.cell(6, 10).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(6, 11).border = Border(bottom=thin)

    # Row 7: Fund Cluster & Date
    ws.row_dimensions[7].height = 16.5
    ws.merge_cells("B7:C7")
    ws.cell(7, 2, "Fund Cluster:")
    ws.cell(7, 2).font = Font(name="Calibri", size=12.0, bold=True)
    ws.cell(7, 2).alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("D7:I7")
    ws.cell(7, 4, "Internally Generated Fund")
    ws.cell(7, 4).font = Font(name="Calibri", size=12.0, bold=True)
    ws.cell(7, 4).alignment = Alignment(horizontal="center", vertical="center")
    for c in range(4, 10):
        ws.cell(7, c).border = Border(top=thin, bottom=thin)

    ws.cell(7, 10, "           Date:")
    ws.cell(7, 10).font = Font(name="Calibri", size=12.0, bold=True)
    ws.cell(7, 10).alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(7, 11).border = Border(top=thin, bottom=thin)

    # Row 8: Spacer
    ws.row_dimensions[8].height = 16.0
    ws.merge_cells("B8:K8")
    ws.cell(8, 2, " ")

    # Row 9: Division Sub-banners
    ws.row_dimensions[9].height = 16.5
    ws.merge_cells("B9:I9")
    ws.cell(9, 2, "To be be filled up by the Supply and/or Property Division/Unit")
    ws.cell(9, 2).font = Font(name="Calibri", size=10.0, italic=True)
    ws.cell(9, 2).alignment = Alignment(horizontal="center", vertical="center")
    for c in range(2, 10):
        ws.cell(9, c).border = Border(
            top=medium,
            bottom=medium,
            left=medium if c == 2 else None,
            right=medium if c == 9 else None,
        )

    ws.merge_cells("J9:K9")
    ws.cell(9, 10, "To be filled up by the Accounting Division/Unit")
    ws.cell(9, 10).font = Font(name="Calibri", size=9.5, italic=True)
    ws.cell(9, 10).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(9, 10).border = Border(top=medium, bottom=medium, left=medium)
    ws.cell(9, 11).border = Border(top=medium, bottom=medium, right=medium)

    # Rows 10 & 11: Main Table Column Headers
    ws.row_dimensions[10].height = 16.0
    ws.row_dimensions[11].height = 16.0

    ws.merge_cells("B10:B11")
    ws.cell(10, 2, "RIS No.")

    ws.merge_cells("C10:D10")
    ws.cell(10, 3, "Responsibility")
    ws.merge_cells("C11:D11")
    ws.cell(11, 3, "Center Code")

    ws.cell(10, 5, "Stock ")
    ws.cell(11, 5, "No.")

    ws.merge_cells("F10:G11")
    ws.cell(10, 6, "Item")

    ws.merge_cells("H10:H11")
    ws.cell(10, 8, "Unit")

    ws.cell(10, 9, "Quantity")
    ws.cell(11, 9, "Issued")

    ws.merge_cells("J10:J11")
    ws.cell(10, 10, "Unit Cost")

    ws.merge_cells("K10:K11")
    ws.cell(10, 11, "Amount")

    for r_hdr in (10, 11):
        for c in range(2, 12):
            cell = ws.cell(r_hdr, c)
            cell.font = Font(name="Calibri", size=11.0, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(
                top=medium if r_hdr == 10 else None,
                bottom=medium if r_hdr == 11 else None,
                left=medium if c in (2, 10) else (thin if c in (3, 5, 6, 8, 9, 11) else None),
                right=medium if c in (9, 11) else (thin if c in (2, 4, 5, 7, 8, 10) else None),
            )

    # --- MAIN TABLE DATA ROWS (Starting at Row 12) ---
    start_data_row = 12
    effective_rows = rsmi_rows if rsmi_rows else [("", DEFAULT_RCC, "", "", "", 0)]

    for idx, (ris_no, rcc_code, stock_no, desc, unit, qty) in enumerate(effective_rows):
        r = start_data_row + idx
        desc_len = len(str(desc or ""))
        ws.row_dimensions[r].height = 28.0 if desc_len > 36 else 22.0

        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

        c_ris = ws.cell(r, 2, ris_no if ris_no else None)
        c_rcc = ws.cell(r, 3, rcc_code if stock_no else None)
        c_stk = ws.cell(r, 5, stock_no if stock_no else None)
        c_itm = ws.cell(r, 6, desc if desc else None)
        c_unt = ws.cell(r, 8, unit if unit else None)
        c_qty = ws.cell(r, 9, qty if stock_no else None)
        c_uc = ws.cell(r, 10, None)  # Left blank for Accounting Division
        c_amt = ws.cell(r, 11, f"=+J{r}*I{r}")

        c_ris.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_rcc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_stk.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_itm.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_unt.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_qty.alignment = Alignment(horizontal="center", vertical="center")
        c_uc.alignment = Alignment(horizontal="right", vertical="center")
        c_amt.alignment = Alignment(horizontal="right", vertical="center")

        c_qty.number_format = int_fmt
        c_uc.number_format = acct_fmt
        c_amt.number_format = acct_fmt

        top_b = medium if idx == 0 else thin
        for c in range(2, 12):
            cell = ws.cell(r, c)
            cell.font = Font(name="Calibri", size=11.0, bold=False)
            cell.border = Border(
                top=top_b,
                bottom=thin,
                left=medium if c in (2, 10) else (thin if c in (3, 5, 6, 8, 9, 11) else None),
                right=medium if c in (9, 11) else (thin if c in (2, 4, 5, 7, 8, 10) else None),
            )

    last_data_row = start_data_row + len(effective_rows) - 1

    # --- MAIN TABLE TOTAL ROW ---
    main_tot_row = last_data_row + 1
    ws.row_dimensions[main_tot_row].height = 16.5
    ws.merge_cells(start_row=main_tot_row, start_column=2, end_row=main_tot_row, end_column=8)

    c_main_qty_tot = ws.cell(main_tot_row, 9, f"=SUM(I{start_data_row}:I{last_data_row})")
    c_main_qty_tot.font = Font(name="Calibri", size=11.0, bold=True)
    c_main_qty_tot.alignment = Alignment(horizontal="center", vertical="center")
    c_main_qty_tot.number_format = int_fmt

    ws.cell(main_tot_row, 10, None).number_format = acct_fmt

    c_main_amt_tot = ws.cell(main_tot_row, 11, f"=SUM(K{start_data_row}:K{last_data_row})")
    c_main_amt_tot.font = Font(name="Calibri", size=11.0, bold=True)
    c_main_amt_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_main_amt_tot.number_format = acct_fmt

    for c in range(2, 12):
        ws.cell(main_tot_row, c).border = Border(
            top=medium,
            bottom=medium,
            left=medium if c in (2, 10) else (thin if c in (9, 11) else None),
            right=medium if c in (9, 11) else (thin if c in (8, 10) else None),
        )

    # --- RECAPITULATION SECTION ---
    recap_hdr1 = main_tot_row + 1
    recap_hdr2 = main_tot_row + 2
    ws.row_dimensions[recap_hdr1].height = 16.0
    ws.row_dimensions[recap_hdr2].height = 16.0

    ws.merge_cells(start_row=recap_hdr1, start_column=3, end_row=recap_hdr1, end_column=5)
    ws.cell(recap_hdr1, 3, "Recapitulation:")
    ws.cell(recap_hdr1, 3).font = Font(name="Calibri", size=11.0, bold=True)
    ws.cell(recap_hdr1, 3).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=recap_hdr1, start_column=6, end_row=recap_hdr2, end_column=7)
    ws.cell(recap_hdr1, 6, "Item")
    ws.cell(recap_hdr1, 6).font = Font(name="Calibri", size=10.0, bold=True)
    ws.cell(recap_hdr1, 6).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=recap_hdr1, start_column=8, end_row=recap_hdr1, end_column=11)
    ws.cell(recap_hdr1, 8, "Recapitulation:")
    ws.cell(recap_hdr1, 8).font = Font(name="Calibri", size=11.0, bold=True)
    ws.cell(recap_hdr1, 8).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=recap_hdr2, start_column=3, end_row=recap_hdr2, end_column=4)
    ws.cell(recap_hdr2, 3, "Stock No.")
    ws.cell(recap_hdr2, 5, "Quantity")
    ws.cell(recap_hdr2, 8, "Unit")
    ws.cell(recap_hdr2, 9, "Unit Cost")
    ws.cell(recap_hdr2, 10, "Total Cost")
    ws.cell(recap_hdr2, 11, "UACS Object Code")

    for c in (3, 5, 8, 9, 10, 11):
        cell = ws.cell(recap_hdr2, c)
        cell.font = Font(name="Calibri", size=9.5, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Borders for Recapitulation Headers
    for c in range(2, 12):
        ws.cell(recap_hdr1, c).border = Border(
            top=medium,
            bottom=thin if c not in (2, 6, 7) else None,
            left=medium if c == 2 else (thin if c in (3, 6, 8) else None),
            right=medium if c == 11 else (thin if c in (2, 5, 7) else None),
        )
        ws.cell(recap_hdr2, c).border = Border(
            top=thin if c not in (2, 6, 7) else None,
            bottom=thin if c != 2 else None,
            left=medium if c in (2, 10) else (thin if c in (3, 5, 6, 8, 9, 11) else None),
            right=medium if c == 11 else (thin if c in (2, 4, 5, 7, 8, 9, 10) else None),
        )

    recap_rows = _build_recapitulation(rsmi_rows)
    if not recap_rows:
        recap_rows = [("", 0, "", "")]

    recap_start_row = recap_hdr2 + 1
    for idx, (stock_no, r_qty, desc, unit) in enumerate(recap_rows):
        r = recap_start_row + idx
        desc_len = len(str(desc or ""))
        ws.row_dimensions[r].height = 26.0 if desc_len > 36 else 20.0

        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

        c_stk = ws.cell(r, 3, stock_no if stock_no else None)
        c_qty = ws.cell(r, 5, r_qty if stock_no else None)
        c_itm = ws.cell(r, 6, desc if desc else None)
        c_unt = ws.cell(r, 8, unit if unit else None)
        c_uc = ws.cell(r, 9, None)  # Blank for Accounting Division
        c_tc = ws.cell(r, 10, f"=+I{r}*E{r}")
        c_uacs = ws.cell(r, 11, None)  # Blank for Accounting Division

        c_stk.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_qty.alignment = Alignment(horizontal="center", vertical="center")
        c_itm.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_unt.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_uc.alignment = Alignment(horizontal="right", vertical="center")
        c_tc.alignment = Alignment(horizontal="right", vertical="center")
        c_uacs.alignment = Alignment(horizontal="center", vertical="center")

        c_qty.number_format = int_fmt
        c_uc.number_format = acct_fmt
        c_tc.number_format = acct_fmt

        for c in range(2, 12):
            cell = ws.cell(r, c)
            cell.font = Font(name="Calibri", size=11.0, bold=False)
            cell.border = Border(
                top=thin if c != 2 else None,
                bottom=thin if c != 2 else None,
                left=medium if c in (2, 10) else (thin if c in (3, 5, 6, 8, 9, 11) else None),
                right=medium if c == 11 else (thin if c in (2, 4, 5, 7, 8, 9, 10) else None),
            )

    recap_last_row = recap_start_row + len(recap_rows) - 1

    # --- RECAPITULATION TOTAL ROW ---
    recap_tot_row = recap_last_row + 1
    ws.row_dimensions[recap_tot_row].height = 16.5
    ws.merge_cells(start_row=recap_tot_row, start_column=2, end_row=recap_tot_row, end_column=4)
    ws.merge_cells(start_row=recap_tot_row, start_column=6, end_row=recap_tot_row, end_column=7)

    c_rec_qty_tot = ws.cell(recap_tot_row, 5, f"=SUM(E{recap_start_row}:E{recap_last_row})")
    c_rec_qty_tot.font = Font(name="Calibri", size=11.0, bold=True)
    c_rec_qty_tot.alignment = Alignment(horizontal="center", vertical="center")
    c_rec_qty_tot.number_format = int_fmt

    c_rec_var_qty = ws.cell(recap_tot_row, 8, f"=+I{main_tot_row}-E{recap_tot_row}")
    c_rec_var_qty.font = Font(name="Calibri", size=8.0, bold=False, color="FF0000")
    c_rec_var_qty.alignment = Alignment(horizontal="center", vertical="center")
    c_rec_var_qty.number_format = var_int_fmt

    c_rec_tc_tot = ws.cell(recap_tot_row, 10, f"=SUM(J{recap_start_row}:J{recap_last_row})")
    c_rec_tc_tot.font = Font(name="Calibri", size=11.0, bold=True)
    c_rec_tc_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_rec_tc_tot.number_format = acct_fmt

    c_rec_var_tc = ws.cell(recap_tot_row, 11, f"=+K{main_tot_row}-J{recap_tot_row}")
    c_rec_var_tc.font = Font(name="Calibri", size=8.0, bold=False, color="FF0000")
    c_rec_var_tc.alignment = Alignment(horizontal="right", vertical="center")
    c_rec_var_tc.number_format = acct_fmt

    for c in range(2, 12):
        ws.cell(recap_tot_row, c).border = Border(
            top=medium,
            bottom=medium,
            left=medium if c in (2, 10) else (thin if c in (5, 6, 8, 9, 11) else None),
            right=medium if c == 11 else (thin if c in (4, 5, 7, 8, 9, 10) else None),
        )

    # --- CERTIFICATION & SIGNATORY BLOCK (Rows recap_tot_row + 1 to recap_tot_row + 8) ---
    r_s1 = recap_tot_row + 1  # Row 36 in template: "Posted by:"
    r_s2 = recap_tot_row + 2  # Row 37 in template: "I hereby certify..."
    r_s3 = recap_tot_row + 3  # Row 38 in template: blank spacer
    r_s4 = recap_tot_row + 4  # Row 39 in template: Printed Name of Supply and/or Property Custodian
    r_s5 = recap_tot_row + 5  # Row 40 in template: "Signature over Printed Name of Supply and/or"
    r_s6 = recap_tot_row + 6  # Row 41 in template: "Property Custodian"
    r_f1 = recap_tot_row + 7  # Row 42 in template: "Original - Accounting Division..."
    r_f2 = recap_tot_row + 8  # Row 43 in template: "Copy 2 - Supply and Property Custodian's file"

    ws.row_dimensions[r_s1].height = 15.0
    ws.row_dimensions[r_s2].height = 15.0
    ws.row_dimensions[r_s3].height = 15.0
    ws.row_dimensions[r_s4].height = 16.5
    ws.row_dimensions[r_s5].height = 15.0
    ws.row_dimensions[r_s6].height = 15.0
    ws.row_dimensions[r_f1].height = 15.0
    ws.row_dimensions[r_f2].height = 15.0

    # Merges in Signatory Block matching RSMI FORMAT.xlsx
    ws.merge_cells(start_row=r_s1, start_column=2, end_row=r_s1, end_column=9)
    ws.merge_cells(start_row=r_s1, start_column=10, end_row=r_s1, end_column=11)

    ws.merge_cells(start_row=r_s2, start_column=2, end_row=r_s2, end_column=9)
    ws.merge_cells(start_row=r_s2, start_column=10, end_row=r_s2, end_column=11)

    ws.merge_cells(start_row=r_s3, start_column=2, end_row=r_s3, end_column=9)
    ws.merge_cells(start_row=r_s3, start_column=10, end_row=r_s3, end_column=11)

    ws.merge_cells(start_row=r_s4, start_column=3, end_row=r_s4, end_column=7)
    ws.merge_cells(start_row=r_s4, start_column=8, end_row=r_s4, end_column=9)
    ws.merge_cells(start_row=r_s4, start_column=10, end_row=r_s4, end_column=11)

    ws.merge_cells(start_row=r_s5, start_column=3, end_row=r_s5, end_column=7)
    ws.merge_cells(start_row=r_s5, start_column=8, end_row=r_s5, end_column=9)
    ws.merge_cells(start_row=r_s5, start_column=10, end_row=r_s5, end_column=11)

    ws.merge_cells(start_row=r_s6, start_column=3, end_row=r_s6, end_column=7)
    ws.merge_cells(start_row=r_s6, start_column=8, end_row=r_s6, end_column=9)
    ws.merge_cells(start_row=r_s6, start_column=10, end_row=r_s6, end_column=11)

    # Populate Signatory Texts
    ws.cell(r_s1, 10, "Posted by:")
    ws.cell(r_s1, 10).font = Font(name="Calibri", size=11.0, bold=False)
    ws.cell(r_s1, 10).alignment = Alignment(horizontal="left", vertical="center")

    ws.cell(r_s2, 2, "I hereby certify to the correctness of the above information.")
    ws.cell(r_s2, 2).font = Font(name="Calibri", size=11.0, bold=False)
    ws.cell(r_s2, 2).alignment = Alignment(horizontal="center", vertical="center")

    # Printed Name of Supply and/or Property Custodian (Encircled in Image 4)
    clean_custodian = str(custodian_name or DEFAULT_CUSTODIAN).strip()
    ws.cell(r_s4, 3, clean_custodian)
    ws.cell(r_s4, 3).font = Font(name="Calibri", size=12.0, bold=True, color="000000")
    ws.cell(r_s4, 3).alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(r_s4, 10, "_________________________        ______")
    ws.cell(r_s4, 10).font = Font(name="Calibri", size=11.0, bold=False)
    ws.cell(r_s4, 10).alignment = Alignment(horizontal="left", vertical="center")

    ws.cell(r_s5, 3, "Signature  over Printed Name of Supply and/or")
    ws.cell(r_s5, 3).font = Font(name="Calibri", size=10.0, bold=False)
    ws.cell(r_s5, 3).alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(r_s5, 10, "Signature Over Printed Name of            Date")
    ws.cell(r_s5, 10).font = Font(name="Calibri", size=10.0, bold=False)
    ws.cell(r_s5, 10).alignment = Alignment(horizontal="left", vertical="center")

    ws.cell(r_s6, 3, "Property Custodian")
    ws.cell(r_s6, 3).font = Font(name="Calibri", size=10.0, bold=False)
    ws.cell(r_s6, 3).alignment = Alignment(horizontal="center", vertical="top")

    ws.cell(r_s6, 10, "Designated Accounting Staff")
    ws.cell(r_s6, 10).font = Font(name="Calibri", size=10.0, bold=False)
    ws.cell(r_s6, 10).alignment = Alignment(horizontal="left", vertical="top")

    # Signatory Block Borders
    for r_sig in (r_s1, r_s2, r_s3, r_s4, r_s5, r_s6):
        for c in range(2, 12):
            top_side = medium if r_sig == r_s1 else (thin if (r_sig == r_s5 and 3 <= c <= 7) else None)
            bot_side = medium if r_sig == r_s6 else (thin if (r_sig == r_s4 and 3 <= c <= 7) else None)
            left_side = medium if c in (2, 10) else None
            right_side = medium if c in (9, 11) else None
            ws.cell(r_sig, c).border = Border(top=top_side, bottom=bot_side, left=left_side, right=right_side)

    # Bottom Footers (Original / Copy 2)
    ws.merge_cells(start_row=r_f1, start_column=4, end_row=r_f1, end_column=11)
    ws.cell(r_f1, 2, "Original").font = Font(name="Calibri", size=10.0, italic=True)
    ws.cell(r_f1, 2).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(r_f1, 3, "-").font = Font(name="Calibri", size=10.0)
    ws.cell(r_f1, 3).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(r_f1, 4, "Accounting Division/Unit to support the JEV drawn to record issues of stocks").font = Font(name="Calibri", size=10.0)

    ws.merge_cells(start_row=r_f2, start_column=4, end_row=r_f2, end_column=11)
    ws.cell(r_f2, 2, "Copy 2").font = Font(name="Calibri", size=10.0, italic=True)
    ws.cell(r_f2, 2).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(r_f2, 3, "-").font = Font(name="Calibri", size=10.0)
    ws.cell(r_f2, 3).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(r_f2, 4, "Supply and Property Custodian's file").font = Font(name="Calibri", size=10.0)

    # Configure Page Setup on RSMI sheet so all columns (B:K) fit cleanly on 1 page wide when printed
    ws.print_area = f"B1:K{r_f2}"
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

    # --- SHEET 2: RIS COUNT SUMMARY ---
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    data_font = Font(name="Calibri", size=11, bold=False)
    bold_font = Font(name="Calibri", size=11, bold=True)
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    ws_sum = out_wb.create_sheet(title="RIS Count Summary")
    sum_headers = ["RIS No.", "Office / Ward", "Item Count", "Total Quantity", "Source File"]
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
    ws_sum.print_area = f"A1:E{tot_row}"
    ws_sum.sheet_properties.pageSetUpPr.fitToPage = True
    ws_sum.page_setup.paperSize = ws_sum.PAPERSIZE_A4
    ws_sum.page_setup.orientation = ws_sum.ORIENTATION_PORTRAIT
    ws_sum.page_setup.fitToWidth = 1
    ws_sum.page_setup.fitToHeight = 0
    ws_sum.print_options.horizontalCentered = True
    ws_sum.page_margins.left = 0.35
    ws_sum.page_margins.right = 0.35
    ws_sum.page_margins.top = 0.5
    ws_sum.page_margins.bottom = 0.5
    ws_sum.page_margins.header = 0.25
    ws_sum.page_margins.footer = 0.25

    try:
        out_wb.save(output_xlsx)
    except PermissionError:
        pass
    if output_xls:
        try:
            out_wb.save(output_xls)
        except PermissionError:
            pass
    out_wb.close()
    return grand_qty

def generate_populated_rsmi(
    final_dir,
    master_filename=None,
    month_upper="OCTOBER",
    year_str="2026",
    custodian_name=DEFAULT_CUSTODIAN,
):
    """
    Generates a separate RSMI workbook specifically for `RIS - <MONTH>_<YEAR>_POPULATED.xlsx`,
    named `RSMI-<MONTH>_<YEAR>-Populated.xlsx` (and `.xls`) inside `final_dir`.
    """
    if not master_filename:
        for f in os.listdir(final_dir):
            if "POPULATED" in f.upper() and f.lower().endswith(".xlsx") and not f.startswith("~$") and not f.upper().startswith("RSMI"):
                master_filename = f
                break

    if not master_filename:
        return None

    master_path = os.path.join(final_dir, master_filename)
    if not os.path.exists(master_path):
        return None

    m_match = re.search(r'RIS\s*-\s*([A-Z]+)_(\d{4})_POPULATED', master_filename, re.IGNORECASE)
    if m_match:
        month_upper = m_match.group(1).upper()
        year_str = m_match.group(2)

    pop_xlsx_name = f"RSMI-{month_upper}_{year_str}-Populated.xlsx"
    pop_xls_name = f"RSMI-{month_upper}_{year_str}-Populated.xls"
    pop_xlsx_path = os.path.join(final_dir, pop_xlsx_name)
    pop_xls_path = os.path.join(final_dir, pop_xls_name)

    office, ris_no, rcc_code, rows, item_count, total_qty = _read_ris_file_items(master_path)
    summary_rows = [{
        "ris_no": ris_no,
        "office": office,
        "item_count": item_count,
        "total_qty": total_qty,
        "filename": master_filename,
    }]

    _write_rsmi_workbook(rows, summary_rows, pop_xlsx_path, pop_xls_path, custodian_name=custodian_name)

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

def generate_rsmi_from_folder(final_dir, master_filename=None, excluded_files=None, custodian_name=None):
    """
    Scans all generated Per-Ward RIS .xlsx files inside `final_dir` (excluding Master Populated RIS,
    RSMI files, and any files in `excluded_files`), and builds `RSMI.xlsx` + `RSMI.xls`
    in the official ARMMC RSMI FORMAT.
    """
    output_xlsx = os.path.join(final_dir, "RSMI.xlsx")
    output_xls = os.path.join(final_dir, "RSMI.xls")
    excluded_set = set(excluded_files or [])

    meta_path = os.path.join(final_dir, "_rsmi_meta.json")
    if custodian_name:
        try:
            import json
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump({"custodian_name": custodian_name}, f)
        except Exception:
            pass
    elif os.path.exists(meta_path):
        try:
            import json
            with open(meta_path, "r", encoding="utf-8") as f:
                custodian_name = json.load(f).get("custodian_name", DEFAULT_CUSTODIAN)
        except Exception:
            custodian_name = DEFAULT_CUSTODIAN
    else:
        custodian_name = DEFAULT_CUSTODIAN

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
        office, ris_no, rcc_code, rows, item_count, total_qty = _read_ris_file_items(fpath)
        rsmi_rows.extend(rows)
        summary_rows.append({
            "ris_no": ris_no,
            "office": office,
            "item_count": item_count,
            "total_qty": total_qty,
            "filename": fname,
        })

    grand_qty = _write_rsmi_workbook(rsmi_rows, summary_rows, output_xlsx, output_xls, custodian_name=custodian_name)

    return {
        "rsmi_path": output_xlsx,
        "rsmi_xls_path": output_xls,
        "total_ris_files": len(summary_rows),
        "total_medicine_rows": len(rsmi_rows),
        "grand_total_qty": grand_qty,
        "summary_rows": summary_rows,
    }
