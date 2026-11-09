import os
import re
import math
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.pagebreak import Break
from .rcc_lookup import RccLookup

RSMI_ITEMS_PER_PAGE = 10

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
    raw_ris_no = ws.cell(6, 4).value or ws.cell(6, 5).value
    ris_no = extract_ris_no(raw_ris_no)
    lookup = rcc_lookup or _RCC_LOOKUP
    if office and lookup:
        rcc_code = lookup.resolve_rcc(office)
    else:
        raw_rcc = ws.cell(5, 4).value or ws.cell(5, 5).value
        rcc_code = extract_rcc(raw_rcc)

    non_item_labels = {
        "REQUISITION AND ISSUE SLIP",
        "DIVISION:",
        "OFFICE:",
        "REQUISITION",
        "STOCK NO.",
        "STOCK NO",
        "TOTAL",
        "SIGNATURE",
        "PRINTED NAME:",
        "NAME:",
        "DESIGNATION:",
        "DATE:",
    }

    rows = []
    item_count = 0
    total_qty = 0

    for r in range(9, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_a_str = str(col_a).strip() if col_a is not None else ""
        if not col_a_str:
            continue
        u_a = col_a_str.upper()
        if (
            u_a in non_item_labels
            or u_a.startswith("ENTITY NAME")
            or u_a.startswith("PURPOSE:")
            or u_a.startswith("TOTAL")
        ):
            continue

        col_b = ws.cell(r, 2).value
        col_c = ws.cell(r, 3).value
        col_d = ws.cell(r, 4).value

        if col_d is None:
            continue
        try:
            qty_f = float(col_d)
            qty_num = int(qty_f) if qty_f.is_integer() else qty_f
        except Exception:
            continue

        unit_str = str(col_b).strip() if col_b is not None else ""
        desc_str = str(col_c).strip() if col_c is not None else ""

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

    # Generous column widths so RIS No., Responsibility Center Code, Stock No., and Accounting headers never clip or break mid-code
    col_widths = {
        "A": 1.5,
        "B": 18.0,   # RIS No. (fits OCT-2026-PHARM and ward RIS codes cleanly)
        "C": 13.0,   # Responsibility Center Code (left half) / Recap Stock No. (left half)
        "D": 13.0,   # Responsibility Center Code (right half) / Recap Stock No. (right half) -> C+D = 26 chars (fits 24-char RCC on 1 line)
        "E": 15.0,   # Stock No. / Recap Quantity (fits DMCSP000018 & PRODMR000002 on 1 line)
        "F": 22.0,   # Item (left half)
        "G": 24.0,   # Item (right half) -> F+G = 46 chars wide, wraps long medicine descriptions cleanly
        "H": 10.5,   # Unit
        "I": 11.0,   # Quantity Issued / Recap Unit Cost
        "J": 12.0,   # Unit Cost / Recap Total Cost
        "K": 13.5,   # Amount / Recap UACS Object Code -> J+K = 25.5 chars (fits Accounting subheader & UACS Object Code)
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    thin = Side(style="thin", color="000000")

    acct_fmt = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'
    int_fmt = '#,##0'
    var_int_fmt = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'

    def _write_rsmi_top_banner(base_r):
        """
        Writes the 7-row top banner of RSMI (Title, Kind of Supplies, Spacers, Entity Name, Fund Cluster, Spacer)
        starting at `base_r` (returns next available row `base_r + 7`).
        """
        r_title = base_r
        r_kind = base_r + 1
        r_sp1 = base_r + 2
        r_sp2 = base_r + 3
        r_ent = base_r + 4
        r_fc = base_r + 5
        r_sp3 = base_r + 6

        ws.row_dimensions[r_title].height = 20.0
        ws.merge_cells(start_row=r_title, start_column=2, end_row=r_title, end_column=11)
        ws.cell(r_title, 2, "REPORT OF SUPPLIES AND MATERIALS ISSUED")
        ws.cell(r_title, 2).font = Font(name="Calibri", size=16.0, bold=True)
        ws.cell(r_title, 2).alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[r_kind].height = 15.0
        ws.merge_cells(start_row=r_kind, start_column=5, end_row=r_kind, end_column=6)
        ws.cell(r_kind, 5, "Kind of Supplies: ")
        ws.cell(r_kind, 5).font = Font(name="Calibri", size=12.0, bold=False)
        ws.cell(r_kind, 5).alignment = Alignment(horizontal="right", vertical="center")
        ws.merge_cells(start_row=r_kind, start_column=7, end_row=r_kind, end_column=9)
        for c in range(7, 10):
            ws.cell(r_kind, c).border = Border(bottom=thin)

        ws.row_dimensions[r_sp1].height = 10.0
        ws.row_dimensions[r_sp2].height = 10.0
        ws.merge_cells(start_row=r_sp1, start_column=2, end_row=r_sp1, end_column=11)
        ws.merge_cells(start_row=r_sp2, start_column=2, end_row=r_sp2, end_column=11)

        # Entity Name & Serial No.
        ws.row_dimensions[r_ent].height = 15.5
        ws.merge_cells(start_row=r_ent, start_column=2, end_row=r_ent, end_column=3)
        ws.cell(r_ent, 2, "Entity Name:")
        ws.cell(r_ent, 2).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_ent, 2).alignment = Alignment(horizontal="left", vertical="center")

        ws.cell(r_ent, 4, "'AMANG' RODRIGUEZ MEMORIAL MEDICAL CENTER")
        ws.cell(r_ent, 4).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_ent, 4).alignment = Alignment(horizontal="left", vertical="center")

        ws.cell(r_ent, 10, "Serial No.:")
        ws.cell(r_ent, 10).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_ent, 10).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(r_ent, 11).border = Border(bottom=thin)

        # Fund Cluster & Date
        ws.row_dimensions[r_fc].height = 16.0
        ws.merge_cells(start_row=r_fc, start_column=2, end_row=r_fc, end_column=3)
        ws.cell(r_fc, 2, "Fund Cluster:")
        ws.cell(r_fc, 2).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_fc, 2).alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(start_row=r_fc, start_column=4, end_row=r_fc, end_column=9)
        ws.cell(r_fc, 4, "Internally Generated Fund")
        ws.cell(r_fc, 4).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_fc, 4).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(4, 10):
            ws.cell(r_fc, c).border = Border(top=thin, bottom=thin)

        ws.cell(r_fc, 10, "           Date:")
        ws.cell(r_fc, 10).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_fc, 10).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(r_fc, 11).border = Border(top=thin, bottom=thin)

        # Spacer
        ws.row_dimensions[r_sp3].height = 10.0
        ws.merge_cells(start_row=r_sp3, start_column=2, end_row=r_sp3, end_column=11)
        ws.cell(r_sp3, 2, " ")
        return r_sp3 + 1

    def _write_main_table_headers(base_r):
        """
        Writes the 3-row Main Table column header block starting at `base_r`
        (returns next available row `base_r + 3`).
        """
        r_sub = base_r
        r_h1 = base_r + 1
        r_h2 = base_r + 2

        ws.row_dimensions[r_sub].height = 18.0
        ws.merge_cells(start_row=r_sub, start_column=2, end_row=r_sub, end_column=9)
        ws.cell(r_sub, 2, "To be filled up by the Supply and/or Property Division/Unit")
        ws.cell(r_sub, 2).font = Font(name="Calibri", size=10.0, italic=True)
        ws.cell(r_sub, 2).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(2, 10):
            ws.cell(r_sub, c).border = Border(
                top=thin,
                bottom=thin,
                left=thin if c == 2 else None,
                right=thin if c == 9 else None,
            )

        ws.merge_cells(start_row=r_sub, start_column=10, end_row=r_sub, end_column=11)
        ws.cell(r_sub, 10, "To be filled up by the Accounting Division/Unit")
        ws.cell(r_sub, 10).font = Font(name="Calibri", size=9.0, italic=True)
        ws.cell(r_sub, 10).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(r_sub, 10).border = Border(top=thin, bottom=thin, left=thin)
        ws.cell(r_sub, 11).border = Border(top=thin, bottom=thin, right=thin)

        ws.row_dimensions[r_h1].height = 18.0
        ws.row_dimensions[r_h2].height = 18.0

        ws.merge_cells(start_row=r_h1, start_column=2, end_row=r_h2, end_column=2)
        ws.cell(r_h1, 2, "RIS No.")

        ws.merge_cells(start_row=r_h1, start_column=3, end_row=r_h1, end_column=4)
        ws.cell(r_h1, 3, "Responsibility")
        ws.merge_cells(start_row=r_h2, start_column=3, end_row=r_h2, end_column=4)
        ws.cell(r_h2, 3, "Center Code")

        ws.cell(r_h1, 5, "Stock ")
        ws.cell(r_h2, 5, "No.")

        ws.merge_cells(start_row=r_h1, start_column=6, end_row=r_h2, end_column=7)
        ws.cell(r_h1, 6, "Item")

        ws.merge_cells(start_row=r_h1, start_column=8, end_row=r_h2, end_column=8)
        ws.cell(r_h1, 8, "Unit")

        ws.cell(r_h1, 9, "Quantity")
        ws.cell(r_h2, 9, "Issued")

        ws.merge_cells(start_row=r_h1, start_column=10, end_row=r_h2, end_column=10)
        ws.cell(r_h1, 10, "Unit Cost")

        ws.merge_cells(start_row=r_h1, start_column=11, end_row=r_h2, end_column=11)
        ws.cell(r_h1, 11, "Amount")

        for r_hdr in (r_h1, r_h2):
            for c in range(2, 12):
                cell = ws.cell(r_hdr, c)
                cell.font = Font(name="Calibri", size=11.0, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(
                    top=thin if r_hdr == r_h1 else None,
                    bottom=thin if r_hdr == r_h2 else None,
                    left=thin if c in (2, 3, 5, 6, 8, 9, 10, 11) else None,
                    right=thin if c in (2, 4, 5, 7, 8, 9, 10, 11) else None,
                )
        return r_h2 + 1

    def _write_recap_headers(base_r):
        """
        Writes the 2-row Recapitulation column header block starting at `base_r`
        (returns next available row `base_r + 2`).
        """
        recap_hdr1 = base_r
        recap_hdr2 = base_r + 1
        ws.row_dimensions[recap_hdr1].height = 18.0
        ws.row_dimensions[recap_hdr2].height = 20.0

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

        for c in range(2, 12):
            ws.cell(recap_hdr1, c).border = Border(
                top=thin,
                bottom=thin if c not in (2, 6, 7) else None,
                left=thin if c in (2, 3, 6, 8) else None,
                right=thin if c in (2, 5, 7, 11) else None,
            )
            ws.cell(recap_hdr2, c).border = Border(
                top=thin if c not in (2, 6, 7) else None,
                bottom=thin if c != 2 else None,
                left=thin if c in (2, 3, 5, 6, 8, 9, 10, 11) else None,
                right=thin if c in (2, 4, 5, 7, 8, 9, 10, 11) else None,
            )
        return recap_hdr2 + 1

    def _write_signatory_block(base_r):
        """
        Writes the 8-row Certification & Signatory Block + Original/Copy 2 footers
        starting at `base_r` (returns last row `r_f2 = base_r + 7`).
        """
        r_s1 = base_r      # "Posted by:"
        r_s2 = base_r + 1  # "I hereby certify..."
        r_s3 = base_r + 2  # blank spacer
        r_s4 = base_r + 3  # Printed Name of Supply and/or Property Custodian
        r_s5 = base_r + 4  # "Signature over Printed Name of Supply and/or"
        r_s6 = base_r + 5  # "Property Custodian"
        r_f1 = base_r + 6  # "Original - Accounting Division..."
        r_f2 = base_r + 7  # "Copy 2 - Supply and Property Custodian's file"

        ws.row_dimensions[r_s1].height = 14.5
        ws.row_dimensions[r_s2].height = 14.5
        ws.row_dimensions[r_s3].height = 14.0
        ws.row_dimensions[r_s4].height = 16.0
        ws.row_dimensions[r_s5].height = 14.5
        ws.row_dimensions[r_s6].height = 14.5
        ws.row_dimensions[r_f1].height = 14.5
        ws.row_dimensions[r_f2].height = 14.5

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

        ws.cell(r_s1, 10, "Posted by:")
        ws.cell(r_s1, 10).font = Font(name="Calibri", size=11.0, bold=False)
        ws.cell(r_s1, 10).alignment = Alignment(horizontal="left", vertical="center")

        ws.cell(r_s2, 2, "I hereby certify to the correctness of the above information.")
        ws.cell(r_s2, 2).font = Font(name="Calibri", size=11.0, bold=False)
        ws.cell(r_s2, 2).alignment = Alignment(horizontal="center", vertical="center")

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

        for r_sig in (r_s1, r_s2, r_s3, r_s4, r_s5, r_s6):
            for c in range(2, 12):
                top_side = thin if r_sig == r_s1 else (thin if (r_sig == r_s5 and 3 <= c <= 7) else None)
                bot_side = thin if r_sig == r_s6 else (thin if (r_sig == r_s4 and 3 <= c <= 7) else None)
                left_side = thin if c in (2, 10) else None
                right_side = thin if c in (9, 11) else None
                ws.cell(r_sig, c).border = Border(top=top_side, bottom=bot_side, left=left_side, right=right_side)

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

        return r_f2

    # --- COMBINED SINGLE-PAGE RSMI LOOP (Part 1 Main Table + Part 2 Recapitulation on Every Page) ---
    effective_rows = rsmi_rows if rsmi_rows else [("", DEFAULT_RCC, "", "", "", 0)]
    main_chunks = [
        effective_rows[i:i + RSMI_ITEMS_PER_PAGE]
        for i in range(0, len(effective_rows), RSMI_ITEMS_PER_PAGE)
    ]
    main_data_ranges = []
    recap_data_ranges = []

    ws.row_dimensions[1].height = 8.0
    cur_r = 2

    for p_idx, page_rows in enumerate(main_chunks):
        is_last_page = (p_idx == len(main_chunks) - 1)

        # 1. Top Banner + 1st Part (Main Issuance Table) Headers
        cur_r = _write_rsmi_top_banner(cur_r)
        data_start_r = _write_main_table_headers(cur_r)

        # 2. 1st Part Data Rows (always standardized to RSMI_ITEMS_PER_PAGE = 10 slots per page)
        for slot_idx in range(RSMI_ITEMS_PER_PAGE):
            r = data_start_r + slot_idx
            if slot_idx < len(page_rows):
                ris_no, rcc_code, stock_no, desc, unit, qty = page_rows[slot_idx]
            else:
                ris_no, rcc_code, stock_no, desc, unit, qty = ("", "", "", "", "", None)

            desc_str = str(desc or "")
            est_lines = max(1, math.ceil(len(desc_str) / 42.0)) if desc_str else 1
            if est_lines == 1:
                ws.row_dimensions[r].height = 22.0
            elif est_lines == 2:
                ws.row_dimensions[r].height = 29.0
            else:
                ws.row_dimensions[r].height = 38.0

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

            for c in range(2, 12):
                cell = ws.cell(r, c)
                cell.font = Font(name="Calibri", size=10.0, bold=False)
                cell.border = Border(
                    top=thin,
                    bottom=thin,
                    left=thin if c in (2, 3, 5, 6, 8, 9, 10, 11) else None,
                    right=thin if c in (2, 4, 5, 7, 8, 9, 10, 11) else None,
                )

        last_data_r = data_start_r + RSMI_ITEMS_PER_PAGE - 1
        main_data_ranges.append((data_start_r, last_data_r))

        # 3. 1st Part Total / Continuation Row
        main_tot_row = last_data_r + 1
        ws.row_dimensions[main_tot_row].height = 20.0

        if is_last_page:
            ws.merge_cells(start_row=main_tot_row, start_column=2, end_row=main_tot_row, end_column=8)
            c_lbl = ws.cell(main_tot_row, 2, "TOTAL")
            c_lbl.font = Font(name="Calibri", size=11.0, bold=True)
            c_lbl.alignment = Alignment(horizontal="center", vertical="center")

            c_main_qty_tot = ws.cell(main_tot_row, 9, _build_chunked_sum_formula("I", main_data_ranges))
            c_main_qty_tot.font = Font(name="Calibri", size=11.0, bold=True)
            c_main_qty_tot.alignment = Alignment(horizontal="center", vertical="center")
            c_main_qty_tot.number_format = int_fmt

            ws.cell(main_tot_row, 10, None).number_format = acct_fmt

            c_main_amt_tot = ws.cell(main_tot_row, 11, _build_chunked_sum_formula("K", main_data_ranges))
            c_main_amt_tot.font = Font(name="Calibri", size=11.0, bold=True)
            c_main_amt_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_main_amt_tot.number_format = acct_fmt
        else:
            ws.merge_cells(start_row=main_tot_row, start_column=2, end_row=main_tot_row, end_column=5)
            c_lbl = ws.cell(main_tot_row, 2, "TOTAL")
            c_lbl.font = Font(name="Calibri", size=11.0, bold=True)
            c_lbl.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells(start_row=main_tot_row, start_column=6, end_row=main_tot_row, end_column=8)
            c_cont = ws.cell(main_tot_row, 6, "Continue to the next page...")
            c_cont.font = Font(name="Calibri", size=10.5, bold=True, italic=True)
            c_cont.alignment = Alignment(horizontal="center", vertical="center")

            ws.cell(main_tot_row, 9, None)
            ws.cell(main_tot_row, 10, None)
            ws.cell(main_tot_row, 11, None)

        for c in range(2, 12):
            ws.cell(main_tot_row, c).border = Border(
                top=thin,
                bottom=thin,
                left=thin if c in (2, 6, 9, 10, 11) else None,
                right=thin if c in (5, 8, 9, 10, 11) else None,
            )

        # 4. 2nd Part (Recapitulation Section) on the SAME Page — reflects the exact items from the 1st Part of this page
        recap_start_row = _write_recap_headers(main_tot_row + 1)
        page_recap_rows = _build_recapitulation(page_rows)
        if not page_recap_rows:
            page_recap_rows = [("", 0, "", "")]

        for slot_idx in range(RSMI_ITEMS_PER_PAGE):
            r = recap_start_row + slot_idx
            if slot_idx < len(page_recap_rows):
                stock_no, r_qty, desc, unit = page_recap_rows[slot_idx]
            else:
                stock_no, r_qty, desc, unit = ("", None, "", "")

            desc_str = str(desc or "")
            est_lines = max(1, math.ceil(len(desc_str) / 42.0)) if desc_str else 1
            if est_lines == 1:
                ws.row_dimensions[r].height = 22.0
            elif est_lines == 2:
                ws.row_dimensions[r].height = 29.0
            else:
                ws.row_dimensions[r].height = 38.0

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
                cell.font = Font(name="Calibri", size=10.0, bold=False)
                cell.border = Border(
                    top=thin if c != 2 else None,
                    bottom=thin if c != 2 else None,
                    left=thin if c in (2, 3, 5, 6, 8, 9, 10, 11) else None,
                    right=thin if c in (2, 4, 5, 7, 8, 9, 10, 11) else None,
                )

        recap_last_row = recap_start_row + RSMI_ITEMS_PER_PAGE - 1
        recap_data_ranges.append((recap_start_row, recap_last_row))

        # 5. 2nd Part (Recapitulation) Total / Continuation Row on this Page
        recap_tot_row = recap_last_row + 1
        ws.row_dimensions[recap_tot_row].height = 20.0
        ws.merge_cells(start_row=recap_tot_row, start_column=2, end_row=recap_tot_row, end_column=4)
        ws.merge_cells(start_row=recap_tot_row, start_column=6, end_row=recap_tot_row, end_column=7)

        c_rec_lbl = ws.cell(recap_tot_row, 2, "TOTAL")
        c_rec_lbl.font = Font(name="Calibri", size=11.0, bold=True)
        c_rec_lbl.alignment = Alignment(horizontal="center", vertical="center")

        if is_last_page:
            c_rec_qty_tot = ws.cell(recap_tot_row, 5, _build_chunked_sum_formula("E", recap_data_ranges))
            c_rec_qty_tot.font = Font(name="Calibri", size=11.0, bold=True)
            c_rec_qty_tot.alignment = Alignment(horizontal="center", vertical="center")
            c_rec_qty_tot.number_format = int_fmt

            c_rec_var_qty = ws.cell(recap_tot_row, 8, f"=+I{main_tot_row}-E{recap_tot_row}")
            c_rec_var_qty.font = Font(name="Calibri", size=8.0, bold=False, color="FF0000")
            c_rec_var_qty.alignment = Alignment(horizontal="center", vertical="center")
            c_rec_var_qty.number_format = var_int_fmt

            c_rec_tc_tot = ws.cell(recap_tot_row, 10, _build_chunked_sum_formula("J", recap_data_ranges))
            c_rec_tc_tot.font = Font(name="Calibri", size=11.0, bold=True)
            c_rec_tc_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_rec_tc_tot.number_format = acct_fmt

            c_rec_var_tc = ws.cell(recap_tot_row, 11, f"=+K{main_tot_row}-J{recap_tot_row}")
            c_rec_var_tc.font = Font(name="Calibri", size=8.0, bold=False, color="FF0000")
            c_rec_var_tc.alignment = Alignment(horizontal="right", vertical="center")
            c_rec_var_tc.number_format = acct_fmt
        else:
            ws.cell(recap_tot_row, 5, None)
            c_rec_cont = ws.cell(recap_tot_row, 6, "Continue to the next page...")
            c_rec_cont.font = Font(name="Calibri", size=10.5, bold=True, italic=True)
            c_rec_cont.alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(recap_tot_row, 8, None)
            ws.cell(recap_tot_row, 9, None)
            ws.cell(recap_tot_row, 10, None)
            ws.cell(recap_tot_row, 11, None)

        for c in range(2, 12):
            ws.cell(recap_tot_row, c).border = Border(
                top=thin,
                bottom=thin,
                left=thin if c in (2, 5, 6, 8, 9, 10, 11) else None,
                right=thin if c in (4, 5, 7, 8, 9, 10, 11) else None,
            )

        # 6. Certification & Signatory Block at the bottom of this page
        r_f2 = _write_signatory_block(recap_tot_row + 1)
        if not is_last_page:
            ws.row_breaks.append(Break(id=r_f2))
        cur_r = r_f2 + 1

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
