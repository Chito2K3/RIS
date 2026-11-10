import os
import re
import math
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.pagebreak import Break
from .rcc_lookup import RccLookup

RSMI_ITEMS_PER_PAGE = 16

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

FUND_CLUSTERS = [
    {
        "code": "RAF",
        "sheet_title": "RSMI - RAF",
        "label": "Regular Agency Fund",
        "full_label": "Regular Agency Fund (RAF)",
        "file_suffix": "RAF",
    },
    {
        "code": "IGF",
        "sheet_title": "RSMI - IGF",
        "label": "Internally Generated Fund",
        "full_label": "Internally Generated Fund (IGF)",
        "file_suffix": "IGF",
    },
    {
        "code": "DON",
        "sheet_title": "RSMI - DON",
        "label": "Donations",
        "full_label": "Donation Fund (DON)",
        "file_suffix": "DON",
    },
]

def classify_fund_cluster(stock_no):
    """
    Classifies a medicine/supply item into its designated Philippine Government Fund Cluster
    based on its official ARMMC Stock Number prefix:
      - `RAF` (Regular Agency Fund / GAA / Fund 101): Begins with `DMR`, `PROD` (e.g. `PRODMR...`), or `DMNP`
      - `IGF` (Internally Generated Fund / Income / Consignment / Fund 161): Begins with `DMEP`, `DMCSP`, `DMC`
      - `DON` (Donations / Trust / Grants / Fund 184): Begins with `DMDON`, `DMD`
    Non-standard / accountable forms (e.g. `AFS000001`) default to `IGF`.
    """
    s = re.sub(r'^[^\w]+', '', str(stock_no or "")).strip().upper()
    if s.startswith("DMR") or s.startswith("PROD") or s.startswith("DMNP"):
        return "RAF"
    elif s.startswith("DMDON") or s.startswith("DMD"):
        return "DON"
    elif s.startswith("DMEP") or s.startswith("DMCSP") or s.startswith("DMC"):
        return "IGF"
    return "IGF"

def is_rsmi_or_master_file(fname, master_filename=None):
    u = fname.upper()
    if fname in SKIP_FILES:
        return True
    if u.startswith("RSMI") or "-RSMI" in u or "_RSMI" in u:
        return True
    if master_filename and fname.upper() == master_filename.upper():
        return True
    if "POPULATED" in u:
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
        "SIGNATURE:",
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
    Aggregates rsmi_rows by Stock No. preserving first-seen order.
    Returns list of (stock_no, qty_sum, description, unit).
    """
    recap_map = {}
    for r in rsmi_rows:
        if len(r) == 6:
            ris_no, rcc, stock_no, desc, unit, qty = r
        elif len(r) == 5:
            ris_no, stock_no, unit, desc, qty = r
        else:
            continue
        if not stock_no:
            continue
        try:
            q_int = int(float(qty)) if qty is not None else 0
        except Exception:
            q_int = 0
        if stock_no not in recap_map:
            recap_map[stock_no] = {
                "stock_no": stock_no,
                "qty": q_int,
                "desc": desc,
                "unit": unit,
            }
        else:
            recap_map[stock_no]["qty"] += q_int
    recap_rows = []
    for s_no, info in recap_map.items():
        recap_rows.append((info["stock_no"], info["qty"], info["desc"], info["unit"]))
    return recap_rows

def _populate_rsmi_sheet(ws, rsmi_rows, fund_cluster_label="Internally Generated Fund (IGF)", custodian_name=DEFAULT_CUSTODIAN):
    """
    Populates an openpyxl worksheet with the official ARMMC RSMI layout:
      - 12 items per page
      - Col B..I: RIS No., RCC, Stock No., Item, Unit, Quantity Issued
      - Recapitulation Table on the SAME page aggregating that page's items
      - Signatory Block at the bottom of every page
      - Dynamic Fund Cluster label in Row 7 (e.g. Regular Agency Fund, Internally Generated Fund, Donations)
      - Printable page breaks & A4 setup
    """
    col_widths = {
        "A": 1.5,
        "B": 16.5,   # RIS No.
        "C": 11.5,   # RCC (left) / Recap Stock No. (left)
        "D": 11.5,   # RCC (right) / Recap Stock No. (right) -> C+D = 23 chars
        "E": 13.5,   # Stock No. / Recap Quantity
        "F": 20.0,   # Item (left)
        "G": 21.0,   # Item (right) -> F+G = 41 chars
        "H": 9.5,    # Unit
        "I": 10.0,   # Quantity Issued / Recap Unit Cost
        "J": 11.0,   # Unit Cost / Recap Total Cost
        "K": 12.5,   # Amount / Recap UACS Object Code
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    thin = Side(style="thin", color="000000")
    acct_fmt = '_(* #,##0.00_);_(* \\(#,##0.00\\);;@'
    int_fmt = '#,##0'
    var_int_fmt = '_(* #,##0_);_(* \\(#,##0\\);;@'

    def _write_top_banner(base_r):
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

        ws.row_dimensions[r_sp1].height = 2.0
        ws.row_dimensions[r_sp2].height = 2.0
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
        ws.cell(r_fc, 4, str(fund_cluster_label or "Internally Generated Fund (IGF)"))
        ws.cell(r_fc, 4).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_fc, 4).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(4, 10):
            ws.cell(r_fc, c).border = Border(top=thin, bottom=thin)

        ws.cell(r_fc, 10, "           Date:")
        ws.cell(r_fc, 10).font = Font(name="Calibri", size=12.0, bold=True)
        ws.cell(r_fc, 10).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(r_fc, 11).border = Border(top=thin, bottom=thin)

        # Spacer
        ws.row_dimensions[r_sp3].height = 2.0
        ws.merge_cells(start_row=r_sp3, start_column=2, end_row=r_sp3, end_column=11)
        ws.cell(r_sp3, 2, " ")
        return r_sp3 + 1

    def _write_main_headers(base_r):
        r_sub = base_r
        r_h1 = base_r + 1
        r_h2 = base_r + 2

        ws.row_dimensions[r_sub].height = 16.0
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

        ws.row_dimensions[r_h1].height = 16.0
        ws.row_dimensions[r_h2].height = 16.0

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
        recap_hdr1 = base_r
        recap_hdr2 = base_r + 1
        ws.row_dimensions[recap_hdr1].height = 16.0
        ws.row_dimensions[recap_hdr2].height = 18.0

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

    def _write_signatory_block(base_r, page_num=None, total_pages=None):
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

        ws.row_dimensions[r_s1].height = 14.0
        ws.row_dimensions[r_s2].height = 14.0
        ws.row_dimensions[r_s3].height = 18.0
        ws.row_dimensions[r_s4].height = 16.0
        ws.row_dimensions[r_s5].height = 14.0
        ws.row_dimensions[r_s6].height = 14.0
        ws.row_dimensions[r_f1].height = 14.0
        ws.row_dimensions[r_f2].height = 14.0

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

        ws.merge_cells(start_row=r_f2, start_column=4, end_row=r_f2, end_column=10)
        ws.cell(r_f2, 2, "Copy 2").font = Font(name="Calibri", size=10.0, italic=True)
        ws.cell(r_f2, 2).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r_f2, 3, "-").font = Font(name="Calibri", size=10.0)
        ws.cell(r_f2, 3).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r_f2, 4, "Supply and Property Custodian's file").font = Font(name="Calibri", size=10.0)

        if page_num is not None and total_pages is not None:
            c_pg = ws.cell(r_f2, 11, f"{page_num} of {total_pages}")
            c_pg.font = Font(name="Calibri", size=10.0, italic=True)
            c_pg.alignment = Alignment(horizontal="right", vertical="center")

        return r_f2

    # --- MAIN PAGE LOOP ---
    effective_rows = rsmi_rows if rsmi_rows else [("", DEFAULT_RCC, "", "", "", None)]
    main_chunks = [
        effective_rows[i:i + RSMI_ITEMS_PER_PAGE]
        for i in range(0, len(effective_rows), RSMI_ITEMS_PER_PAGE)
    ]
    main_data_ranges = []
    recap_data_ranges = []

    ws.row_dimensions[1].height = 8.0
    cur_r = 2
    r_f2_last = cur_r

    for p_idx, page_rows in enumerate(main_chunks):
        is_last_page = (p_idx == len(main_chunks) - 1)

        # 1. Top Banner + Main Table Headers
        cur_r = _write_top_banner(cur_r)
        data_start_r = _write_main_headers(cur_r)

        # 2. 1st Part Data Rows (always standardized to RSMI_ITEMS_PER_PAGE = 16 slots per page)
        for slot_idx in range(RSMI_ITEMS_PER_PAGE):
            r = data_start_r + slot_idx
            if slot_idx < len(page_rows):
                ris_no, rcc_code, stock_no, desc, unit, qty = page_rows[slot_idx]
            else:
                ris_no, rcc_code, stock_no, desc, unit, qty = ("", "", "", "", "", None)

            desc_str = str(desc or "")
            est_lines = max(1, math.ceil(len(desc_str) / 38.0)) if desc_str else 1
            if est_lines == 1:
                ws.row_dimensions[r].height = 18.0
            elif est_lines == 2:
                ws.row_dimensions[r].height = 23.5
            else:
                ws.row_dimensions[r].height = 29.5

            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

            c_ris = ws.cell(r, 2, ris_no if ris_no else None)
            c_rcc = ws.cell(r, 3, rcc_code if stock_no else None)
            c_stk = ws.cell(r, 5, stock_no if stock_no else None)
            c_itm = ws.cell(r, 6, desc if desc else None)
            c_unt = ws.cell(r, 8, unit if unit else None)
            c_qty = ws.cell(r, 9, qty if stock_no else None)
            c_uc = ws.cell(r, 10, None)
            c_amt = ws.cell(r, 11, None)

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

        # 3. Main Total / Continuation Row
        main_tot_row = last_data_r + 1
        ws.row_dimensions[main_tot_row].height = 18.0

        if is_last_page:
            ws.merge_cells(start_row=main_tot_row, start_column=2, end_row=main_tot_row, end_column=8)
            c_lbl = ws.cell(main_tot_row, 2, "TOTAL")
            c_lbl.font = Font(name="Calibri", size=11.0, bold=True)
            c_lbl.alignment = Alignment(horizontal="center", vertical="center")

            c_main_qty_tot = ws.cell(main_tot_row, 9, _build_chunked_sum_formula("I", main_data_ranges))
            c_main_qty_tot.font = Font(name="Calibri", size=11.0, bold=True)
            c_main_qty_tot.alignment = Alignment(horizontal="center", vertical="center")
            c_main_qty_tot.number_format = int_fmt

            ws.cell(main_tot_row, 10, None)
            c_main_amt_tot = ws.cell(main_tot_row, 11, None)
        else:
            ws.merge_cells(start_row=main_tot_row, start_column=2, end_row=main_tot_row, end_column=5)
            c_lbl = ws.cell(main_tot_row, 2, "TOTAL")
            c_lbl.font = Font(name="Calibri", size=11.0, bold=True)
            c_lbl.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells(start_row=main_tot_row, start_column=6, end_row=main_tot_row, end_column=8)
            c_cont = ws.cell(main_tot_row, 6, f"Continue to the next page... ({p_idx + 1} of {len(main_chunks)})")
            c_cont.font = Font(name="Calibri", size=10.0, bold=True, italic=True)
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

        # 4. Recapitulation Section for this Page
        recap_start_row = _write_recap_headers(main_tot_row + 1)
        page_recap_rows = _build_recapitulation(page_rows)
        if not page_recap_rows:
            page_recap_rows = [("", None, "", "")]

        for slot_idx in range(RSMI_ITEMS_PER_PAGE):
            r = recap_start_row + slot_idx
            if slot_idx < len(page_recap_rows):
                stock_no, r_qty, desc, unit = page_recap_rows[slot_idx]
            else:
                stock_no, r_qty, desc, unit = ("", None, "", "")

            desc_str = str(desc or "")
            est_lines = max(1, math.ceil(len(desc_str) / 38.0)) if desc_str else 1
            if est_lines == 1:
                ws.row_dimensions[r].height = 18.0
            elif est_lines == 2:
                ws.row_dimensions[r].height = 23.5
            else:
                ws.row_dimensions[r].height = 29.5

            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

            c_stk = ws.cell(r, 3, stock_no if stock_no else None)
            c_qty = ws.cell(r, 5, r_qty if stock_no else None)
            c_itm = ws.cell(r, 6, desc if desc else None)
            c_unt = ws.cell(r, 8, unit if unit else None)
            c_uc = ws.cell(r, 9, None)
            c_tc = ws.cell(r, 10, None)
            c_uacs = ws.cell(r, 11, None)

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

        # 5. Recapitulation Total / Continuation Row
        recap_tot_row = recap_last_row + 1
        ws.row_dimensions[recap_tot_row].height = 18.0
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

            ws.cell(recap_tot_row, 8, None)
            ws.cell(recap_tot_row, 9, None)
            ws.cell(recap_tot_row, 10, None)
            ws.cell(recap_tot_row, 11, None)
        else:
            ws.cell(recap_tot_row, 5, None)
            c_rec_cont = ws.cell(recap_tot_row, 6, f"Continue to the next page... ({p_idx + 1} of {len(main_chunks)})")
            c_rec_cont.font = Font(name="Calibri", size=10.0, bold=True, italic=True)
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
        r_f2 = _write_signatory_block(recap_tot_row + 1, page_num=p_idx + 1, total_pages=len(main_chunks))
        r_f2_last = r_f2
        if not is_last_page:
            ws.row_breaks.append(Break(id=r_f2))
        cur_r = r_f2 + 1

    # Page setup
    ws.print_area = f"B1:K{r_f2_last}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.15
    ws.page_margins.right = 0.15
    ws.page_margins.top = 0.15
    ws.page_margins.bottom = 0.15
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0
    ws.oddFooter.right.text = "&P of &N"
    ws.oddFooter.right.font = "Calibri"
    ws.oddFooter.right.size = 10

def _write_summary_sheet(out_wb, summary_rows, rsmi_rows, title="RIS Count Summary"):
    """
    Creates Sheet 'RIS Count Summary' detailing per-RIS item counts, total quantity,
    and a clear breakdown across RAF, IGF, and DON Fund Clusters.
    """
    thin = Side(style="thin", color="000000")
    header_font = Font(name="Calibri", size=11, bold=True, color="000000")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    data_font = Font(name="Calibri", size=11, bold=False)
    bold_font = Font(name="Calibri", size=11, bold=True)
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    ws_sum = out_wb.create_sheet(title=title)
    sum_headers = [
        "RIS No.", "Office / Ward",
        "RAF Qty (DMR/PROD/DMNP)", "IGF Qty (DMEP/DMC)", "DON Qty (DMD/DMDON)",
        "Item Count", "Total Quantity", "Source File"
    ]
    ws_sum.row_dimensions[1].height = 24
    for col_idx, h_text in enumerate(sum_headers, start=1):
        cell = ws_sum.cell(1, col_idx, h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_all

    grand_raf_qty = 0
    grand_igf_qty = 0
    grand_don_qty = 0
    grand_items = 0
    grand_qty = 0

    # Build index of items by ris_no for fast lookup
    items_by_ris = {}
    for r in rsmi_rows:
        r_no = r[0]
        items_by_ris.setdefault(r_no, []).append(r)

    for idx, srow in enumerate(summary_rows, start=2):
        r_no = srow.get("ris_no") or ""
        matched_items = items_by_ris.get(r_no, [])

        raf_q = sum(int(float(it[5])) for it in matched_items if classify_fund_cluster(it[2]) == "RAF")
        igf_q = sum(int(float(it[5])) for it in matched_items if classify_fund_cluster(it[2]) == "IGF")
        don_q = sum(int(float(it[5])) for it in matched_items if classify_fund_cluster(it[2]) == "DON")

        c1 = ws_sum.cell(idx, 1, r_no if r_no else None)
        c2 = ws_sum.cell(idx, 2, srow["office"])
        c3 = ws_sum.cell(idx, 3, raf_q)
        c4 = ws_sum.cell(idx, 4, igf_q)
        c5 = ws_sum.cell(idx, 5, don_q)
        c6 = ws_sum.cell(idx, 6, srow["item_count"])
        c7 = ws_sum.cell(idx, 7, srow["total_qty"])
        c8 = ws_sum.cell(idx, 8, srow["filename"])

        for c in (c1, c2, c3, c4, c5, c6, c7, c8):
            c.font = data_font
            c.border = border_all

        c1.alignment = align_center
        c2.alignment = align_left
        c3.alignment = align_right
        c3.number_format = "#,##0"
        c4.alignment = align_right
        c4.number_format = "#,##0"
        c5.alignment = align_right
        c5.number_format = "#,##0"
        c6.alignment = align_right
        c6.number_format = "#,##0"
        c7.alignment = align_right
        c7.number_format = "#,##0"
        c8.alignment = align_left

        grand_raf_qty += raf_q
        grand_igf_qty += igf_q
        grand_don_qty += don_q
        grand_items += srow["item_count"]
        grand_qty += srow["total_qty"]

    tot_row = len(summary_rows) + 2
    ws_sum.cell(tot_row, 1, "TOTAL").font = bold_font
    ws_sum.cell(tot_row, 1).alignment = align_center
    ws_sum.cell(tot_row, 2, f"{len(summary_rows)} RIS File(s)").font = bold_font
    ws_sum.cell(tot_row, 2).alignment = align_left
    ws_sum.cell(tot_row, 3, grand_raf_qty).font = bold_font
    ws_sum.cell(tot_row, 3).alignment = align_right
    ws_sum.cell(tot_row, 3).number_format = "#,##0"
    ws_sum.cell(tot_row, 4, grand_igf_qty).font = bold_font
    ws_sum.cell(tot_row, 4).alignment = align_right
    ws_sum.cell(tot_row, 4).number_format = "#,##0"
    ws_sum.cell(tot_row, 5, grand_don_qty).font = bold_font
    ws_sum.cell(tot_row, 5).alignment = align_right
    ws_sum.cell(tot_row, 5).number_format = "#,##0"
    ws_sum.cell(tot_row, 6, grand_items).font = bold_font
    ws_sum.cell(tot_row, 6).alignment = align_right
    ws_sum.cell(tot_row, 6).number_format = "#,##0"
    ws_sum.cell(tot_row, 7, grand_qty).font = bold_font
    ws_sum.cell(tot_row, 7).alignment = align_right
    ws_sum.cell(tot_row, 7).number_format = "#,##0"
    ws_sum.cell(tot_row, 8, "").font = bold_font
    for col_idx in range(1, 9):
        ws_sum.cell(tot_row, col_idx).border = border_all
        ws_sum.cell(tot_row, col_idx).fill = header_fill

    ws_sum.column_dimensions["A"].width = 24
    ws_sum.column_dimensions["B"].width = 34
    ws_sum.column_dimensions["C"].width = 16
    ws_sum.column_dimensions["D"].width = 19
    ws_sum.column_dimensions["E"].width = 19
    ws_sum.column_dimensions["F"].width = 13
    ws_sum.column_dimensions["G"].width = 15
    ws_sum.column_dimensions["H"].width = 44
    ws_sum.freeze_panes = "A2"
    ws_sum.print_area = f"A1:H{tot_row}"
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

def _write_rsmi_workbook(
    rsmi_rows,
    summary_rows,
    output_xlsx,
    output_xls=None,
    custodian_name=DEFAULT_CUSTODIAN,
    standalone_prefix="RSMI",
):
    """
    Generates the official ARMMC RSMI deliverables segregated strictly by Fund Cluster:
      1. Consolidated Master Workbook (`output_xlsx` & `output_xls`):
         - Sheet 1: `RSMI - RAF` (Row 7: Regular Agency Fund (RAF), only DMR items)
         - Sheet 2: `RSMI - IGF` (Row 7: Internally Generated Fund (IGF), only DMEP, DMC, DMCSP items)
         - Sheet 3: `RSMI - DON` (Row 7: Donation Fund (DON), only DMD, DMDON items)
         - Sheet 4: `RIS Count Summary` (Full breakdown across all 3 funds)
      2. Standalone Dedicated Workbooks for each Fund Cluster in the same directory:
         - `{prefix} - RAF.xlsx` & `.xls`
         - `{prefix} - IGF.xlsx` & `.xls`
         - `{prefix} - DON.xlsx` & `.xls`
    """
    # 1. Segregate rows by Fund Cluster
    cluster_buckets = {"RAF": [], "IGF": [], "DON": []}
    for row in rsmi_rows:
        stock_no = row[2]
        fc = classify_fund_cluster(stock_no)
        cluster_buckets[fc].append(row)

    # 2. Build Consolidated Workbook
    out_wb = openpyxl.Workbook()
    for idx, fc_meta in enumerate(FUND_CLUSTERS):
        fc_code = fc_meta["code"]
        fc_title = fc_meta["sheet_title"]
        fc_label = fc_meta["full_label"]
        rows_in_cluster = cluster_buckets[fc_code]

        if idx == 0:
            ws = out_wb.active
            ws.title = fc_title
        else:
            ws = out_wb.create_sheet(title=fc_title)

        _populate_rsmi_sheet(
            ws,
            rsmi_rows=rows_in_cluster,
            fund_cluster_label=fc_label,
            custodian_name=custodian_name,
        )

    _write_summary_sheet(out_wb, summary_rows, rsmi_rows, title="RIS Count Summary")

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

    # 3. Build Standalone Workbooks per Fund Cluster
    out_dir = os.path.dirname(os.path.abspath(output_xlsx))
    fund_files = {}

    for fc_meta in FUND_CLUSTERS:
        fc_code = fc_meta["code"]
        fc_label = fc_meta["full_label"]
        rows_in_cluster = cluster_buckets[fc_code]

        single_xlsx_name = f"{standalone_prefix} - {fc_code}.xlsx"
        single_xls_name = f"{standalone_prefix} - {fc_code}.xls"
        single_xlsx_path = os.path.join(out_dir, single_xlsx_name)
        single_xls_path = os.path.join(out_dir, single_xls_name)

        s_wb = openpyxl.Workbook()
        ws_single = s_wb.active
        ws_single.title = "RSMI"

        _populate_rsmi_sheet(
            ws_single,
            rsmi_rows=rows_in_cluster,
            fund_cluster_label=fc_label,
            custodian_name=custodian_name,
        )
        _write_summary_sheet(s_wb, summary_rows, rows_in_cluster, title="RIS Count Summary")

        try:
            s_wb.save(single_xlsx_path)
        except PermissionError:
            pass
        if single_xls_path:
            try:
                s_wb.save(single_xls_path)
            except PermissionError:
                pass
        s_wb.close()

        c_qty = sum(int(float(r[5])) for r in rows_in_cluster)
        fund_files[fc_code] = {
            "xlsx": single_xlsx_name,
            "xls": single_xls_name,
            "xlsx_path": single_xlsx_path,
            "xls_path": single_xls_path,
            "item_count": len(rows_in_cluster),
            "total_qty": c_qty,
            "label": fc_label,
        }

    grand_qty = sum(int(float(r[5])) for r in rsmi_rows)
    return {
        "grand_qty": grand_qty,
        "fund_files": fund_files,
        "breakdown": {
            "RAF": {"items": len(cluster_buckets["RAF"]), "qty": fund_files["RAF"]["total_qty"]},
            "IGF": {"items": len(cluster_buckets["IGF"]), "qty": fund_files["IGF"]["total_qty"]},
            "DON": {"items": len(cluster_buckets["DON"]), "qty": fund_files["DON"]["total_qty"]},
        },
    }

def generate_populated_rsmi(
    final_dir,
    master_filename=None,
    month_upper="OCTOBER",
    year_str="2026",
    custodian_name=DEFAULT_CUSTODIAN,
):
    """
    Generates segregated RSMI workbooks specifically for `RIS - <MONTH>_<YEAR>_POPULATED.xlsx`,
    including:
      - `RSMI-<MONTH>_<YEAR>-Populated.xlsx` / `.xls` (Multi-sheet: RAF, IGF, DON, Summary)
      - `RSMI-<MONTH>_<YEAR>-Populated - RAF.xlsx` / `.xls` (DMR only)
      - `RSMI-<MONTH>_<YEAR>-Populated - IGF.xlsx` / `.xls` (DMEP, DMC, DMCSP only)
      - `RSMI-<MONTH>_<YEAR>-Populated - DON.xlsx` / `.xls` (DMD, DMDON only)
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

    base_prefix = f"RSMI-{month_upper}_{year_str}-Populated"
    pop_xlsx_name = f"{base_prefix}.xlsx"
    pop_xls_name = f"{base_prefix}.xls"
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

    write_res = _write_rsmi_workbook(
        rows,
        summary_rows,
        pop_xlsx_path,
        pop_xls_path,
        custodian_name=custodian_name,
        standalone_prefix=base_prefix,
    )

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
        "fund_files": write_res["fund_files"],
        "breakdown": write_res["breakdown"],
    }

def generate_rsmi_from_folder(final_dir, master_filename=None, excluded_files=None, custodian_name=None):
    """
    Scans all generated Per-Ward RIS .xlsx files inside `final_dir` (excluding Master Populated RIS,
    RSMI files, and any files in `excluded_files`), and builds:
      - `RSMI.xlsx` & `RSMI.xls` (Multi-sheet: RAF, IGF, DON, Summary)
      - `RSMI - RAF.xlsx` & `RSMI - RAF.xls`
      - `RSMI - IGF.xlsx` & `RSMI - IGF.xls`
      - `RSMI - DON.xlsx` & `RSMI - DON.xls`
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

    write_res = _write_rsmi_workbook(
        rsmi_rows,
        summary_rows,
        output_xlsx,
        output_xls,
        custodian_name=custodian_name,
        standalone_prefix="RSMI",
    )

    return {
        "rsmi_path": output_xlsx,
        "rsmi_xls_path": output_xls,
        "total_ris_files": len(summary_rows),
        "total_medicine_rows": len(rsmi_rows),
        "grand_total_qty": write_res["grand_qty"],
        "fund_files": write_res["fund_files"],
        "breakdown": write_res["breakdown"],
        "summary_rows": summary_rows,
    }
