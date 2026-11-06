import os
import re
import datetime
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import pandas as pd

ITEM_CODE_RE = re.compile(r'^(AFS|DMC|DMCSP|DMD|DMDON|DMEP|DMNP|DMR|PRODMR)\d+', re.IGNORECASE)

def is_item_code(val):
    if val is None:
        return False
    s = str(val).strip()
    return bool(ITEM_CODE_RE.match(s))

def format_excel_date(val, datemode=0):
    if isinstance(val, (int, float)) and val > 30000:
        try:
            dt = xlrd.xldate_as_datetime(val, datemode)
            return dt.strftime("%m/%d/%Y %I:%M %p")
        except Exception:
            return str(val)
    return str(val or "").strip()

def save_styled_dataframe(df, xlsx_path, csv_path=None, sheet_name="Cleaned_Data"):
    if csv_path:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    header_font = Font(name="Calibri", size=11, bold=True)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    data_font = Font(name="Calibri", size=11, bold=False)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    cols = list(df.columns)
    ws.row_dimensions[1].height = 22
    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(1, c_idx, str(col_name))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            if pd.isna(val):
                val = None
            cell = ws.cell(r_idx, c_idx, val)
            cell.font = data_font
            cell.border = border
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0.##"

    # Auto-adjust column widths cleanly
    for c_idx, col_name in enumerate(cols, start=1):
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        max_len = max(len(str(col_name)), 12)
        if "Description" in str(col_name) or "Name" in str(col_name):
            max_len = 48
        elif "Code" in str(col_name):
            max_len = 16
        ws.column_dimensions[col_letter].width = min(max_len + 4, 55)

    ws.freeze_panes = "A2"
    wb.save(xlsx_path)
    wb.close()

def clean_consumption_file(file_path, cleaned_dir):
    """
    Detects whether a raw consumption .xls/.xlsx/.csv file is:
      A) Detailed Ancillary Patient Consumption Report (e.g. CONSUMP_ANCILLIARY_OCT_2026.xls)
      B) Consumption Summary Report (e.g. CONSUMP_SUMMARY_OCT_2026.xls or CONS_CONSUMP_OCT_2026.xls)
      C) Already-cleaned consumption table
    Unmerges & strips all empty cells/rows/columns, saves the cleaned product in `cleaned_dir`,
    and returns a list of summary dicts:
      [{'item_code': ..., 'description': ..., 'rendered_qty': ..., 'returned_qty': ..., 'total_qty': ...}]
    """
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1].lower()

    # Read all sheets into 2D lists of raw values (ignoring merges & empty spacer cells)
    sheets_rows = []
    datemode = 0

    if ext == ".xls":
        wb = xlrd.open_workbook(file_path)
        datemode = wb.datemode
        for s in wb.sheets():
            s_rows = []
            for r in range(s.nrows):
                non_empty = [s.cell_value(r, c) for c in range(s.ncols) if str(s.cell_value(r, c)).strip() != ""]
                if non_empty:
                    s_rows.append(non_empty)
            sheets_rows.append(s_rows)
    elif ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            # Unmerge all merged cells in memory
            for m in list(ws.merged_cells.ranges):
                ws.unmerge_cells(range_string=str(m))
            s_rows = []
            for row in ws.iter_rows(values_only=True):
                non_empty = [v for v in row if v is not None and str(v).strip() != ""]
                if non_empty:
                    s_rows.append(list(non_empty))
            sheets_rows.append(s_rows)
        wb.close()
    elif ext == ".csv":
        df_csv = pd.read_csv(file_path)
        s_rows = [list(df_csv.columns)] + df_csv.values.tolist()
        sheets_rows.append(s_rows)

    # Detect if Detailed Ancillary Report ("Item Consumption Detailed Report" or "Item Code:" markers)
    is_detailed = False
    for s_rows in sheets_rows:
        for r in s_rows[:25]:
            joined = " ".join(str(x) for x in r)
            if "Detailed Report" in joined or "Item Code:" in joined or "SlipNo" in joined:
                is_detailed = True
                break

    if is_detailed:
        return _clean_detailed_ancillary(sheets_rows, datemode, base_name, os.path.basename(file_path), cleaned_dir)
    else:
        return _clean_summary_consumption(sheets_rows, base_name, os.path.basename(file_path), cleaned_dir)

def _clean_detailed_ancillary(sheets_rows, datemode, base_name, source_filename, cleaned_dir):
    detailed_records = []
    summary_items = []

    cur_code = ""
    cur_name = ""
    cur_dept = "PHARMACY"
    running_qty = 0.0

    for s_rows in sheets_rows:
        for r in s_rows:
            first_str = str(r[0]).strip()
            # Check for Item Code:
            if first_str == "Item Code:" and len(r) >= 2:
                cur_code = str(r[1]).strip()
                running_qty = 0.0
                continue
            if first_str == "Item Name:" and len(r) >= 2:
                cur_name = str(r[1]).strip()
                continue
            if first_str.startswith("Department") and len(r) >= 2:
                cur_dept = str(r[1]).strip()
                continue
            if "SUB TOTAL" in first_str.upper():
                sub_val = running_qty
                if len(r) >= 2:
                    try:
                        sub_val = float(r[1])
                    except Exception:
                        pass
                if cur_code:
                    summary_items.append({
                        "Item Code": cur_code,
                        "Description": cur_name,
                        "Rendered Quantity": sub_val,
                        "Returned Quantity": 0.0,
                        "Total Quantity": sub_val,
                    })
                continue
            if "GRAND TOTAL" in first_str.upper() or first_str.startswith("AMANG RODRIGUEZ") or first_str.startswith("Item Consumption") or first_str.startswith("For The Period") or first_str == "SlipNo" or first_str.startswith("Prepared by") or first_str.startswith("Nurse in Charge") or first_str.startswith("Run Date"):
                continue

            # Otherwise check if patient transaction row:
            # Usually [RenderDate(float/str), PatientName, Age, Gender, Classification, Quantity, Amount, RoomNo/RenderedBy...]
            if cur_code and len(r) >= 6:
                # Find where 'MEDICINES' or classification is
                try:
                    render_date = format_excel_date(r[0], datemode)
                    patient_name = str(r[1]).strip()
                    age = r[2]
                    gender = str(r[3]).strip()
                    classification = str(r[4]).strip()
                    qty = float(r[5])
                    amt = float(r[6]) if len(r) > 6 and isinstance(r[6], (int, float)) else 0.0
                    room_no = str(r[7]).strip() if len(r) > 8 else ""
                    rendered_by = str(r[8]).strip() if len(r) > 8 else (str(r[7]).strip() if len(r) > 7 else "")
                    running_qty += qty
                    detailed_records.append({
                        "Item Code": cur_code,
                        "Item Description": cur_name,
                        "Render Date": render_date,
                        "Patient Name": patient_name,
                        "Age": age,
                        "Gender": gender,
                        "Classification": classification,
                        "Quantity": qty,
                        "Amount": amt,
                        "Room / Bed": room_no,
                        "Rendered By": rendered_by
                    })
                except Exception:
                    pass

    # Save cleaned patient-level CSV and aggregated summary XLSX + CSV in cleaned_dir
    df_det = pd.DataFrame(detailed_records)
    det_csv_path = os.path.join(cleaned_dir, f"{base_name}_PATIENT_ROWS_CLEANED.csv")
    if not df_det.empty:
        df_det.to_csv(det_csv_path, index=False, encoding="utf-8-sig")

    df_sum = pd.DataFrame(summary_items)
    sum_xlsx_path = os.path.join(cleaned_dir, f"{base_name}_CLEANED.xlsx")
    sum_csv_path = os.path.join(cleaned_dir, f"{base_name}_CLEANED.csv")
    save_styled_dataframe(df_sum, sum_xlsx_path, sum_csv_path, sheet_name="Cleaned_Consumption")

    return {
        "type": "detailed_ancillary",
        "source_file": source_filename,
        "cleaned_xlsx": sum_xlsx_path,
        "cleaned_csv": sum_csv_path,
        "patient_rows_count": len(df_det),
        "items": summary_items,
    }

def _clean_summary_consumption(sheets_rows, base_name, source_filename, cleaned_dir):
    summary_items = []
    for s_rows in sheets_rows:
        for r in s_rows:
            if not r:
                continue
            code_cand = str(r[0]).strip()
            if is_item_code(code_cand) and len(r) >= 4:
                desc = str(r[1]).strip()
                # Check if r[2] is Generic Name (string) or Rendered Quantity (number)
                nums = []
                generic_name = "(NONE)"
                for val in r[2:]:
                    try:
                        nums.append(float(str(val).replace(",", "").strip()))
                    except Exception:
                        if generic_name == "(NONE)" and str(val).strip():
                            generic_name = str(val).strip()
                rendered_q = nums[0] if len(nums) > 0 else 0.0
                returned_q = nums[1] if len(nums) > 1 else 0.0
                total_q = nums[2] if len(nums) > 2 else (rendered_q - returned_q)
                wh_q = nums[3] if len(nums) > 3 else 0.0
                dept_q = nums[4] if len(nums) > 4 else 0.0

                summary_items.append({
                    "Item Code": code_cand,
                    "Description": desc,
                    "Generic Name": generic_name,
                    "Rendered Quantity": rendered_q,
                    "Returned Quantity": returned_q,
                    "Total Quantity": total_q,
                    "Warehouse Transfer Qty": wh_q,
                    "Departmental Issuance Qty": dept_q,
                })

    df_sum = pd.DataFrame(summary_items)
    sum_xlsx_path = os.path.join(cleaned_dir, f"{base_name}_CLEANED.xlsx")
    sum_csv_path = os.path.join(cleaned_dir, f"{base_name}_CLEANED.csv")
    save_styled_dataframe(df_sum, sum_xlsx_path, sum_csv_path, sheet_name="Cleaned_Summary")

    return {
        "type": "summary_consumption",
        "source_file": source_filename,
        "cleaned_xlsx": sum_xlsx_path,
        "cleaned_csv": sum_csv_path,
        "patient_rows_count": 0,
        "items": summary_items,
    }

def consolidate_consumption_items(cleaned_results, cleaned_dir=None):
    """
    Merges consumption items across the cleaned consumption reports (e.g.
    CONSUMP_ANCILLIARY_OCT_2026.xls and CONS_CONSUMP_OCT_2026_ORIGINAL_BACKUP.xls)
    into 1 single unified clean file (`CONSUMPTION_COMBINED_CLEANED.xlsx` and `.csv`)
    inside `cleaned_dir`, and returns the consolidated items list for Master RIS generation.
    """
    merged = {}
    ordered_codes = []

    def _sort_key(r):
        if r.get("type") == "summary_consumption":
            return (0, 0 if "CONS_CONSUMP" not in r.get("source_file", "").upper() else 1)
        return (1, 0)

    ordered_res = sorted(cleaned_results, key=_sort_key)
    for res in ordered_res:
        for item in res["items"]:
            code = str(item["Item Code"]).strip()
            desc = str(item["Description"]).strip()
            qty = float(item["Total Quantity"])
            rendered = float(item.get("Rendered Quantity", qty))
            returned = float(item.get("Returned Quantity", 0.0))
            if code not in merged:
                ordered_codes.append(code)
                merged[code] = {
                    "Item Code": code,
                    "Description": desc,
                    "Rendered Quantity": rendered,
                    "Returned Quantity": returned,
                    "Quantity": qty,
                }
            elif merged[code]["Quantity"] == 0 and qty != 0:
                merged[code]["Quantity"] = qty
                merged[code]["Rendered Quantity"] = rendered
                merged[code]["Returned Quantity"] = returned

    has_reg_summary = any(
        r.get("type") == "summary_consumption" and "CONS_CONSUMP" not in r.get("source_file", "").upper()
        for r in cleaned_results
    )
    if not has_reg_summary:
        anc_codes = []
        cons_codes = []
        for res in cleaned_results:
            is_cons = "CONS_CONSUMP" in res.get("source_file", "").upper()
            for item in res["items"]:
                c = str(item["Item Code"]).strip()
                if not is_cons and c not in anc_codes:
                    anc_codes.append(c)
                elif is_cons and c not in anc_codes and c not in cons_codes:
                    cons_codes.append(c)
        ordered_codes = sorted(anc_codes) + cons_codes

    final_items = [merged[k] for k in ordered_codes]

    if cleaned_dir and final_items:
        df_comb = pd.DataFrame([
            {
                "Item Code": x["Item Code"],
                "Description": x["Description"],
                "Rendered Quantity": x["Rendered Quantity"],
                "Returned Quantity": x["Returned Quantity"],
                "Total Quantity": x["Quantity"],
            }
            for x in final_items
        ])
        comb_xlsx = os.path.join(cleaned_dir, "CONSUMPTION_COMBINED_CLEANED.xlsx")
        comb_csv = os.path.join(cleaned_dir, "CONSUMPTION_COMBINED_CLEANED.csv")
        save_styled_dataframe(df_comb, comb_xlsx, comb_csv, sheet_name="Combined_Consumption")

    return final_items

def clean_issuance_file(file_path, cleaned_dir):
    """
    Cleans raw issuance matrix file (e.g. issuance_storage_oct_2026.xls/.csv or issuance_disp_oct_2026.xls/.csv):
    - Unmerges any merged cells if Excel format
    - Strips empty rows and non-item rows
    - Strips ward columns whose sum is 0
    - Saves cleaned .xlsx and .csv into `cleaned_dir`
    """
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".csv":
        df = pd.read_csv(file_path)
    elif ext == ".xls":
        try:
            wb = xlrd.open_workbook(file_path)
            s = wb.sheet_by_index(0)
            data = [[s.cell_value(r, c) for c in range(s.ncols)] for r in range(s.nrows)]
            # Find header row containing 'Item Code'
            hdr_idx = 0
            for idx, r in enumerate(data):
                if any("ITEM CODE" in str(v).upper() for v in r):
                    hdr_idx = idx
                    break
            df = pd.DataFrame(data[hdr_idx + 1:], columns=[str(c).strip() for c in data[hdr_idx]])
        except Exception:
            # Fallback if .xls is actually CSV/HTML disguised as .xls
            df = pd.read_csv(file_path)
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        for m in list(ws.merged_cells.ranges):
            ws.unmerge_cells(range_string=str(m))
        data = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        hdr_idx = 0
        for idx, r in enumerate(data):
            if any("ITEM CODE" in str(v or "").upper() for v in r):
                hdr_idx = idx
                break
        df = pd.DataFrame(data[hdr_idx + 1:], columns=[str(c or "").strip() for c in data[hdr_idx]])

    # Identify Item Code and Item Description columns
    cols = list(df.columns)
    code_col = cols[0]
    desc_col = cols[1]
    df = df.rename(columns={code_col: "Item Code", desc_col: "Item Description"})

    # Keep only valid item rows (exclude blank rows and TOTAL row)
    df["Item Code"] = df["Item Code"].astype(str).str.strip()
    df["Item Description"] = df["Item Description"].astype(str).str.strip()
    df = df[df["Item Code"].apply(is_item_code)].copy()

    # Ward columns: all columns after Item Code & Item Description except 'Total Issuance' / 'TOTAL'
    ward_cols = [
        c for c in df.columns[2:]
        if str(c).strip() != "" and not str(c).upper().startswith("UNNAMED") and str(c).strip().upper() not in ("TOTAL", "TOTAL ISSUANCE")
    ]

    active_ward_cols = []
    for wc in ward_cols:
        df[wc] = pd.to_numeric(df[wc], errors="coerce").fillna(0)
        if df[wc].sum() > 0:
            active_ward_cols.append(wc)

    df_clean = df[["Item Code", "Item Description"] + active_ward_cols].copy()
    df_clean["Total Issuance"] = df_clean[active_ward_cols].sum(axis=1)
    # Keep only items that have at least 0 or more (preserve item row order)
    df_clean = df_clean.reset_index(drop=True)

    clean_xlsx = os.path.join(cleaned_dir, f"{base_name}_CLEANED.xlsx")
    clean_csv = os.path.join(cleaned_dir, f"{base_name}_CLEANED.csv")
    save_styled_dataframe(df_clean, clean_xlsx, clean_csv, sheet_name="Cleaned_Issuance")

    return {
        "source_file": os.path.basename(file_path),
        "cleaned_xlsx": clean_xlsx,
        "cleaned_csv": clean_csv,
        "df": df_clean,
        "active_wards": active_ward_cols,
        "total_units": int(df_clean["Total Issuance"].sum()),
    }

def combine_issuance_tables(cleaned_issuance_list, cleaned_dir):
    """
    Combines 1 or more cleaned issuance tables (e.g. dispensing + storage)
    preserving the exact row appearance order and summing quantities for identical Item Codes
    across all active Ward columns.
    """
    all_wards_set = set()
    for res in cleaned_issuance_list:
        all_wards_set.update(res["active_wards"])
    all_wards = sorted(all_wards_set)

    # Preserve order of appearance of Item Codes across the tables
    ordered_codes = []
    code_to_desc = {}
    code_ward_qty = {}

    for res in cleaned_issuance_list:
        df = res["df"]
        wards = res["active_wards"]
        for _, row in df.iterrows():
            code = str(row["Item Code"]).strip()
            desc = str(row["Item Description"]).strip()
            if code not in code_ward_qty:
                ordered_codes.append(code)
                code_to_desc[code] = desc
                code_ward_qty[code] = {w: 0 for w in all_wards}
            for w in wards:
                val = float(row[w]) if pd.notna(row[w]) else 0.0
                code_ward_qty[code][w] += val

    combined_rows = []
    for code in ordered_codes:
        row_dict = {
            "Item Code": code,
            "Item Description": code_to_desc[code],
        }
        total_item = 0
        for w in all_wards:
            q = code_ward_qty[code][w]
            q_int = int(q) if float(q).is_integer() else q
            row_dict[w] = q_int
            total_item += q_int
        if total_item > 0:
            row_dict["TOTAL"] = total_item
            combined_rows.append(row_dict)

    df_comb = pd.DataFrame(combined_rows)
    comb_xlsx = os.path.join(cleaned_dir, "ISSUANCE_COMBINED_CLEANED.xlsx")
    comb_csv = os.path.join(cleaned_dir, "ISSUANCE_COMBINED_CLEANED.csv")
    save_styled_dataframe(df_comb, comb_xlsx, comb_csv, sheet_name="Combined_Issuance")

    return {
        "combined_xlsx": comb_xlsx,
        "combined_csv": comb_csv,
        "df": df_comb,
        "active_wards": all_wards,
        "total_units": int(df_comb["TOTAL"].sum()) if not df_comb.empty else 0,
    }