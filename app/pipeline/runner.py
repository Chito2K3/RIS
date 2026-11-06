import os
from .unit_lookup import UnitLookup
from .cleaner import (
    clean_consumption_file,
    consolidate_consumption_items,
    clean_issuance_file,
    combine_issuance_tables,
)
from .ris_generator import (
    build_ris_workbook,
    generate_all_ward_ris_files,
)
from .rsmi_generator import generate_rsmi_from_folder, generate_populated_rsmi

MONTH_MAP = {
    "JANUARY": "JAN", "FEBRUARY": "FEB", "MARCH": "MAR", "APRIL": "APR",
    "MAY": "MAY", "JUNE": "JUN", "JULY": "JUL", "AUGUST": "AUG",
    "SEPTEMBER": "SEP", "OCTOBER": "OCT", "NOVEMBER": "NOV", "DECEMBER": "DEC"
}

def run_full_pipeline(
    consumption_files,
    issuance_files,
    masterlist_path,
    output_base_dir,
    month_name="OCTOBER",
    year="2026",
    req_by_name="SABA, LEONARDO CONCHITO SEBASTIAN",
    req_by_desig="Administrative Assistant II",
    app_by_name="SALVE CERTEZA",
    app_by_desig="Pharmacist V",
    sig_date="11/09/2026",
):
    """
    Executes the complete 4-stage pipeline:
      1. Clean raw Consumption & Issuance files -> save into segregated `01_CLEANED_DATA/`
      2. Generate Master Hospital RIS (`RIS - <MONTH>_<YEAR>_POPULATED.xlsx`) -> save into `02_RIS_FINAL/`
      3. Generate Per-Ward RIS files (`<WARD>_RIS_<MON>.xlsx`) -> save into `02_RIS_FINAL/`
      4. Generate Consolidated `RSMI.xlsx` -> save into `02_RIS_FINAL/`
    """
    month_upper = str(month_name).strip().upper()
    month_title = month_upper.capitalize()
    month_short = MONTH_MAP.get(month_upper, month_upper[:3])
    year_str = str(year).strip()
    month_year_label = f"{month_title} {year_str}"

    cleaned_dir = os.path.join(output_base_dir, "01_CLEANED_DATA")
    final_dir = os.path.join(output_base_dir, "02_RIS_FINAL")
    # Clean previous generated files in 01_CLEANED_DATA and 02_RIS_FINAL so each run is fresh
    for d in (cleaned_dir, final_dir):
        if os.path.exists(d):
            for old_f in os.listdir(d):
                old_p = os.path.join(d, old_f)
                if os.path.isfile(old_p):
                    try:
                        os.remove(old_p)
                    except Exception:
                        pass
        os.makedirs(d, exist_ok=True)

    # Load Unit Lookup from Masterlist
    unit_lookup = UnitLookup(masterlist_path)

    # Stage 1A: Clean Consumption files & Combine into 1 unified clean file
    cleaned_consump_results = []
    for c_file in consumption_files:
        if c_file and os.path.exists(c_file):
            res = clean_consumption_file(c_file, cleaned_dir)
            cleaned_consump_results.append(res)

    consolidated_consump_items = consolidate_consumption_items(cleaned_consump_results, cleaned_dir=cleaned_dir)
    combined_consump_xlsx = os.path.join(cleaned_dir, "CONSUMPTION_COMBINED_CLEANED.xlsx") if consolidated_consump_items else None
    total_consump_qty = sum(int(float(x["Quantity"])) for x in consolidated_consump_items)

    # Stage 2: Generate Master Populated RIS if consumption files were provided
    master_ris_filename = f"RIS - {month_upper}_{year_str}_POPULATED.xlsx"
    master_ris_path = None
    if consolidated_consump_items:
        master_ris_path = os.path.join(final_dir, master_ris_filename)
        build_ris_workbook(
            dest_path=master_ris_path,
            office_name="PHARMACY",
            ris_no_text=f"RIS No.  {month_short}-{year_str}-PHARM",
            items=consolidated_consump_items,
            unit_lookup=unit_lookup,
            is_master=True,
            month_year_label=month_year_label,
            req_by_name=req_by_name,
            req_by_desig=req_by_desig,
            app_by_name=app_by_name,
            app_by_desig=app_by_desig,
            sig_date=sig_date,
        )

    # Stage 1B: Clean Issuance files
    cleaned_issuance_results = []
    for i_file in issuance_files:
        if i_file and os.path.exists(i_file):
            res = clean_issuance_file(i_file, cleaned_dir)
            cleaned_issuance_results.append(res)

    # Stage 3: Combine Issuances & Generate Per-Ward RIS files
    generated_wards = []
    combined_issuance_info = None
    if cleaned_issuance_results:
        combined_issuance_info = combine_issuance_tables(cleaned_issuance_results, cleaned_dir)
        generated_wards = generate_all_ward_ris_files(
            df_combined_issuance=combined_issuance_info["df"],
            active_wards=combined_issuance_info["active_wards"],
            final_dir=final_dir,
            unit_lookup=unit_lookup,
            month_short=month_short,
            year=year_str,
            month_year_label=month_year_label,
            req_by_name=req_by_name,
            req_by_desig=req_by_desig,
            app_by_name=app_by_name,
            app_by_desig=app_by_desig,
            sig_date=sig_date,
        )

    # Stage 4A: Generate separate RSMI-<MONTH>_<YEAR>-Populated.xlsx / .xls for Master Populated RIS
    populated_rsmi_info = None
    if master_ris_path and os.path.exists(master_ris_path):
        populated_rsmi_info = generate_populated_rsmi(
            final_dir=final_dir,
            master_filename=master_ris_filename,
            month_upper=month_upper,
            year_str=year_str,
        )

    # Stage 4B: Generate Ward Issuance RSMI.xlsx / RSMI.xls inside 02_RIS_FINAL
    rsmi_info = generate_rsmi_from_folder(final_dir, master_filename=master_ris_filename)

    return {
        "output_base_dir": output_base_dir,
        "cleaned_dir": cleaned_dir,
        "final_dir": final_dir,
        "month_year_label": month_year_label,
        "consumption": {
            "files_cleaned": [r["source_file"] for r in cleaned_consump_results],
            "combined_cleaned_xlsx": combined_consump_xlsx,
            "patient_rows_cleaned": sum(r.get("patient_rows_count", 0) for r in cleaned_consump_results),
            "master_ris_filename": master_ris_filename if master_ris_path else None,
            "master_ris_path": master_ris_path,
            "populated_rsmi": populated_rsmi_info,
            "item_count": len(consolidated_consump_items),
            "total_qty": total_consump_qty,
        },
        "issuance": {
            "files_cleaned": [r["source_file"] for r in cleaned_issuance_results],
            "ward_count": len(generated_wards),
            "total_ward_items": sum(w["item_count"] for w in generated_wards),
            "total_qty": sum(w["total_qty"] for w in generated_wards),
            "wards": generated_wards,
        },
        "rsmi": rsmi_info,
    }

