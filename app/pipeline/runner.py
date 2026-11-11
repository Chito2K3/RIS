import os
import importlib
from . import ris_generator as _ris_gen_mod
from . import rsmi_generator as _rsmi_gen_mod
importlib.reload(_ris_gen_mod)
importlib.reload(_rsmi_gen_mod)
from .unit_lookup import UnitLookup
from .rcc_lookup import RccLookup
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
    rsmi_custodian_name=None,
    progress_callback=None,
    mirror_outputs=False,
):
    """
    Executes the complete 4-stage pipeline:
      1. Clean raw Consumption & Issuance files -> save into segregated `01_CLEANED_DATA/`
      2. Generate Master Hospital RIS (`RIS - <MONTH>_<YEAR>_POPULATED.xlsx`) -> save into `02_RIS_FINAL/`
      3. Generate Per-Ward RIS files (`<WARD>_RIS_<MON>.xlsx`) -> save into `02_RIS_FINAL/`
      4. Generate Consolidated `RSMI.xlsx` -> save into `02_RIS_FINAL/`
    """
    def _emit(pct, stage, msg):
        if progress_callback:
            try:
                progress_callback(pct, stage, msg)
            except Exception:
                pass

    month_upper = str(month_name).strip().upper()
    month_title = month_upper.capitalize()
    month_short = MONTH_MAP.get(month_upper, month_upper[:3])
    year_str = str(year).strip()
    month_year_label = f"{month_title} {year_str}"

    cleaned_dir = os.path.join(output_base_dir, "01_CLEANED_DATA")
    final_dir = os.path.join(output_base_dir, "02_RIS_FINAL")
    os.makedirs(cleaned_dir, exist_ok=True)
    os.makedirs(final_dir, exist_ok=True)

    # Clean previous generated files in 01_CLEANED_DATA before Stage 1
    for old_f in os.listdir(cleaned_dir):
        old_p = os.path.join(cleaned_dir, old_f)
        if os.path.isfile(old_p):
            try:
                os.remove(old_p)
            except Exception:
                pass

    # Load Unit Lookup from Masterlist & RCC Lookup from `rs code.xls`
    _emit(8, 1, "Loading Masterlist units & RCC reference codes...")
    unit_lookup = UnitLookup(masterlist_path)
    rcc_lookup = RccLookup()
    effective_custodian = str(rsmi_custodian_name or app_by_name or "SALVE CERTEZA").strip()
    locked_files = []

    # Stage 1A: Clean Consumption files & Combine into 1 unified clean file
    _emit(15, 1, "Stage 1/4: Cleaning raw Consumption report files...")
    cleaned_consump_results = []
    for c_file in consumption_files:
        if c_file and os.path.exists(c_file):
            res = clean_consumption_file(c_file, cleaned_dir)
            cleaned_consump_results.append(res)

    consolidated_consump_items = consolidate_consumption_items(cleaned_consump_results, cleaned_dir=cleaned_dir)
    combined_consump_xlsx = os.path.join(cleaned_dir, "CONSUMPTION_COMBINED_CLEANED.xlsx") if consolidated_consump_items else None
    total_consump_qty = sum(int(float(x["Quantity"])) for x in consolidated_consump_items)

    # Stage 1B: Clean Issuance files before touching 02_RIS_FINAL
    _emit(28, 1, "Stage 1/4: Cleaning raw Issuance matrices into 01_CLEANED_DATA...")
    cleaned_issuance_results = []
    for i_file in issuance_files:
        if i_file and os.path.exists(i_file):
            res = clean_issuance_file(i_file, cleaned_dir)
            cleaned_issuance_results.append(res)

    # Now clear old .xlsx/.xls files in 02_RIS_FINAL right before writing new ones
    for old_f in os.listdir(final_dir):
        old_p = os.path.join(final_dir, old_f)
        if os.path.isfile(old_p):
            try:
                os.remove(old_p)
            except PermissionError:
                if old_f not in locked_files and old_f.lower().endswith((".xlsx", ".xls")):
                    locked_files.append(old_f)
            except Exception:
                pass

    # Stage 2: Generate Master Populated RIS if consumption files were provided
    _emit(38, 2, "Stage 2/4: Generating Master Populated RIS workbook...")
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
            rcc_code=rcc_lookup.resolve_rcc("PHARMACY"),
            req_by_name=req_by_name,
            req_by_desig=req_by_desig,
            app_by_name=app_by_name,
            app_by_desig=app_by_desig,
            sig_date=sig_date,
            locked_files=locked_files,
        )

    # Stage 3: Combine Issuances & Generate Per-Ward RIS files
    _emit(46, 3, "Stage 3/4: Generating Per-Ward RIS slips...")
    generated_wards = []
    combined_issuance_info = None
    if cleaned_issuance_results:
        combined_issuance_info = combine_issuance_tables(cleaned_issuance_results, cleaned_dir)
        def _ward_prog(idx, total, ward_name):
            pct = 46 + int(32 * (idx / max(total, 1)))
            _emit(pct, 3, f"Stage 3/4: Building Ward RIS {idx}/{total} — {ward_name}")

        generated_wards = generate_all_ward_ris_files(
            df_combined_issuance=combined_issuance_info["df"],
            active_wards=combined_issuance_info["active_wards"],
            final_dir=final_dir,
            unit_lookup=unit_lookup,
            rcc_lookup=rcc_lookup,
            month_short=month_short,
            year=year_str,
            month_year_label=month_year_label,
            req_by_name=req_by_name,
            req_by_desig=req_by_desig,
            app_by_name=app_by_name,
            app_by_desig=app_by_desig,
            sig_date=sig_date,
            locked_files=locked_files,
            progress_callback=_ward_prog,
        )

    # Stage 4A: Generate separate RSMI-<MONTH>_<YEAR>-Populated.xlsx / .xls for Master Populated RIS
    _emit(82, 4, "Stage 4/4: Compiling Master Populated RSMI workbooks...")
    populated_rsmi_info = None
    if master_ris_path and os.path.exists(master_ris_path):
        populated_rsmi_info = generate_populated_rsmi(
            final_dir=final_dir,
            master_filename=master_ris_filename,
            month_upper=month_upper,
            year_str=year_str,
            custodian_name=effective_custodian,
        )
        for lf in (populated_rsmi_info or {}).get("locked_files", []):
            if lf not in locked_files:
                locked_files.append(lf)

    # Stage 4B: Generate Ward Issuance RSMI.xlsx / RSMI.xls inside 02_RIS_FINAL
    _emit(91, 4, "Stage 4/4: Compiling Ward RSMI (.xlsx / .xls) & Fund Clusters...")
    rsmi_info = generate_rsmi_from_folder(
        final_dir,
        master_filename=master_ris_filename,
        custodian_name=effective_custodian,
    )
    for lf in rsmi_info.get("locked_files", []):
        if lf not in locked_files:
            locked_files.append(lf)

    # Optional explicit mirroring across secondary PC directories (disabled by default)
    if mirror_outputs:
        try:
            import shutil
            folder_name = os.path.basename(os.path.normpath(output_base_dir))
            mirror_roots = [
                r"C:\Users\PC 127\Desktop\RIS NEW WITH APP\RIS",
                r"E:\RIS NEW WITH APP\RIS",
            ]
            norm_out = os.path.normcase(os.path.abspath(output_base_dir))
            for m_root in mirror_roots:
                if os.path.exists(m_root):
                    m_out = os.path.join(m_root, folder_name)
                    if os.path.normcase(os.path.abspath(m_out)) != norm_out:
                        for sub_name, src_sub in [("01_CLEANED_DATA", cleaned_dir), ("02_RIS_FINAL", final_dir)]:
                            dst_sub = os.path.join(m_out, sub_name)
                            os.makedirs(dst_sub, exist_ok=True)
                            for fn in os.listdir(src_sub):
                                s_f = os.path.join(src_sub, fn)
                                d_f = os.path.join(dst_sub, fn)
                                if os.path.isfile(s_f):
                                    try:
                                        shutil.copy2(s_f, d_f)
                                    except Exception:
                                        pass
        except Exception:
            pass

    unassigned_map = {}
    for entry in (populated_rsmi_info or {}).get("unassigned_prefixes", []):
        unassigned_map[entry["stock_no"]] = entry
    for entry in rsmi_info.get("unassigned_prefixes", []):
        unassigned_map[entry["stock_no"]] = entry

    audit_summary = {
        "locked_files": locked_files,
        "unit_fallbacks": list(unit_lookup.fallback_items.values()),
        "rcc_fallbacks": list(rcc_lookup.fallback_wards.values()),
        "unassigned_prefixes": list(unassigned_map.values()),
    }

    _emit(100, 4, "Completed! All RIS & RSMI files generated.")

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
        "audit": audit_summary,
    }

