import os
import re
import xlrd
import openpyxl

DEFAULT_RCC_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    "rs code.xls",
)
DEFAULT_PHARMACY_RCC = "13-001-14-00012-03-02-33"


class RccLookup:
    """
    Loads ARMMC Responsibility Center Codes from `rs code.xls`:
      - Column B (index 1): Responsibility Center Code (e.g. `13-001-14-00012-03-02-10`)
      - Column E (index 4): Short Office/Ward Code (e.g. `ERMED`, `PEDWARD`, `PHARMA`)
      - Column F (index 5): Full Department/Ward Name (e.g. `EMERGENCY MEDICINE`, `PEDIA WARD`, `PHARMACY`)
    Resolves any HIS Ward / Office name to its designated Responsibility Center Code,
    with intelligent fallback to the parent department when a ward/section name is not
    listed verbatim in `rs code.xls` (e.g., `EMERGENCY ROOM DEPARTMENT` or `ER WARD` -> `EMERGENCY MEDICINE`).
    """

    def __init__(self, rs_code_path=None):
        self.rs_code_path = rs_code_path or DEFAULT_RCC_PATH
        self.short_to_rcc = {}
        self.full_to_rcc = {}
        self.norm_to_rcc = {}
        self.entries = []
        self.fallback_wards = {}
        if self.rs_code_path and os.path.exists(self.rs_code_path):
            self.load_rs_code_file(self.rs_code_path)

    @staticmethod
    def _norm(text):
        s = str(text or "").strip().upper()
        s = re.sub(r"[^A-Z0-9]+", " ", s)
        return re.sub(r"\s+", " ", s).strip()

    def load_rs_code_file(self, path):
        ext = os.path.splitext(path)[1].lower()
        rows = []
        if ext == ".xls":
            wb = xlrd.open_workbook(path)
            s = wb.sheet_by_index(0)
            for r in range(s.nrows):
                row_vals = [s.cell_value(r, c) for c in range(s.ncols)]
                rows.append(row_vals)
        else:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            for row_vals in ws.iter_rows(values_only=True):
                rows.append(list(row_vals))
            wb.close()

        for row in rows:
            if len(row) < 6:
                continue
            rcc = str(row[1] or "").strip()
            short_code = str(row[4] or "").strip().upper()
            full_name = str(row[5] or "").strip().upper()
            alias_code = str(row[6] or "").strip().upper() if len(row) > 6 and row[6] else ""
            if not rcc or not rcc.startswith("13"):
                continue

            self.entries.append((rcc, short_code, full_name))
            if short_code:
                self.short_to_rcc[short_code] = rcc
                self.norm_to_rcc[self._norm(short_code)] = rcc
            if full_name:
                self.full_to_rcc[full_name] = rcc
                self.norm_to_rcc[self._norm(full_name)] = rcc
            if alias_code:
                self.short_to_rcc[alias_code] = rcc
                self.norm_to_rcc[self._norm(alias_code)] = rcc

    def _get_by_short(self, short_code, fallback=DEFAULT_PHARMACY_RCC):
        return self.short_to_rcc.get(short_code.upper(), fallback)

    def resolve_rcc(self, office_name):
        rcc = self._resolve_rcc_internal(office_name, track_default=True)
        return rcc

    def _resolve_rcc_internal(self, office_name, track_default=False):
        if not office_name:
            return self._get_by_short("PHARMA")

        raw_u = str(office_name).strip().upper()
        norm_u = self._norm(raw_u)

        # 1. Direct exact match against Column E (short code) or Column F (full name)
        if raw_u in self.full_to_rcc:
            return self.full_to_rcc[raw_u]
        if raw_u in self.short_to_rcc:
            return self.short_to_rcc[raw_u]
        if norm_u in self.norm_to_rcc:
            return self.norm_to_rcc[norm_u]

        # 2. Explicit HIS Ward / Department synonym & parent-department fallback rules
        # Pharmacy & Consignment / Storage / Warehouse
        if "PHARMACY" in raw_u or "DRUGS AND MEDICINES" in raw_u:
            return self._get_by_short("PHARMA")

        # Emergency variants (e.g. EMERGENCY ROOM DEPARTMENT, EMERGENCY MEDICINE DEPARTMENT, ER WARD, EMERGENCY WARD, BUCAS, ANIMAL BITE, HEMB, ERID)
        if "ER OB" in raw_u or "EMERGENCY OB" in raw_u:
            return self._get_by_short("ER OBGYNE")
        if "ER PED" in raw_u or "EMERGENCY PED" in raw_u:
            return self._get_by_short("ER PED")
        if "ER OPTH" in raw_u or "ER OPHTH" in raw_u or "EMERGENCY OPTH" in raw_u or "EMERGENCY OPHTH" in raw_u:
            return self._get_by_short("ER OPTHA")
        if "ER SURG" in raw_u or "EMERGENCY SURG" in raw_u or "ERSURGE" in raw_u:
            return self._get_by_short("ERSURGE")
        if "ER FAM" in raw_u or "EMERGENCY FAM" in raw_u:
            return self._get_by_short("ER FAMED")
        if "EMERGENCY ROOM LAB" in raw_u:
            return self._get_by_short("LAB")
        if (
            "EMERGENCY" in raw_u
            or raw_u.startswith("ER ")
            or raw_u.startswith("ER-")
            or "ERID" in raw_u
            or "ANIMAL BITE" in raw_u
            or "BUCAS" in raw_u
            or "HEMB" in raw_u
        ):
            # Use EMERGENCY MEDICINE (13-001-14-00012-03-02-10)
            return self._get_by_short("ERMED")

        # Triage
        if "TRIAGE" in raw_u:
            return self._get_by_short("TRIAGE")

        # Neonatal ICU / Newborn Care
        if "NICU" in raw_u or "NEONATAL" in raw_u or "NEWBORN" in raw_u:
            return self._get_by_short("NICU")

        # Intensive Care Units & Specialized Care
        if "ICU" in raw_u or "INTENSIVE CARE" in raw_u or "INTERMEDIATE MEDICAL CARE" in raw_u or "BURN UNIT" in raw_u:
            return self._get_by_short("ICU")

        # OB-GYNE variants
        if "DELIVERY ROOM" in raw_u or "LABOR ROOM" in raw_u or "LABOR-DR" in raw_u:
            return self._get_by_short("LABOR-DR")
        if "OPD OB" in raw_u or "OB-GYNE OPD" in raw_u or "OBGYNE OPD" in raw_u:
            return self._get_by_short("OPD OBGYNE")
        if "OB-GYNE WARD" in raw_u or "OBGYNE WARD" in raw_u or "OB WARD" in raw_u or "HIGH-RISK PREGNANCY" in raw_u:
            return self._get_by_short("OBWARD")
        if "OBSTETRICS" in raw_u or "OB-GYNE" in raw_u or "OBGYNE" in raw_u or "OB COMPLEX" in raw_u or "OB OPERATING" in raw_u or "FAMILY PLANNING" in raw_u or "WOMEN AND CHILD" in raw_u:
            return self._get_by_short("OB-GYNE")

        # Pediatrics variants
        if "OPD PED" in raw_u or "PEDIA OPD" in raw_u or "PEDIATRIC OPD" in raw_u:
            return self._get_by_short("OPD PED")
        if "PEDIA WARD" in raw_u or "PEDIATRIC WARD" in raw_u or "PEDWARD" in raw_u:
            return self._get_by_short("PEDWARD")
        if "PEDIA" in raw_u:
            return self._get_by_short("PEDIA")

        # Surgery & Ortho variants
        if "SURGERY OPD" in raw_u or "OPD SURG" in raw_u or "ORTHO-URO OPD" in raw_u or "ORTHO OPD" in raw_u:
            return self._get_by_short("OPDSURGE")
        if "SURGERY WARD" in raw_u or "SURGICAL WARD" in raw_u or "ORTHO WARD" in raw_u:
            return self._get_by_short("SURGWARD")
        if "SURGERY" in raw_u or "SURGICAL" in raw_u:
            return self._get_by_short("SURGERY")

        # Operating Room & PACU / Anesthesia
        if "PACU" in raw_u or "RECOVERY ROOM" in raw_u or "POST ANESTHESIA" in raw_u:
            return self._get_by_short("PACU")
        if "OPERATING ROOM" in raw_u or "OR-RR" in raw_u or "ANESTHESIA" in raw_u:
            return self._get_by_short("OR-RR")

        # Ophthalmology & ENT
        if "OPD OPHTH" in raw_u or "OPD OPTH" in raw_u or "OPHTHALMOLOGY OPD" in raw_u:
            return self._get_by_short("OPD OPTHA")
        if "OPHTHALMOLOGY WARD" in raw_u or "OPTHALMOLOGY WARD" in raw_u or "EENT" in raw_u or re.search(r"\bENT\b", raw_u):
            return self._get_by_short("EENTWARD")
        if "OPHTHALMOLOGY" in raw_u or "OPTHALMOLOGY" in raw_u:
            return self._get_by_short("OPTHA")

        # Medicine Ward & Internal Medicine sub-units
        if "MEDICINE OPD" in raw_u or "OPD MED" in raw_u:
            return self._get_by_short("OPDMED")
        if (
            "MEDICINE WARD" in raw_u
            or "COMMUNICABLE" in raw_u
            or "ISOLATION WARD" in raw_u
            or "TRANSITION WARD" in raw_u
            or "INFIRMARY WARD" in raw_u
        ):
            return self._get_by_short("MEDWARD")
        if "PAYWARD" in raw_u or "SUITE ROOM" in raw_u:
            return self._get_by_short("PW")
        if (
            "INTERNAL MEDICINE" in raw_u
            or "CANCER" in raw_u
            or "DERMATOLOGY" in raw_u
            or "TREATMENT HUB" in raw_u
            or "NORTHERN TAGALOG" in raw_u
        ):
            return self._get_by_short("MEDICINE")

        # Dialysis / Hemodialysis
        if "DIALYSIS" in raw_u or "HEMODIALYSIS" in raw_u or "HEMO" in raw_u:
            return self._get_by_short("HEMO")

        # Psychiatry
        if "PSYCHIATRY" in raw_u or "PSYCHIATRIC" in raw_u or "APU" in raw_u:
            return self._get_by_short("APU")

        # Family & Community Medicine / TB / Palliative / Telemedicine
        if "OPD FAM" in raw_u or "OPD INTEGRATED PALLIATIVE" in raw_u:
            return self._get_by_short("OPD FAMED")
        if "FAMILY" in raw_u or "COMMUNITY MEDICINE" in raw_u or "TB PROGRAM" in raw_u or "PALLIATIVE" in raw_u or "COMPLEMENTARY" in raw_u or "TELEMEDICINE" in raw_u:
            return self._get_by_short("DFM")

        # Radiology, Imaging, 2D Echo, ECG, Diagnostics
        if "2D ECHO" in raw_u or "ECG" in raw_u:
            return self._get_by_short("ECG")
        if (
            "CT-SCAN" in raw_u
            or "CT SCAN" in raw_u
            or "RADIOLOGY" in raw_u
            or "X-RAY" in raw_u
            or "ULTRASOUND" in raw_u
            or "MAMMOGRAPHY" in raw_u
            or "MAGNETIC RESONANCE" in raw_u
            or "DIAGNOSTIC" in raw_u
        ):
            return self._get_by_short("RADIOLOGY")

        # Laboratory & Pathology
        if "HISTOPATHOLOGY" in raw_u or "PATHOLOGY" in raw_u:
            return self._get_by_short("PATHO")
        if "LAB" in raw_u or "BLOOD BANK" in raw_u or "MICROBIOLOGY" in raw_u:
            return self._get_by_short("LAB")

        # Other Allied / Admin / Nursing / Finance sections in rs code.xls
        if "REHAB" in raw_u:
            return self._get_by_short("REHABMED")
        if "DIABETIC" in raw_u:
            return self._get_by_short("DIA")
        if "ADMITTING" in raw_u:
            return self._get_by_short("ADMITTING")
        if "HEALTH INFORMATION" in raw_u or "HIMS" in raw_u or "OPD RECORDS" in raw_u:
            return self._get_by_short("HIMS")
        if "SOCIAL SERVICE" in raw_u or "SWA" in raw_u:
            return self._get_by_short("SWA")
        if "NUTRITION" in raw_u or "DIETETICS" in raw_u:
            return self._get_by_short("NDS")
        if "CENTRAL SUPPLY" in raw_u or "CSR" in raw_u or "CSP" in raw_u:
            return self._get_by_short("CSP")
        if "NURSING ADMIN" in raw_u or "NURSE SUPERVISOR" in raw_u:
            return self._get_by_short("NSO")
        if "NURSING" in raw_u or "CHIEF NURSE" in raw_u:
            return self._get_by_short("CNO")
        if "ENGINEERING" in raw_u:
            return self._get_by_short("ENGINEERING")
        if "HOUSEKEEP" in raw_u or "HOUSE KEEPING" in raw_u:
            return self._get_by_short("HOUSEKEEP")
        if "LAUNDRY" in raw_u or "LINEN" in raw_u:
            return self._get_by_short("LAUNDRY/LINEN")
        if "BIOMED" in raw_u:
            return self._get_by_short("BIOMEDICAL")
        if "HUMAN RESOURCE" in raw_u or "HRM" in raw_u:
            return self._get_by_short("HRM")
        if "SECURITY" in raw_u:
            return self._get_by_short("SECURITY")
        if "MATERIALS MANAGEMENT" in raw_u or "MMS" in raw_u:
            return self._get_by_short("MMS")
        if "PROCUREMENT" in raw_u:
            return self._get_by_short("PS")
        if "CHIEF ADMINISTRATIVE" in raw_u:
            return self._get_by_short("CAO")
        if "ACCOUNTING" in raw_u:
            return self._get_by_short("ACCTG")
        if "BILLING" in raw_u:
            return self._get_by_short("BILLING")
        if "PHILHEALTH" in raw_u or "CLAIMS" in raw_u:
            return self._get_by_short("CLAIMS")
        if "BUDGET" in raw_u:
            return self._get_by_short("BUDGET")
        if "CASH" in raw_u:
            return self._get_by_short("CASHIER")
        if "FINANCE" in raw_u:
            return self._get_by_short("FMO")

        # 3. Substring scan against loaded `rs code.xls` entries
        for rcc, sc, fn in self.entries:
            if fn and (fn in raw_u or raw_u in fn):
                return rcc

        default_rcc = self._get_by_short("PHARMA")
        if track_default and raw_u and raw_u not in self.fallback_wards:
            self.fallback_wards[raw_u] = {"office": str(office_name).strip(), "rcc": default_rcc}
        return default_rcc
