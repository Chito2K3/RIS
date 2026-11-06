import os
import re
import openpyxl

class UnitLookup:
    def __init__(self, masterlist_path=None):
        self.unit_map = {}
        self.desc_map = {}
        if masterlist_path and os.path.exists(masterlist_path):
            self.load_masterlist(masterlist_path)

    @staticmethod
    def clean_desc(d):
        d = re.sub(r'^\([A-Z]+\)\s*', '', str(d or '').strip().upper())
        return re.sub(r'\s+', ' ', d)

    def load_masterlist(self, masterlist_path):
        wb_ml = openpyxl.load_workbook(masterlist_path, read_only=True, data_only=True)
        sheet_name = "MASTERLIST_2026" if "MASTERLIST_2026" in wb_ml.sheetnames else wb_ml.sheetnames[0]
        ws_ml = wb_ml[sheet_name]
        for row in ws_ml.iter_rows(min_row=6, max_row=3500, max_col=8, values_only=True):
            if not row or len(row) < 8:
                continue
            c_long = str(row[1] or "").strip()
            dci = str(row[2] or "").strip()
            desc = str(row[3] or "").strip()
            uom = str(row[7] or "").strip()
            if uom and uom != "None":
                if dci:
                    self.unit_map[dci] = uom
                if c_long:
                    self.unit_map[c_long] = uom
                if desc:
                    self.desc_map[self.clean_desc(desc)] = uom
        wb_ml.close()

    def resolve_unit(self, code, desc):
        code_str = str(code or "").strip()
        desc_str = str(desc or "").strip()
        if code_str in self.unit_map:
            return self.unit_map[code_str]
        cd = self.clean_desc(desc_str)
        if cd in self.desc_map:
            return self.desc_map[cd]
        u_desc = desc_str.upper()
        if "PRESCRIPTION PAD" in u_desc or "PAD" in u_desc:
            return "Pad"
        if "PRE-FILLED SYRINGE" in u_desc or "PRE‐FILLED SYRINGE" in u_desc or "PREFILLED SYRINGE" in u_desc:
            return "PFS"
        if "BOTTLE/BAG" in u_desc:
            return "Bottle/bag"
        if "AMPULE/VIAL" in u_desc:
            return "Ampule"
        if "TABLET/CAPSULE" in u_desc or "CAPSULE/TABLET" in u_desc:
            return "Capsule/Tablet"
        if "TABLET" in u_desc:
            return "Tablet"
        if "CAPSULE" in u_desc:
            return "Capsule"
        if "AMPULE" in u_desc:
            return "Ampule"
        if "VIAL" in u_desc or "LYOPHILIZED POWDER" in u_desc or "IFOSFAMIDE" in u_desc or "METHYLPREDNISOLONE 500 MG IV" in u_desc:
            return "Vial"
        if "BOTTLE" in u_desc or "DROPS" in u_desc or "SYRUP" in u_desc or "SUSPENSION" in u_desc:
            return "Bottle"
        if "TUBE" in u_desc or "OINTMENT" in u_desc or "CREAM" in u_desc or "GEL" in u_desc:
            return "Tube"
        if "NEBULE" in u_desc:
            return "Nebule"
        if "SACHET" in u_desc:
            return "Sachet"
        if "SUPPOSITORY" in u_desc or "DULCOLAX" in u_desc:
            return "Suppository"
        if "1CC" in u_desc or "0.2CC" in u_desc or " CC" in u_desc:
            return "cc"
        return "Piece"

