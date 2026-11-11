import io
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock

import openpyxl

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from app.pipeline.ris_generator import (
    RIS_TARGET_PAGE_HEIGHT_PT,
    _compute_ris_page_heights,
    build_ris_workbook,
)
from app.pipeline.rsmi_generator import (
    RSMI_TARGET_PAGE_HEIGHT_PT,
    _compute_rsmi_page_heights,
    _write_rsmi_workbook,
    classify_fund_cluster,
    is_known_fund_prefix,
)
from app.pipeline.unit_lookup import UnitLookup
from app.pipeline.rcc_lookup import RccLookup
from app.pipeline.utils import safe_save_workbook
from app.server import RISAppHandler, parse_multipart_form, is_path_within_root


class TestFundClusterAndAudit(unittest.TestCase):
    def test_fund_cluster_prefixes(self):
        self.assertEqual(classify_fund_cluster("DMR000012"), "RAF")
        self.assertEqual(classify_fund_cluster("PRODMR000045"), "RAF")
        self.assertEqual(classify_fund_cluster("DMNP000101"), "RAF")

        self.assertEqual(classify_fund_cluster("DMEP000088"), "IGF")
        self.assertEqual(classify_fund_cluster("DMCSP000002"), "IGF")
        self.assertEqual(classify_fund_cluster("DMC000411"), "IGF")

        self.assertEqual(classify_fund_cluster("DMD000019"), "DON")
        self.assertEqual(classify_fund_cluster("DMDON000374"), "DON")

        self.assertTrue(is_known_fund_prefix("PRODMR000045"))
        self.assertFalse(is_known_fund_prefix("AFS000001"))

    def test_safe_save_workbook_captures_locked_files(self):
        locked = []
        mock_wb = MagicMock()
        mock_wb.save.side_effect = PermissionError("File is open")

        ok = safe_save_workbook(mock_wb, r"C:\tmp\RSMI - IGF.xlsx", locked_files=locked)
        self.assertFalse(ok)
        self.assertEqual(locked, ["RSMI - IGF.xlsx"])

        # Deduplicate repeated lock notices
        safe_save_workbook(mock_wb, r"C:\tmp\RSMI - IGF.xlsx", locked_files=locked)
        self.assertEqual(locked, ["RSMI - IGF.xlsx"])

    def test_unit_and_rcc_fallback_audit_tracking(self):
        ul = UnitLookup()
        self.assertEqual(ul.resolve_unit("DMR999999", "PARACETAMOL 500MG TABLET"), "Tablet")
        self.assertEqual(ul.resolve_unit("XYZ000001", "UNUSUAL MEDICAL GADGET"), "Piece")
        self.assertIn("DMR999999", ul.fallback_items)
        self.assertEqual(ul.fallback_items["DMR999999"]["reason"], "heuristic")
        self.assertEqual(ul.fallback_items["XYZ000001"]["reason"], "default_piece")

        rcc = RccLookup()
        rcc.resolve_rcc("PHARMACY")
        rcc.resolve_rcc("PEDIA WARD")
        self.assertEqual(len(rcc.fallback_wards), 0)

        rcc.resolve_rcc("MYSTERY TEST CLINIC 999")
        self.assertIn("MYSTERY TEST CLINIC 999", rcc.fallback_wards)


class TestDynamicHeightsAndZeroHyphens(unittest.TestCase):
    def test_ris_dynamic_page_heights(self):
        sparse_h, footer_h = _compute_ris_page_heights([1] * 27, header_pt=160.0)
        self.assertEqual(len(sparse_h), 27)
        total_sparse = 160.0 + sum(sparse_h) + sum(footer_h.values())
        self.assertAlmostEqual(total_sparse, RIS_TARGET_PAGE_HEIGHT_PT, delta=1.0)

        dense_h, dense_footer = _compute_ris_page_heights([4] * 27, header_pt=160.0)
        self.assertTrue(all(h >= 15.5 for h in dense_h))
        self.assertLessEqual(160.0 + sum(dense_h) + sum(dense_footer.values()), RIS_TARGET_PAGE_HEIGHT_PT + 2.0)

    def test_rsmi_dynamic_page_heights(self):
        sparse_info = [(1, False)] * 16
        m_h, r_h, sig_h = _compute_rsmi_page_heights(sparse_info, [1] * 16, is_first_page=False)
        self.assertEqual(len(m_h), 16)
        self.assertEqual(len(r_h), 16)
        total = 72.5 + 48.0 + 18.0 + sum(m_h) + 34.0 + 18.0 + sum(r_h) + 100.0 + sig_h
        self.assertAlmostEqual(total, RSMI_TARGET_PAGE_HEIGHT_PT, delta=1.0)

    def test_rsmi_workbook_no_hyphens_and_valid_formulas(self):
        sample_rows = [
            ("OCT-2026-PEDIA", "13-001-14-00012-03-04-13", "PRODMR000001", "AMOXICILLIN 500MG CAPSULE", "Capsule", 10),
            ("OCT-2026-PEDIA", "13-001-14-00012-03-04-13", "DMEP000100", "(EP) SALBUTAMOL 2MG TABLET", "Tablet", 5),
            ("OCT-2026-PEDIA", "13-001-14-00012-03-04-13", "DMDON000050", "(DON) VITAMIN C 500MG TABLET", "Tablet", 20),
            ("OCT-2026-PEDIA", "13-001-14-00012-03-04-13", "AFS000001", "ACCOUNTABLE FORM TEST", "Pad", 2),
        ]
        summary_rows = [{
            "ris_no": "OCT-2026-PEDIA",
            "office": "PEDIA WARD",
            "item_count": 4,
            "total_qty": 37,
            "filename": "PEDIA WARD_RIS_OCT.xlsx",
            "raf_qty": 10,
            "igf_qty": 7,
            "don_qty": 20,
        }]

        with tempfile.TemporaryDirectory() as tmpdir:
            final_dir = os.path.join(tmpdir, "02_RIS_FINAL")
            sum_dir = os.path.join(tmpdir, "04_RIS_COUNT_SUMMARY")
            os.makedirs(final_dir, exist_ok=True)
            out_xlsx = os.path.join(final_dir, "RSMI.xlsx")
            out_xls = os.path.join(final_dir, "RSMI.xls")
            sum_xlsx = os.path.join(sum_dir, "RIS_COUNT_SUMMARY.xlsx")
            sum_xls = os.path.join(sum_dir, "RIS_COUNT_SUMMARY.xls")
            res = _write_rsmi_workbook(
                sample_rows,
                summary_rows,
                out_xlsx,
                out_xls,
                summary_xlsx=sum_xlsx,
                summary_xls=sum_xls,
            )

            self.assertEqual(res["breakdown"]["RAF"]["items"], 1)
            self.assertEqual(res["breakdown"]["IGF"]["items"], 2)
            self.assertEqual(res["breakdown"]["DON"]["items"], 1)
            self.assertEqual([x["stock_no"] for x in res["unassigned_prefixes"]], ["AFS000001"])

            wb = openpyxl.load_workbook(out_xlsx, data_only=False)
            self.assertEqual(wb.sheetnames, ["RSMI - RAF", "RSMI - IGF", "RSMI - DON"])
            ws_raf = wb["RSMI - RAF"]
            self.assertFalse(ws_raf.oddFooter.right.text)
            # Check no hyphen placeholders in accounting cells J/K
            for r in range(12, ws_raf.max_row + 1):
                for c in (10, 11):
                    v = ws_raf.cell(r, c).value
                    self.assertNotEqual(str(v).strip(), "-")
            wb.close()

            wb_don = openpyxl.load_workbook(os.path.join(final_dir, "RSMI - DON.xlsx"), data_only=False)
            self.assertEqual(wb_don.sheetnames, ["RSMI"])
            self.assertFalse(wb_don["RSMI"].oddFooter.right.text)
            wb_don.close()

            self.assertTrue(os.path.exists(sum_xlsx))
            self.assertTrue(os.path.exists(sum_xls))
            wb_sum = openpyxl.load_workbook(sum_xlsx, data_only=False)
            self.assertEqual(wb_sum.sheetnames, ["RIS Count Summary"])
            self.assertGreaterEqual(wb_sum["RIS Count Summary"].row_dimensions[1].height, 34)
            wb_sum.close()

    def test_build_ris_workbook_creates_valid_sheet(self):
        ul = UnitLookup()
        with tempfile.TemporaryDirectory() as tmpdir:
            dest = os.path.join(tmpdir, "TEST_RIS.xlsx")
            build_ris_workbook(
                dest_path=dest,
                office_name="PEDIA WARD",
                ris_no_text="RIS No.  OCT-2026-PEDIA",
                items=[{"Item Code": "DMR000001", "Description": "AMOXICILLIN 500MG CAPSULE", "Quantity": 12}],
                unit_lookup=ul,
            )
            self.assertTrue(os.path.exists(dest))


class TestServerMultipartAndEndpoints(unittest.TestCase):
    def test_parse_multipart_form_handles_text_and_multiple_files(self):
        boundary = "----TestBoundary123"
        body = (
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="month_name"\r\n\r\nOCTOBER\r\n'
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="consumption_files"; filename="a.xls"\r\n'
            "Content-Type: application/vnd.ms-excel\r\n\r\nFILE_A\r\n"
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="consumption_files"; filename=""\r\n'
            "Content-Type: application/octet-stream\r\n\r\n\r\n"
            f"--{boundary}--\r\n"
        ).encode("utf-8")
        fields, files = parse_multipart_form(body, f"multipart/form-data; boundary={boundary}")
        self.assertEqual(fields["month_name"], "OCTOBER")
        self.assertEqual(len(files["consumption_files"]), 1)
        self.assertEqual(files["consumption_files"][0][0], "a.xls")
        self.assertEqual(files["consumption_files"][0][1], b"FILE_A")

    def test_path_guard_blocks_outside_root(self):
        self.assertTrue(is_path_within_root(os.path.join(ROOT_DIR, "OUTPUT_OCTOBER_2026")))
        self.assertFalse(is_path_within_root(r"C:\Windows\System32"))

    @staticmethod
    def _invoke_handler(method, path, body_bytes=b"", headers=None, parse_json=True):
        raw_out = io.BytesIO()
        handler = RISAppHandler.__new__(RISAppHandler)
        handler.path = path
        handler.request_version = "HTTP/1.1"
        handler.command = method
        handler.requestline = f"{method} {path} HTTP/1.1"
        handler.client_address = ("127.0.0.1", 5050)
        handler.rfile = io.BytesIO(body_bytes)
        handler.wfile = raw_out
        handler.headers = {"Content-Length": str(len(body_bytes)), **(headers or {})}
        handler.close_connection = True
        if method == "GET":
            handler.do_GET()
        else:
            handler.do_POST()
        raw_text = raw_out.getvalue().decode("utf-8", errors="replace")
        head, _, body = raw_text.partition("\r\n\r\n")
        status_code = int(head.splitlines()[0].split()[1])
        return status_code, (json.loads(body) if parse_json else body)

    def test_handler_endpoints_in_memory(self):
        status_html, raw_html = self._invoke_handler("GET", "/", parse_json=False)
        self.assertEqual(status_html, 200)
        self.assertIn("ARMMC Pharmacy", raw_html)
        self.assertIn("Export RSMI PDFs", raw_html)

        status, data = self._invoke_handler("GET", "/api/defaults")
        self.assertEqual(status, 200)
        self.assertIn("outputs", data)
        self.assertIn("libreoffice_available", data)

        status_p, p_data = self._invoke_handler("GET", "/api/progress?job_id=missing")
        self.assertEqual(status_p, 200)
        self.assertTrue(p_data["ok"])

        oct_final = os.path.join(ROOT_DIR, "OUTPUT_OCTOBER_2026", "02_RIS_FINAL")
        if os.path.isdir(oct_final):
            status_oct, oct_data = self._invoke_handler(
                "POST",
                "/api/load_output",
                body_bytes=json.dumps({"folder_name": "OUTPUT_OCTOBER_2026"}).encode("utf-8"),
                headers={"Content-Type": "application/json"},
            )
            self.assertEqual(status_oct, 200)
            self.assertTrue(oct_data["ok"])
            self.assertGreater(oct_data["result"]["rsmi"]["total_ris_files"], 0)

        payload = json.dumps({"final_dir": r"C:\Windows\System32"}).encode("utf-8")
        status_err, err_data = self._invoke_handler(
            "POST",
            "/api/load_output",
            body_bytes=payload,
            headers={"Content-Type": "application/json"},
        )
        self.assertEqual(status_err, 404)
        self.assertFalse(err_data["ok"])


if __name__ == "__main__":
    unittest.main()
