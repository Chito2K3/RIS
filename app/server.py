import os
import sys
import json
import shutil
import zipfile
import tempfile
import subprocess
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import cgi

# Add parent directory so `pipeline` package can be imported cleanly
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from pipeline.runner import run_full_pipeline
from pipeline.rsmi_generator import generate_rsmi_from_folder

DEFAULT_MASTERLIST = os.path.join(
    ROOT_DIR, "RIS_OCTOBER_FINAL", "INVENTORY UTILIZATION REPORT_MASTERLIST_2026_V30.xlsx"
)

LAST_RUN_STATE = {}

class RISAppHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        sys.stdout.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), format % args))

    def _send_json(self, data, status=200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            tpl_path = os.path.join(APP_DIR, "templates", "index.html")
            with open(tpl_path, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return

        if parsed.path == "/api/defaults":
            # Detect available files in workspace for quick 1-click loading or drag-and-drop
            workspace_files = sorted([
                f for f in os.listdir(ROOT_DIR)
                if os.path.isfile(os.path.join(ROOT_DIR, f)) and not f.startswith("~$")
            ])
            self._send_json({
                "root_dir": ROOT_DIR,
                "default_masterlist": DEFAULT_MASTERLIST if os.path.exists(DEFAULT_MASTERLIST) else "",
                "workspace_files": workspace_files,
            })
            return

        if parsed.path == "/api/download_zip":
            qs = parse_qs(parsed.query)
            folder = qs.get("folder", [""])[0]
            if not folder or not os.path.exists(folder):
                self._send_json({"error": "Folder not found"}, status=404)
                return
            tmp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
            tmp_zip.close()
            with zipfile.ZipFile(tmp_zip.name, "w", zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(folder):
                    for file in files:
                        full_p = os.path.join(root, file)
                        rel_p = os.path.relpath(full_p, folder)
                        zf.write(full_p, rel_p)
            with open(tmp_zip.name, "rb") as f:
                data = f.read()
            os.unlink(tmp_zip.name)
            zip_name = os.path.basename(os.path.normpath(folder)) + ".zip"
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", f'attachment; filename="{zip_name}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        self.send_error(404, "Not Found")

    def do_POST(self):
        global LAST_RUN_STATE
        parsed = urlparse(self.path)

        if parsed.path == "/api/open_folder":
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length).decode("utf-8")) if length > 0 else {}
            folder = os.path.normpath(body.get("folder", ""))
            if folder and os.path.exists(folder):
                if os.name == "nt":
                    subprocess.Popen(["explorer.exe", folder])
                self._send_json({"ok": True, "opened": folder})
            else:
                self._send_json({"ok": False, "error": f"Folder not found: {folder}"}, status=400)
            return

        if parsed.path == "/api/update_rsmi_exclusions":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = json.loads(self.rfile.read(length).decode("utf-8")) if length > 0 else {}
                final_dir = os.path.normpath(body.get("final_dir", ""))
                excluded_files = body.get("excluded_files", [])

                if not final_dir or not os.path.exists(final_dir):
                    self._send_json({"ok": False, "error": "Final RIS folder not found."}, status=400)
                    return

                excl_dir = os.path.join(final_dir, "_EXCLUDED_RIS")
                os.makedirs(excl_dir, exist_ok=True)

                moved_files = []
                for fname in excluded_files:
                    safe_fn = os.path.basename(fname)
                    src_p = os.path.join(final_dir, safe_fn)
                    dst_p = os.path.join(excl_dir, safe_fn)
                    if os.path.exists(src_p) and os.path.isfile(src_p):
                        shutil.move(src_p, dst_p)
                        moved_files.append(safe_fn)

                rsmi_info = generate_rsmi_from_folder(final_dir)
                excluded_on_disk = sorted([
                    f for f in os.listdir(excl_dir)
                    if f.lower().endswith(".xlsx") and not f.startswith("~$")
                ])
                self._send_json({
                    "ok": True,
                    "moved_files": moved_files,
                    "excluded_count_total": len(excluded_on_disk),
                    "rsmi": rsmi_info,
                })
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path == "/api/restore_exclusions":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = json.loads(self.rfile.read(length).decode("utf-8")) if length > 0 else {}
                final_dir = os.path.normpath(body.get("final_dir", ""))
                excl_dir = os.path.join(final_dir, "_EXCLUDED_RIS")
                restored = []
                if os.path.exists(excl_dir):
                    for fname in os.listdir(excl_dir):
                        src_p = os.path.join(excl_dir, fname)
                        dst_p = os.path.join(final_dir, fname)
                        if os.path.isfile(src_p):
                            shutil.move(src_p, dst_p)
                            restored.append(fname)
                rsmi_info = generate_rsmi_from_folder(final_dir)
                self._send_json({
                    "ok": True,
                    "restored_files": restored,
                    "excluded_count_total": 0,
                    "rsmi": rsmi_info,
                })
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path == "/api/process":
            try:
                form = cgi.FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ={
                        "REQUEST_METHOD": "POST",
                        "CONTENT_TYPE": self.headers.get("Content-Type"),
                    },
                )

                month_name = form.getvalue("month_name", "OCTOBER")
                year = form.getvalue("year", "2026")
                output_folder_name = form.getvalue("output_folder_name", f"OUTPUT_{month_name}_{year}")
                req_by_name = form.getvalue("req_by_name", "SABA, LEONARDO CONCHITO SEBASTIAN")
                req_by_desig = form.getvalue("req_by_desig", "Administrative Assistant II")
                app_by_name = form.getvalue("app_by_name", "SALVE CERTEZA")
                app_by_desig = form.getvalue("app_by_desig", "Pharmacist V")
                sig_date = form.getvalue("sig_date", "11/09/2026")
                use_sample_oct = form.getvalue("use_workspace_october", "false") == "true"

                output_base_dir = os.path.join(ROOT_DIR, output_folder_name)
                upload_staging_dir = os.path.join(output_base_dir, "00_RAW_UPLOADS")
                os.makedirs(upload_staging_dir, exist_ok=True)

                consump_paths = []
                issuance_paths = []
                masterlist_path = DEFAULT_MASTERLIST if os.path.exists(DEFAULT_MASTERLIST) else None

                if use_sample_oct:
                    # Use the workspace October 2026 raw files directly
                    for cand in ["CONSUMP_ANCILLIARY_OCT_2026.xls", "CONSUMP_SUMMARY_OCT_2026_ORIGINAL_BACKUP.xls", "CONS_CONSUMP_OCT_2026_ORIGINAL_BACKUP.xls"]:
                        p = os.path.join(ROOT_DIR, cand)
                        if os.path.exists(p):
                            consump_paths.append(p)
                    for cand in ["issuance_disp_oct_2026.csv", "issuance_storage_oct_2026.csv"]:
                        p = os.path.join(ROOT_DIR, cand)
                        if os.path.exists(p):
                            issuance_paths.append(p)
                else:
                    # Save uploaded consumption files
                    if "consumption_files" in form:
                        c_items = form["consumption_files"]
                        if not isinstance(c_items, list):
                            c_items = [c_items]
                        for item in c_items:
                            if getattr(item, "filename", None):
                                safe_fn = os.path.basename(item.filename)
                                dest_p = os.path.join(upload_staging_dir, safe_fn)
                                with open(dest_p, "wb") as f_out:
                                    shutil.copyfileobj(item.file, f_out)
                                consump_paths.append(dest_p)

                    # Save uploaded issuance files
                    if "issuance_files" in form:
                        i_items = form["issuance_files"]
                        if not isinstance(i_items, list):
                            i_items = [i_items]
                        for item in i_items:
                            if getattr(item, "filename", None):
                                safe_fn = os.path.basename(item.filename)
                                dest_p = os.path.join(upload_staging_dir, safe_fn)
                                with open(dest_p, "wb") as f_out:
                                    shutil.copyfileobj(item.file, f_out)
                                issuance_paths.append(dest_p)

                    # Optional custom masterlist upload
                    if "masterlist_file" in form:
                        m_item = form["masterlist_file"]
                        if getattr(m_item, "filename", None):
                            safe_fn = os.path.basename(m_item.filename)
                            dest_p = os.path.join(upload_staging_dir, safe_fn)
                            with open(dest_p, "wb") as f_out:
                                shutil.copyfileobj(m_item.file, f_out)
                            masterlist_path = dest_p

                if not consump_paths and not issuance_paths:
                    self._send_json({
                        "ok": False,
                        "error": "Please select at least one Consumption file or Issuance file (or click 'Run October 2026 Workspace Files')."
                    }, status=400)
                    return

                result = run_full_pipeline(
                    consumption_files=consump_paths,
                    issuance_files=issuance_paths,
                    masterlist_path=masterlist_path,
                    output_base_dir=output_base_dir,
                    month_name=month_name,
                    year=year,
                    req_by_name=req_by_name,
                    req_by_desig=req_by_desig,
                    app_by_name=app_by_name,
                    app_by_desig=app_by_desig,
                    sig_date=sig_date,
                )
                LAST_RUN_STATE = result
                self._send_json({"ok": True, "result": result})
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                self._send_json({"ok": False, "error": str(e), "traceback": tb}, status=500)
            return

        self.send_error(404, "Not Found")

def run_server(port=5050):
    server = HTTPServer(("127.0.0.1", port), RISAppHandler)
    print(f"============================================================")
    print(f" ARMMC Pharmacy RIS & RSMI Automation App Running!")
    print(f" Open in browser: http://127.0.0.1:{port}")
    print(f"============================================================")
    server.serve_forever()

if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5050
    run_server(port)

