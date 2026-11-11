import os
import re
import sys
import json
import glob
import zipfile
import datetime
import tempfile
import threading
import subprocess
from email import policy
from email.parser import BytesParser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# Add parent directory so `pipeline` package can be imported cleanly
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from pipeline.runner import run_full_pipeline
from pipeline.rsmi_generator import generate_rsmi_from_folder, inspect_existing_output_folder
from pipeline.pdf_exporter import is_libreoffice_available, export_workbooks_to_pdf

LAST_RUN_STATE = {}
PROGRESS_JOBS = {}
PROGRESS_LOCK = threading.Lock()


def set_job_progress(job_id, pct, stage, message, done=False, error=None):
    if not job_id:
        return
    with PROGRESS_LOCK:
        PROGRESS_JOBS[job_id] = {
            "pct": pct,
            "stage": stage,
            "message": message,
            "done": done,
            "error": error,
        }


def find_latest_masterlist(root_dir=ROOT_DIR):
    """Auto-detects the latest `*MASTERLIST*.xlsx` workbook in standard workspace folders."""
    search_dirs = [
        os.path.join(root_dir, "RIS_OCTOBER_FINAL"),
        root_dir,
    ]
    candidates = []
    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        for p in glob.glob(os.path.join(d, "*MASTERLIST*.xlsx")):
            bn = os.path.basename(p)
            if bn.startswith("~$") or not os.path.isfile(p):
                continue
            m = re.search(r'_V(\d+)', bn, re.IGNORECASE)
            ver = int(m.group(1)) if m else 0
            candidates.append((ver, os.path.getmtime(p), p))
    if not candidates:
        return ""
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2]


def is_path_within_root(target_path, root_dir=ROOT_DIR):
    try:
        root_abs = os.path.abspath(root_dir)
        target_abs = os.path.abspath(target_path)
        return os.path.commonpath([root_abs, target_abs]) == root_abs
    except Exception:
        return False


def parse_multipart_form(body_bytes, content_type):
    """Parses multipart/form-data using stdlib `email.parser` (Python 3.13+ safe)."""
    fields = {}
    files = {}
    if not content_type:
        return fields, files
    raw_msg = (
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("latin-1")
        + body_bytes
    )
    msg = BytesParser(policy=policy.default).parsebytes(raw_msg)
    if not msg.is_multipart():
        return fields, files
    for part in msg.iter_parts():
        name = part.get_param("name", header="Content-Disposition")
        filename = part.get_filename()
        if not name:
            continue
        payload = part.get_payload(decode=True) or b""
        if filename is not None:
            if filename:
                files.setdefault(name, []).append((os.path.basename(filename), payload))
        else:
            fields[name] = payload.decode("utf-8", errors="replace")
    return fields, files


def list_existing_outputs(root_dir=ROOT_DIR):
    outputs = []
    if not os.path.isdir(root_dir):
        return outputs
    for entry in os.listdir(root_dir):
        full_base = os.path.join(root_dir, entry)
        final_dir = os.path.join(full_base, "02_RIS_FINAL")
        if not os.path.isdir(full_base) or not os.path.isdir(final_dir):
            continue
        if not entry.upper().startswith("OUTPUT"):
            continue
        xlsx_files = [
            f for f in os.listdir(final_dir)
            if f.lower().endswith(".xlsx") and not f.startswith("~$")
        ]
        cleaned_dir = os.path.join(full_base, "01_CLEANED_DATA")
        pdf_dir = os.path.join(full_base, "03_PRINT_PDF")
        mtime = os.path.getmtime(final_dir)
        outputs.append({
            "name": entry,
            "output_base_dir": full_base,
            "final_dir": final_dir,
            "cleaned_dir": cleaned_dir if os.path.isdir(cleaned_dir) else final_dir,
            "pdf_dir": pdf_dir if os.path.isdir(pdf_dir) else "",
            "xlsx_count": len(xlsx_files),
            "has_rsmi": os.path.exists(os.path.join(final_dir, "RSMI.xlsx")),
            "modified": datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %I:%M %p"),
            "_mtime": mtime,
        })
    outputs.sort(key=lambda x: x["_mtime"], reverse=True)
    for item in outputs:
        item.pop("_mtime", None)
    return outputs


def _mime_for_filename(fname):
    ext = os.path.splitext(fname)[1].lower()
    if ext == ".pdf":
        return "application/pdf"
    if ext == ".xls":
        return "application/vnd.ms-excel"
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class RISAppHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        try:
            if sys.stdout is not None:
                sys.stdout.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), format % args))
                sys.stdout.flush()
        except Exception:
            pass

    def _send_json(self, data, status=200):
        body = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length <= 0:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

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
            latest_ml = find_latest_masterlist()
            workspace_files = sorted([
                f for f in os.listdir(ROOT_DIR)
                if os.path.isfile(os.path.join(ROOT_DIR, f)) and not f.startswith("~$")
            ])
            self._send_json({
                "root_dir": ROOT_DIR,
                "default_masterlist": latest_ml,
                "default_masterlist_name": os.path.basename(latest_ml) if latest_ml else "",
                "workspace_files": workspace_files,
                "libreoffice_available": is_libreoffice_available(),
                "outputs": list_existing_outputs(),
            })
            return

        if parsed.path == "/api/outputs":
            self._send_json({"ok": True, "outputs": list_existing_outputs()})
            return

        if parsed.path == "/api/progress":
            qs = parse_qs(parsed.query)
            job_id = qs.get("job_id", [""])[0]
            with PROGRESS_LOCK:
                info = PROGRESS_JOBS.get(job_id, {"pct": 0, "stage": 0, "message": "", "done": False})
            self._send_json({"ok": True, **info})
            return

        if parsed.path in ("/api/download_rsmi", "/api/download_file"):
            qs = parse_qs(parsed.query)
            folder = qs.get("folder", [""])[0]
            fmt = qs.get("fmt", ["xlsx"])[0]
            custom_file = qs.get("file", [""])[0]
            fname = os.path.basename(custom_file) if custom_file else ("RSMI.xls" if fmt == "xls" else "RSMI.xlsx")
            fpath = os.path.join(folder, fname) if folder else ""
            if not folder or not is_path_within_root(folder) or not os.path.exists(fpath):
                self._send_json({"error": f"{fname} not found"}, status=404)
                return
            with open(fpath, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", _mime_for_filename(fname))
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        if parsed.path == "/api/download_zip":
            qs = parse_qs(parsed.query)
            folder = qs.get("folder", [""])[0]
            if not folder or not is_path_within_root(folder) or not os.path.exists(folder):
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

        if parsed.path == "/api/shutdown":
            self._send_json({"ok": True, "message": "Shutting down server..."})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return

        self._send_json({"ok": False, "error": f"Unknown GET endpoint: {parsed.path}"}, status=404)

    def do_POST(self):
        global LAST_RUN_STATE
        parsed = urlparse(self.path)

        if parsed.path == "/api/open_folder":
            body = self._read_json_body()
            folder = os.path.normpath(body.get("folder", ""))
            if folder and is_path_within_root(folder) and os.path.exists(folder):
                if os.name == "nt":
                    subprocess.Popen(["explorer.exe", folder])
                self._send_json({"ok": True, "opened": folder})
            else:
                self._send_json({"ok": False, "error": f"Folder not found: {folder}"}, status=400)
            return

        if parsed.path == "/api/load_output":
            try:
                body = self._read_json_body()
                folder_name = body.get("folder_name", "")
                final_dir = body.get("final_dir", "")
                if folder_name and not final_dir:
                    final_dir = os.path.join(ROOT_DIR, os.path.basename(folder_name), "02_RIS_FINAL")
                final_dir = os.path.normpath(final_dir)
                if not final_dir or not is_path_within_root(final_dir) or not os.path.isdir(final_dir):
                    self._send_json({"ok": False, "error": "Output folder not found."}, status=404)
                    return
                res = inspect_existing_output_folder(final_dir)
                LAST_RUN_STATE = res
                self._send_json({"ok": True, "result": res})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path == "/api/export_pdfs":
            job_id = ""
            try:
                body = self._read_json_body()
                final_dir = os.path.normpath(body.get("final_dir", ""))
                scope = body.get("scope", "rsmi")
                job_id = body.get("job_id", "")
                if not final_dir or not is_path_within_root(final_dir) or not os.path.isdir(final_dir):
                    self._send_json({"ok": False, "error": "Final RIS folder not found."}, status=400)
                    return

                output_base_dir = os.path.dirname(final_dir)
                pdf_dir = os.path.join(output_base_dir, "03_PRINT_PDF")
                all_xlsx = sorted([
                    f for f in os.listdir(final_dir)
                    if f.lower().endswith(".xlsx") and not f.startswith("~$")
                    and os.path.isfile(os.path.join(final_dir, f))
                ])
                rsmi_targets = [f for f in all_xlsx if f.upper().startswith("RSMI")]
                ris_targets = [f for f in all_xlsx if not f.upper().startswith("RSMI")]

                if scope == "rsmi":
                    selected = rsmi_targets
                elif scope == "ris":
                    selected = ris_targets
                else:
                    selected = rsmi_targets + ris_targets

                target_paths = [os.path.join(final_dir, f) for f in selected]
                set_job_progress(job_id, 5, 4, f"Preparing {len(target_paths)} workbook(s) for PDF export...")

                def _pdf_cb(pct, msg):
                    set_job_progress(job_id, pct, 4, msg)

                res = export_workbooks_to_pdf(target_paths, pdf_dir, progress_cb=_pdf_cb)
                set_job_progress(job_id, 100, 4, f"Exported {res['count']} PDF file(s) to 03_PRINT_PDF/", done=True)
                self._send_json({"ok": True, **res})
            except Exception as e:
                set_job_progress(job_id, 100, 4, str(e), done=True, error=str(e))
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path in ("/api/update_rsmi_exclusions", "/api/save_overwrite_rsmi"):
            try:
                body = self._read_json_body()
                final_dir = os.path.normpath(body.get("final_dir", ""))
                excluded_files = body.get("excluded_files", [])

                if not final_dir or not is_path_within_root(final_dir) or not os.path.exists(final_dir):
                    self._send_json({"ok": False, "error": "Final RIS folder not found."}, status=400)
                    return

                excl_state_path = os.path.join(final_dir, "_rsmi_excluded.json")
                existing_excl = set()
                if os.path.exists(excl_state_path):
                    try:
                        with open(excl_state_path, "r", encoding="utf-8") as f:
                            existing_excl = set(json.load(f))
                    except Exception:
                        pass

                for f in excluded_files:
                    existing_excl.add(os.path.basename(f))

                if any("PHARMACY_RIS_" in f.upper() or "POPULATED" in f.upper() for f in existing_excl):
                    for f in os.listdir(final_dir):
                        if ("PHARMACY_RIS_" in f.upper() or "POPULATED" in f.upper()) and f.lower().endswith(".xlsx"):
                            existing_excl.add(f)

                with open(excl_state_path, "w", encoding="utf-8") as f:
                    json.dump(sorted(existing_excl), f)

                import importlib
                import pipeline.ris_generator
                import pipeline.rsmi_generator
                import pipeline.runner
                importlib.reload(pipeline.ris_generator)
                importlib.reload(pipeline.rsmi_generator)
                importlib.reload(pipeline.runner)

                rsmi_info = pipeline.rsmi_generator.generate_rsmi_from_folder(final_dir, excluded_files=list(existing_excl))
                saved_at = datetime.datetime.now().strftime("%I:%M:%S %p")
                self._send_json({
                    "ok": True,
                    "moved_files": sorted(existing_excl),
                    "excluded_count_total": len(existing_excl),
                    "saved_at": saved_at,
                    "rsmi": rsmi_info,
                })
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return

        if parsed.path == "/api/restore_exclusions":
            try:
                body = self._read_json_body()
                final_dir = os.path.normpath(body.get("final_dir", ""))
                if not final_dir or not is_path_within_root(final_dir) or not os.path.isdir(final_dir):
                    self._send_json({"ok": False, "error": "Final RIS folder not found."}, status=400)
                    return

                excl_dir = os.path.join(final_dir, "_EXCLUDED_RIS")
                restored = []
                if os.path.exists(excl_dir):
                    for fname in os.listdir(excl_dir):
                        src_p = os.path.join(excl_dir, fname)
                        dst_p = os.path.join(final_dir, fname)
                        if os.path.isfile(src_p):
                            import shutil
                            shutil.move(src_p, dst_p)
                            restored.append(fname)
                excl_state_path = os.path.join(final_dir, "_rsmi_excluded.json")
                if os.path.exists(excl_state_path):
                    try:
                        with open(excl_state_path, "r", encoding="utf-8") as f:
                            restored.extend(json.load(f))
                        os.remove(excl_state_path)
                    except Exception:
                        pass

                import importlib
                import pipeline.ris_generator
                import pipeline.rsmi_generator
                import pipeline.runner
                importlib.reload(pipeline.ris_generator)
                importlib.reload(pipeline.rsmi_generator)
                importlib.reload(pipeline.runner)

                rsmi_info = pipeline.rsmi_generator.generate_rsmi_from_folder(final_dir)
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
            job_id = ""
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                raw_body = self.rfile.read(content_length) if content_length > 0 else b""
                fields, uploaded_files = parse_multipart_form(
                    raw_body,
                    self.headers.get("Content-Type", ""),
                )

                job_id = fields.get("job_id", "")
                month_name = fields.get("month_name", "OCTOBER")
                year = fields.get("year", "2026")
                safe_out_name = os.path.basename(fields.get("output_folder_name") or f"OUTPUT_{month_name}_{year}")
                req_by_name = fields.get("req_by_name", "SABA, LEONARDO CONCHITO SEBASTIAN")
                req_by_desig = fields.get("req_by_desig", "Administrative Assistant II")
                app_by_name = fields.get("app_by_name", "SALVE CERTEZA")
                app_by_desig = fields.get("app_by_desig", "Pharmacist V")
                sig_date = fields.get("sig_date", datetime.date.today().strftime("%m/%d/%Y"))
                rsmi_custodian_name = fields.get("rsmi_custodian_name", "").strip() or app_by_name
                use_sample_oct = fields.get("use_workspace_october", "false") == "true"
                mirror_outputs = fields.get("mirror_outputs", "false") == "true"

                output_base_dir = os.path.join(ROOT_DIR, safe_out_name)
                upload_staging_dir = os.path.join(output_base_dir, "00_RAW_UPLOADS")
                os.makedirs(upload_staging_dir, exist_ok=True)

                consump_paths = []
                issuance_paths = []
                masterlist_path = find_latest_masterlist() or None

                if use_sample_oct:
                    for cand in [
                        "CONSUMP_ANCILLIARY_OCT_2026.xls",
                        "CONSUMP_SUMMARY_OCT_2026_ORIGINAL_BACKUP.xls",
                        "CONS_CONSUMP_OCT_2026_ORIGINAL_BACKUP.xls",
                    ]:
                        p = os.path.join(ROOT_DIR, cand)
                        if os.path.exists(p):
                            consump_paths.append(p)
                    for cand in ["issuance_disp_oct_2026.csv", "issuance_storage_oct_2026.csv"]:
                        p = os.path.join(ROOT_DIR, cand)
                        if os.path.exists(p):
                            issuance_paths.append(p)
                else:
                    for safe_fn, payload in uploaded_files.get("consumption_files", []):
                        dest_p = os.path.join(upload_staging_dir, safe_fn)
                        with open(dest_p, "wb") as f_out:
                            f_out.write(payload)
                        consump_paths.append(dest_p)

                    for safe_fn, payload in uploaded_files.get("issuance_files", []):
                        dest_p = os.path.join(upload_staging_dir, safe_fn)
                        with open(dest_p, "wb") as f_out:
                            f_out.write(payload)
                        issuance_paths.append(dest_p)

                    ml_uploads = uploaded_files.get("masterlist_file", [])
                    if ml_uploads:
                        safe_fn, payload = ml_uploads[0]
                        dest_p = os.path.join(upload_staging_dir, safe_fn)
                        with open(dest_p, "wb") as f_out:
                            f_out.write(payload)
                        masterlist_path = dest_p

                if not consump_paths and not issuance_paths:
                    self._send_json({
                        "ok": False,
                        "error": "Please select at least one Consumption file or Issuance file (or click 'Quick-Run October 2026 Workspace Files').",
                    }, status=400)
                    return

                import importlib
                import pipeline.ris_generator
                import pipeline.rsmi_generator
                import pipeline.runner
                importlib.reload(pipeline.ris_generator)
                importlib.reload(pipeline.rsmi_generator)
                importlib.reload(pipeline.runner)

                def _progress_cb(pct, stage, msg):
                    set_job_progress(job_id, pct, stage, msg)

                set_job_progress(job_id, 5, 1, "Initializing pipeline...")
                result = pipeline.runner.run_full_pipeline(
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
                    rsmi_custodian_name=rsmi_custodian_name,
                    progress_callback=_progress_cb,
                    mirror_outputs=mirror_outputs,
                )
                LAST_RUN_STATE = result
                set_job_progress(job_id, 100, 4, "Completed! All files processed and saved.", done=True)
                self._send_json({"ok": True, "result": result})
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                set_job_progress(job_id, 100, 4, str(e), done=True, error=str(e))
                self._send_json({"ok": False, "error": str(e), "traceback": tb}, status=500)
            return

        if parsed.path == "/api/shutdown":
            self._send_json({"ok": True, "message": "Shutting down server..."})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return

        self._send_json({"ok": False, "error": f"Unknown POST endpoint: {parsed.path}"}, status=404)


def run_server(port=5050, open_browser=False):
    class ReusableThreadingHTTPServer(ThreadingHTTPServer):
        allow_reuse_address = True

    server = ReusableThreadingHTTPServer(("127.0.0.1", port), RISAppHandler)
    url = f"http://127.0.0.1:{port}"
    try:
        if sys.stdout is not None:
            print("============================================================", flush=True)
            print(" ARMMC Pharmacy RIS & RSMI Automation App Running!", flush=True)
            print(f" Open in browser: {url}", flush=True)
            print("============================================================", flush=True)
    except Exception:
        pass
    if open_browser:
        import webbrowser
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    open_browser = "--open" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    port = int(args[0]) if len(args) > 0 else 5050
    run_server(port, open_browser=open_browser)
