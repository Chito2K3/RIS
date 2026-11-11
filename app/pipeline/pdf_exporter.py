import os
import shutil
import subprocess
import tempfile
from pathlib import Path


def find_libreoffice():
    """
    Locates LibreOffice headless binary (`soffice.com` preferred on Windows,
    falling back to `soffice` on PATH). Returns full path string or None.
    """
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
        shutil.which("soffice.com"),
        shutil.which("soffice"),
    ]
    for cand in candidates:
        if cand and os.path.exists(cand):
            return cand
    return None


def is_libreoffice_available():
    return find_libreoffice() is not None


def export_workbooks_to_pdf(xlsx_paths, pdf_dir, progress_cb=None, batch_size=10):
    """
    Converts `.xlsx` workbooks in `xlsx_paths` into print-ready A4 Portrait `.pdf`
    files inside `pdf_dir` using headless LibreOffice. Original `.xlsx`/`.xls` files
    remain untouched. Returns summary dict with generated PDF filenames & paths.
    """
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice (soffice.com) was not found on this PC. Install LibreOffice to enable one-click PDF export."
        )

    valid_paths = [os.path.abspath(p) for p in xlsx_paths if p and os.path.exists(p)]
    os.makedirs(pdf_dir, exist_ok=True)
    if not valid_paths:
        return {"pdf_dir": pdf_dir, "generated": [], "count": 0}

    generated = []
    total = len(valid_paths)

    with tempfile.TemporaryDirectory(prefix="armmc_lo_") as lo_profile:
        lo_uri = Path(lo_profile).as_uri()
        for start in range(0, total, batch_size):
            chunk = valid_paths[start : start + batch_size]
            if progress_cb:
                done = min(start, total)
                progress_cb(
                    int(10 + 85 * (done / max(total, 1))),
                    f"Exporting print PDFs ({done}/{total})...",
                )

            cmd = [
                soffice,
                "--headless",
                "--norestore",
                f"-env:UserInstallation={lo_uri}",
                "--convert-to",
                "pdf",
                "--outdir",
                os.path.abspath(pdf_dir),
                *chunk,
            ]
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=180,
            )
            if proc.returncode != 0:
                err_msg = (proc.stderr or proc.stdout or "").strip()
                raise RuntimeError(f"LibreOffice PDF conversion failed (exit {proc.returncode}): {err_msg}")

            for src in chunk:
                pdf_name = os.path.splitext(os.path.basename(src))[0] + ".pdf"
                pdf_path = os.path.join(pdf_dir, pdf_name)
                if os.path.exists(pdf_path):
                    generated.append({
                        "filename": pdf_name,
                        "path": pdf_path,
                        "source_xlsx": os.path.basename(src),
                        "size_kb": round(os.path.getsize(pdf_path) / 1024.0, 1),
                    })

    if progress_cb:
        progress_cb(100, f"Exported {len(generated)} print-ready PDF(s).")

    return {
        "pdf_dir": pdf_dir,
        "generated": generated,
        "count": len(generated),
    }

