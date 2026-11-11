import os


def safe_save_workbook(wb, dest_path, locked_files=None):
    """
    Saves an openpyxl Workbook to `dest_path`, capturing any PermissionError
    (e.g. file open in Excel or LibreOffice Calc) into `locked_files` instead
    of failing silently. Returns True if saved, False if locked.
    """
    try:
        wb.save(dest_path)
        return True
    except PermissionError:
        if locked_files is not None:
            fname = os.path.basename(dest_path)
            if fname not in locked_files:
                locked_files.append(fname)
        return False

