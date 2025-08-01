import os
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from pyxlsb import open_workbook as open_xlsb
import openpyxl

# COM imports for expanding outlines in Excel
try:
    import pythoncom
    import win32com.client
except ImportError:
    win32com = None

def expand_all_outlines(xlsb_path):
    """
    Opens the .xlsb via COM, expands every outline/group, unhides any hidden rows,
    then saves. Requires Windows + pywin32.
    """
    if not win32com:
        return

    pythoncom.CoInitialize()
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    wb = xl.Workbooks.Open(os.path.abspath(xlsb_path))
    for ws in wb.Worksheets:
        try:
            ws.Outline.ShowLevels(RowLevels=256)
        except Exception:
            pass
        try:
            ws.Rows.Hidden = False
        except Exception:
            pass
    wb.Save()
    wb.Close(False)
    xl.Quit()
    pythoncom.CoUninitialize()


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXT = {'xlsb', 'xlsx', 'xls'}
def allowed_file(fname):
    return '.' in fname and fname.rsplit('.', 1)[1].lower() in ALLOWED_EXT


@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        # 1) get uploads
        mf = request.files.get('monthly_file')
        lf = request.files.get('employee_file')
        if not (mf and allowed_file(mf.filename)
                and lf and allowed_file(lf.filename)):
            return "Please upload a .xlsb (monthly) and a .xlsx/.xls (manager list).", 400

        # 2) save to disk
        path_monthly = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(mf.filename))
        path_list    = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(lf.filename))
        mf.save(path_monthly)
        lf.save(path_list)

        # 3) expand outlines so every row is visible
        expand_all_outlines(path_monthly)

        # 4) read the monthly .xlsb
        try:
            with open_xlsb(path_monthly) as wb:
                sheets = list(wb.sheets)
                sheet_name = next(
                    (s for s in sheets if 'missing' in s.lower() and 'time' in s.lower()),
                    None
                )
                if not sheet_name:
                    return f"No sheet matching 'Missing Time'. Found: {sheets}", 400
                ws = wb.get_sheet(sheet_name)
                all_rows = list(ws.rows())
        except Exception as e:
            return f"Error reading monthly sheet: {e}", 500

        # 5) detect header row
        header_idx = None
        headers = []
        for i, row in enumerate(all_rows):
            vals = [(cell.v or '').strip() for cell in row]
            has_level   = any('level' in v.lower() for v in vals)
            has_empname = any('emp' in v.lower() and 'name' in v.lower() for v in vals)
            has_email   = any('email' in v.lower() for v in vals)
            has_missing = any(all(x in v.lower() for x in ('sum','missing','time')) for v in vals)
            if has_level and has_empname and has_email and has_missing:
                header_idx = i
                headers = vals
                break

        if header_idx is None:
            return ("Couldn’t find header row containing "
                    "Level, Emp ID - Name, Email ID, and Sum of Missing Time."), 400

        data = [[cell.v for cell in row] for row in all_rows[header_idx+1:]]

        # 6) locate columns by fuzzy matching
        try:
            idx_level   = next(i for i,h in enumerate(headers) if 'level' in h.lower())
            idx_empinfo = next(i for i,h in enumerate(headers)
                               if 'emp' in h.lower() and 'name' in h.lower())
            idx_email   = next(i for i,h in enumerate(headers) if 'email' in h.lower())
            idx_missing = next(i for i,h in enumerate(headers)
                               if all(x in h.lower() for x in ('sum','missing','time')))
        except StopIteration:
            return ("Monthly sheet missing one of: Level, "
                    "Emp ID - Name, Email ID, or Sum of Missing Time."), 400

        # 7) read manager list and detect its header row by Email (and Name)
        try:
            wb2 = openpyxl.load_workbook(path_list, read_only=True, data_only=True)
            ws2 = wb2.active
            mgr_rows = list(ws2.iter_rows(values_only=True))
        except Exception as e:
            return f"Error reading manager list: {e}", 500

        hdr2 = None
        mgr_data = []
        for i, row in enumerate(mgr_rows):
            vals = [(v or '').strip() for v in row]
            if any('email' in v.lower() for v in vals):
                hdr2 = vals
                mgr_data = mgr_rows[i+1:]
                break
        if hdr2 is None:
            return "Manager list must have a column containing 'Email'.", 400

        # find manager-list columns
        idx_email_mgr = next(i for i,h in enumerate(hdr2) if 'email' in h.lower())
        idx_name_mgr  = next((i for i,h in enumerate(hdr2)
                              if 'name' in h.lower() and 'email' not in h.lower()), None)

        # build lookup sets
        mgr_emails = {
            str(r[idx_email_mgr]).strip().lower()
            for r in mgr_data
            if r[idx_email_mgr]
        }
        mgr_names = set()
        if idx_name_mgr is not None:
            mgr_names = {
                str(r[idx_name_mgr]).strip().lower()
                for r in mgr_data
                if r[idx_name_mgr]
            }

        # 8) filter & build results, tracking current_level
        results = []
        current_level = None

        for row in data:
            # If this row has a non-blank Level cell, treat it as a group header:
            lvl_cell = row[idx_level]
            if lvl_cell is not None and str(lvl_cell).strip():
                current_level = lvl_cell
                continue

            # extract email
            raw_email = row[idx_email]
            email_key = str(raw_email).strip().lower() if raw_email else ''

            # extract the "Name" part from "EmpID - Name"
            empinfo = str(row[idx_empinfo] or '')
            parts = empinfo.split('-', 1)
            name_part = parts[1].strip() if len(parts) > 1 else empinfo.strip()
            name_key = name_part.lower()

            # match by email or by name
            if email_key in mgr_emails or name_key in mgr_names:
                results.append({
                    'Level':        current_level,
                    'Name':         name_part,
                    'Email ID':     raw_email,
                    'Missing Time': row[idx_missing],
                })

        return render_template('result.html', entries=results)

    # GET → show upload form
    return render_template('upload.html')


if __name__ == '__main__':
    app.run(debug=True)
