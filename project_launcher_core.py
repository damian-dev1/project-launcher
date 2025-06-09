import subprocess
import xlwings as xw
from pathlib import Path
import datetime

EXCEL_PATH = Path(r"C:/Users/damian/projects/tests/project_dashboard.xlsm")
SCRIPT_PATH = Path(r"C:/Users/damian/projects/newproject/projectgen.py")


def load_dropdown_values():
    app = xw.App(visible=False)
    wb = app.books.open(str(EXCEL_PATH))
    ws = wb.sheets["DataValidation"]

    values = {}
    headers = ["Type", "Language", "Status"]
    for idx, header in enumerate(headers, start=1):
        col_range = ws.range((2, idx), ws.range((2, idx)).end("down"))
        col_values = [str(c.value).strip() for c in col_range if c.value]
        values[header] = col_values

    app.quit()
    return values

def log_to_excel(name, type_, lang, status, tags, dev, output):
    try:
        # Attach to the already open Excel workbook
        wb = xw.books["project_dashboard.xlsm"]
    except Exception:
        # If not open, open it in background mode
        app = xw.App(visible=False)
        wb = app.books.open(str(EXCEL_PATH))

    ws = wb.sheets["Dashboard"]
    last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row + 1
    ws.range(f"A{last_row}").value = [
        datetime.datetime.now(), name, type_, lang, status, tags, dev, output
    ]

    wb.save()

    # DO NOT CLOSE workbook — leave it open if already open
    if "app" in locals():
        app.quit()  # Only quit if we started a hidden instance



def generate_tags(name, type_, lang):
    parts = set(name.lower().split("-") + [type_.lower(), lang.lower()])
    return ",".join(sorted(parts))


def run_projectgen(name, type_, lang, status, dev, tags):
    cmd = [
        "python", str(SCRIPT_PATH),
        name,
        "--type", type_,
        "--lang", lang,
        "--status", status,
        "--tags", tags,
        "--dev", dev
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def archive_dashboard_rows(excel_path: Path):
    import xlwings as xw
    from datetime import datetime

    app = xw.App(visible=False)
    wb = app.books.open(str(excel_path))

    try:
        dashboard = wb.sheets["Dashboard"]
        archive = None

        # Find or create Archive sheet
        if "Archive" in [s.name for s in wb.sheets]:
            archive = wb.sheets["Archive"]
        else:
            archive = wb.sheets.add("Archive")

        # Read data from Dashboard
        data = dashboard.range("A2").expand("table").value
        if not data:
            print("[INFO] Dashboard is empty, nothing to archive.")
            return

        # Ensure it's in list-of-lists format
        if not isinstance(data[0], list):
            data = [data]

        # Add Date Archived column (timestamp) to each row
        archive_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        archived_rows = [[archive_date] + row for row in data]

        # Find where to append in Archive sheet
        last_row = archive.range("A" + str(archive.cells.last_cell.row)).end("up").row
        if last_row == 1 and not archive.range("A1").value:
            # Add headers if first time
            archive.range("A1").value = [
                "Date Archived", "Created", "Project Name", "Type", "Language",
                "Status", "Tags", "Folder Path", "Notes"
            ]
            last_row = 1

        # Write rows
        archive.range(f"A{last_row + 1}").value = archived_rows
        wb.save()
        print(f"[INFO] Archived {len(archived_rows)} rows to Archive sheet.")

    except Exception as e:
        print(f"[ERROR] Failed to archive Dashboard: {e}")

    finally:
        wb.close()
        app.quit()
