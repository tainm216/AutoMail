import win32com.client
from pathlib import Path


def load_excel(file_path):

    excel = None
    wb = None

    try:

        excel = win32com.client.DispatchEx("Excel.Application")

        excel.Visible = False
        excel.DisplayAlerts = False

        full_path = str(Path(file_path).resolve())

        print(full_path)

        wb = excel.Workbooks.Open(full_path, ReadOnly=True)

        # ==========================================
        # KIỂM TRA SHEET DATA
        # ==========================================

        sheet_names = [str(ws.Name).strip() for ws in wb.Worksheets]

        if "data" not in sheet_names:

            raise ValueError('File Excel không có sheet "data".')

        ws = wb.Worksheets("data")

        # ==========================================
        # ĐỌC DỮ LIỆU
        # ==========================================

        rows = []

        used_range = ws.UsedRange

        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        headers = []

        for col in range(1, col_count + 1):

            header = ws.Cells(1, col).Text

            headers.append(str(header).strip())

        for row in range(2, row_count + 1):

            row_data = {}

            for col in range(1, col_count + 1):

                key = headers[col - 1]

                value = ws.Cells(row, col).Text

                row_data[key] = (str(value).strip())

            rows.append(row_data)

        return rows

    finally:

        # ==========================================
        # LUÔN ĐÓNG FILE EXCEL
        # ==========================================

        if wb is not None:

            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

        # ==========================================
        # LUÔN THOÁT EXCEL INSTANCE CỦA AUTOMAIL
        # ==========================================

        if excel is not None:

            try:
                excel.Quit()
            except Exception:
                pass
