from pathlib import Path
from datetime import datetime

import pandas as pd


class MailLog:

    def __init__(self, output_folder):

        self.output_folder = Path(output_folder)

        self.output_folder.mkdir(parents=True, exist_ok=True)

        # Timestamp được tạo MỘT LẦN cho toàn bộ phiên gửi
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

        self.log_file = (self.output_folder /
                         f"Báo cáo gửi mail_{timestamp}.xlsx")

        self.rows = []

    # ==================================================
    # GHI KẾT QUẢ
    # ==================================================

    def add(self, file_name, email, status, note=""):

        row = {
            "TEN_FILE": file_name,
            "EMAIL": email,
            "TRANG_THAI": status,
            "GHI_CHU": note,
            "THOI_GIAN": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        self.rows.append(row)

        self.save()

    # ==================================================
    # LƯU EXCEL
    # ==================================================

    def save(self):

        df = pd.DataFrame(self.rows,
                          columns=[
                              "TEN_FILE", "EMAIL", "TRANG_THAI", "GHI_CHU",
                              "THOI_GIAN"
                          ])

        df.to_excel(self.log_file, index=False)

    # ==================================================
    # LẤY ĐƯỜNG DẪN FILE LOG
    # ==================================================

    def get_file(self):

        return self.log_file


_current_log = None


def start_mail_log(output_folder):

    global _current_log

    _current_log = MailLog(output_folder)

    return _current_log


def write_mail_log(output_folder, file_name, email, status, note=""):

    global _current_log

    # Nếu chưa khởi tạo phiên log
    if _current_log is None:

        _current_log = MailLog(output_folder)

    _current_log.add(file_name=file_name,
                     email=email,
                     status=status,
                     note=note)


def get_mail_log_file():

    global _current_log

    if _current_log is None:

        return None

    return _current_log.get_file()
