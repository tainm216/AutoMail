from pathlib import Path
from datetime import datetime
import pandas as pd


class MailLogger:

    def __init__(self, output_folder):

        self.output_folder = Path(output_folder)

        self.log_file = (self.output_folder / "BaoCaoGuiThongBao.xlsx")

        self.rows = []

    def add(self, ma_kh, ten_kh, email, pdf_file, status):

        self.rows.append({
            "Thời gian":
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Mã KH":
            ma_kh,
            "Tên KH":
            ten_kh,
            "Email":
            email,
            "File PDF":
            Path(pdf_file).name,
            "Trạng thái":
            status
        })

    def save(self):

        df = pd.DataFrame(self.rows)

        df.to_excel(self.log_file, index=False)
