import shutil
import win32com.client
from pathlib import Path
from excel_reader import load_excel
from datetime import datetime


def get_output_folder():

    now = datetime.now()

    year = now.strftime("%Y")
    month = now.strftime("%m")
    day = now.strftime("%d")

    folder = Path("output") / year / month / day

    folder.mkdir(parents=True, exist_ok=True)

    return folder


class WordTemplate:

    def __init__(self, visible=False, logger=None):

        self.logger = logger

        self.word = win32com.client.DispatchEx("Word.Application")

        self.word.Visible = visible

    def close(self):

        self.word.Quit()

    def replace_placeholder(self, doc, placeholder, value):

        rng = doc.Content

        find = rng.Find

        while find.Execute(FindText=placeholder):

            rng.Text = str(value)

            rng = doc.Range(rng.End, doc.Content.End)

            find = rng.Find

    def generate_docx(self, template_file, data, output_docx):

        template_file = Path(template_file).resolve()
        output_docx = Path(output_docx).resolve()

        shutil.copy(str(template_file), str(output_docx))

        doc = self.word.Documents.Open(str(output_docx))

        for key, value in data.items():

            placeholder = f"{{{{{key}}}}}"

            self.replace_placeholder(doc, placeholder, value)

            self.log(f"Đã thay key {{{{{key}}}}} bằng giá trị {value}")

        doc.Save()
        doc.Close()

        return str(output_docx)

    def docx_to_pdf(self, docx_file, pdf_file=None):

        docx_file = Path(docx_file).resolve()

        if pdf_file is None:

            pdf_file = (output_folder / f"{ma_kh}_{now:%Y%m%d}.pdf")

        pdf_file = Path(pdf_file).resolve()

        doc = self.word.Documents.Open(str(docx_file))

        doc.SaveAs(str(pdf_file), FileFormat=17)

        doc.Close(False)

        return str(pdf_file)

    def log(self, message):

        if self.logger:

            try:
                self.logger.info(message)

            except:
                print(message)

        else:

            print(message)


if __name__ == "__main__":

    rows = load_excel("data.xlsx")

    self.log(f"Tổng số khách hàng: {len(rows)}")

    engine = WordTemplate()

    try:

        output_folder = get_output_folder()

        self.log(f"Output folder: {output_folder}")

        for row in rows:

            ma_kh = row["MA_KHANG"]

            today = datetime.now().strftime("%Y%m%d")

            docx_file = (output_folder / f"{ma_kh}_{today}.docx")

            pdf_file = (output_folder / f"{ma_kh}_{today}.pdf")

            engine.generate_docx("MauThongBao.docx", row, docx_file)

            engine.docx_to_pdf(docx_file, pdf_file)

            Path(docx_file).unlink(missing_ok=True)

            self.log(f"Done: {ma_kh}")

        self.log("Hoàn thành!")

    finally:

        engine.close()
