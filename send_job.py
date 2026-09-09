from pathlib import Path
from datetime import datetime

from ui_helpers import replace_fields

from word_template import (WordTemplate, get_output_folder)

from mail_sender import send_mail

from mail_log import (start_mail_log, write_mail_log)


class SendJob:

    def __init__(self,
                 rows,
                 subject_template,
                 body_template,
                 pdf_filename_template,
                 sender_email,
                 app_password,
                 template_file,
                 logger=None,
                 progress_callback=None):

        self.rows = rows

        self.subject_template = subject_template
        self.body_template = body_template
        self.pdf_filename_template = pdf_filename_template

        self.sender_email = sender_email
        self.app_password = app_password
        self.template_file = template_file

        self.logger = logger
        self.progress_callback = progress_callback

        self.sent_count = 0
        self.error_count = 0

        self.engine = WordTemplate(logger=self.logger)

    # ==================================================
    # LOG
    # ==================================================

    def log(self, message):

        if self.logger:

            try:

                self.logger.info(message)

            except Exception:

                print(message)

        else:

            print(message)

    # ==================================================
    # PROGRESS
    # ==================================================

    def update_progress(self, current, total):

        if self.progress_callback:

            if total <= 0:

                percent = 100

            else:

                percent = int(current * 100 / total)

            self.progress_callback(percent)

    # ==================================================
    # ERROR / RESULT LOG
    # ==================================================

    def write_result_log(self, output_folder, file_name, email, status, ex=""):

        messages = {
            "EMAIL_EMPTY": "Không có địa chỉ email người nhận",
            "WORD_ERROR": "Lỗi tạo file Word",
            "PDF_ERROR": "Lỗi chuyển đổi Word sang PDF",
            "MAIL_ERROR": "Lỗi gửi email",
            "SUCCESS": "Đã gửi thành công"
        }

        note = messages.get(status, "Lỗi không xác định")

        if ex:

            note += f": {ex}"

        write_mail_log(output_folder, file_name, email, status, note)

    # ==================================================
    # SAFE FILE NAME
    # ==================================================

    def make_safe_filename(self, filename):

        invalid_chars = '<>:"/\\|?*'

        for char in invalid_chars:

            filename = filename.replace(char, "_")

        filename = filename.strip()

        if not filename:

            filename = "ThongBao"

        return filename

    # ==================================================
    # UNIQUE PDF
    # ==================================================

    def get_unique_pdf_path(self, output_folder, filename):

        filename = self.make_safe_filename(filename)

        if not filename.lower().endswith(".pdf"):

            filename += ".pdf"

        pdf_file = (output_folder / filename)

        if not pdf_file.exists():

            return pdf_file

        stem = pdf_file.stem
        suffix = pdf_file.suffix

        counter = 1

        while True:

            candidate = (output_folder / f"{stem}({counter}){suffix}")

            if not candidate.exists():

                return candidate

            counter += 1

    # ==================================================
    # RUN
    # ==================================================

    def run(self):

        total = len(self.rows)

        self.log(f"Tổng số KH: {total}")

        output_folder = None

        try:

            output_folder = (get_output_folder())

            # ==================================================
            # TẠO FILE EXCEL BÁO CÁO CHO PHIÊN GỬI NÀY
            # ==================================================

            start_mail_log(output_folder)

            self.log("Đã tạo file báo cáo gửi mail")

            # ==================================================
            # XỬ LÝ TỪNG KHÁCH HÀNG
            # ==================================================

            for index, row in enumerate(self.rows, start=1):

                ma_kh = str(row.get("MA_KHANG", row.get("MA_KHTT",
                                                        "UNKNOWN"))).strip()

                email = str(row.get("EMAIL", "")).strip()

                # Mặc định để luôn có tên file
                pdf_file = None

                try:

                    self.log(f"[{index}/{total}] "
                             f"Đang xử lý {ma_kh}")

                    # ==================================================
                    # TIMESTAMP RIÊNG CHO KHÁCH HÀNG
                    # ==================================================

                    timestamp = (datetime.now().strftime("%Y%m%d_%H%M%S_%f"))

                    # ==================================================
                    # DATA DÙNG CHO PLACEHOLDER
                    # ==================================================

                    row_data = dict(row)

                    row_data["TIMESTAMP"] = (timestamp)

                    # ==================================================
                    # TẠO TÊN PDF
                    # ==================================================

                    pdf_name = replace_fields(self.pdf_filename_template,
                                              row_data)

                    pdf_file = (self.get_unique_pdf_path(
                        output_folder, pdf_name))

                    # Đây là tên file THỰC TẾ
                    # sẽ được dùng trong báo cáo Excel

                    file_name = pdf_file.name

                    # ==================================================
                    # DOCX TẠM
                    # ==================================================

                    docx_file = (output_folder / f"{ma_kh}_{timestamp}.docx")

                    # ==================================================
                    # WORD
                    # ==================================================

                    try:

                        self.engine.generate_docx(self.template_file, row_data,
                                                  docx_file)

                    except Exception as ex:

                        self.write_result_log(output_folder, file_name, email,
                                              "WORD_ERROR", ex)

                        self.error_count += 1

                        self.log(f"Lỗi Word {ma_kh}: "
                                 f"{ex}")

                        continue

                    # ==================================================
                    # PDF
                    # ==================================================

                    try:

                        self.engine.docx_to_pdf(docx_file, pdf_file)

                    except Exception as ex:

                        self.write_result_log(output_folder, file_name, email,
                                              "PDF_ERROR", ex)

                        self.error_count += 1

                        self.log(f"Lỗi PDF {ma_kh}: "
                                 f"{ex}")

                        continue

                    finally:

                        Path(docx_file).unlink(missing_ok=True)

                    # ==================================================
                    # EMAIL EMPTY
                    # ==================================================

                    if not email:

                        self.write_result_log(output_folder, file_name, email,
                                              "EMAIL_EMPTY")

                        self.error_count += 1

                        self.log(f"Không có email: "
                                 f"{ma_kh}")

                        continue

                    # ==================================================
                    # SUBJECT
                    # ==================================================

                    subject = replace_fields(self.subject_template, row_data)

                    # ==================================================
                    # BODY
                    # ==================================================

                    body = replace_fields(self.body_template, row_data)

                    # ==================================================
                    # SEND MAIL
                    # ==================================================

                    try:

                        send_mail(sender_email=(self.sender_email),
                                  app_password=(self.app_password),
                                  receiver_email=email,
                                  subject=subject,
                                  body=body,
                                  pdf_file=str(pdf_file))

                    except Exception as ex:

                        self.write_result_log(output_folder, file_name, email,
                                              "MAIL_ERROR", ex)

                        self.error_count += 1

                        self.log(f"Lỗi gửi mail "
                                 f"{ma_kh}: "
                                 f"{ex}")

                        continue

                    # ==================================================
                    # SUCCESS
                    # ==================================================

                    self.write_result_log(output_folder, file_name, email,
                                          "SUCCESS")

                    self.sent_count += 1

                    self.log(f"Đã gửi tới email: "
                             f"{email}")

                except Exception as ex:

                    self.error_count += 1

                    # Nếu đã tạo được PDF name
                    # thì ghi đúng tên file.
                    #
                    # Nếu lỗi xảy ra trước đó
                    # thì dùng tên theo timestamp.

                    if pdf_file is not None:

                        error_file_name = (pdf_file.name)

                    else:

                        error_file_name = (
                            f"{ma_kh}_"
                            f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
                            f".pdf")

                    self.write_result_log(output_folder, error_file_name,
                                          email, "MAIL_ERROR", ex)

                    self.log(f"Lỗi {ma_kh}: "
                             f"{ex}")

                self.update_progress(index, total)

        except Exception as ex:

            self.log(f"Lỗi nghiêm trọng: {ex}")

            self.error_count += 1

        finally:

            try:

                self.engine.close()

            except Exception:

                pass

        # ==================================================
        # HOÀN TẤT
        # ==================================================

        self.log("")

        self.log("Hoàn thành!")

        self.log(f"Đã gửi: "
                 f"{self.sent_count}")

        self.log(f"Lỗi: "
                 f"{self.error_count}")
