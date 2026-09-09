from pathlib import Path
from datetime import datetime
import json
from resource_path import resource_path
from excel_reader import load_excel
from word_template import WordTemplate
from mail_sender import send_mail
from ui_logger import MailLogger
from template_engine import replace_template

CONFIG_DIR = resource_path("config")


def get_output_folder():

    now = datetime.now()

    folder = (Path("output") / now.strftime("%Y") / now.strftime("%m") /
              now.strftime("%d"))

    folder.mkdir(parents=True, exist_ok=True)

    return folder


def load_config():

    with open(CONFIG_DIR / "config.json", "r", encoding="utf-8") as f:

        return json.load(f)


def load_subject():

    with open(CONFIG_DIR / "subject.html", "r", encoding="utf-8") as f:

        return f.read()


def load_body():

    with open(CONFIG_DIR / "body.html", "r", encoding="utf-8") as f:

        return f.read()


def main():

    config = load_config()

    subject_template = load_subject()

    body_template = load_body()

    sender_email = config["sender_email"]
    app_password = config["app_password"]

    rows = load_excel("data.xlsx")

    print(f"Tổng số khách hàng: {len(rows)}")

    output_folder = get_output_folder()

    logger = MailLogger(output_folder)

    engine = WordTemplate()

    try:

        for index, row in enumerate(rows, start=1):

            ma_kh = str(row.get("MA_KHANG", "")).strip()

            ten_kh = str(row.get("TEN_KH", "")).strip()

            email = str(row.get("EMAIL", "")).strip()

            today = datetime.now().strftime("%Y%m%d")

            pdf_file = (output_folder / f"{ma_kh}_{today}.pdf")

            docx_file = (output_folder / f"{ma_kh}_{today}.docx")

            print(f"[{index}/{len(rows)}] "
                  f"Đang xử lý {ma_kh}")

            try:

                # Tạo DOCX
                engine.generate_docx("MauThongBao.docx", row, docx_file)

                # Xuất PDF
                engine.docx_to_pdf(docx_file, pdf_file)

                # Xóa DOCX
                Path(docx_file).unlink(missing_ok=True)

                # Kiểm tra email

                if not email:

                    logger.add(ma_kh=ma_kh,
                               ten_kh=ten_kh,
                               email="",
                               pdf_file=pdf_file,
                               status="Không có email")

                    continue

                # Gửi mail

                subject = replace_template(subject_template, row)

                body = replace_template(body_template, row)

                send_mail(sender_email=sender_email,
                          app_password=app_password,
                          receiver_email=email,
                          subject=subject,
                          body=body,
                          pdf_file=str(pdf_file))

                logger.add(ma_kh=ma_kh,
                           ten_kh=ten_kh,
                           email=email,
                           pdf_file=pdf_file,
                           status="Đã gửi")

            except Exception as e:

                logger.add(ma_kh=ma_kh,
                           ten_kh=ten_kh,
                           email=email,
                           pdf_file=pdf_file,
                           status=f"Lỗi: {e}")

                print(f"Lỗi {ma_kh}: {e}")

        logger.save()

        print("\nHoàn thành!")

    finally:

        engine.close()


if __name__ == "__main__":
    main()
