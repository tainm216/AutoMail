import tkinter as tk
from ui_helpers import (load_subject, load_body, load_filename, save_templates,
                        insert_html_tag, replace_fields, open_output_folder)
from ui_logger import GuiLogger
from ui_config import ConfigWindow
from tkinter import (ttk, filedialog, messagebox)
from ui_preview import PreviewWindow
import re
import sys
from pathlib import Path
from send_job import SendJob
import threading
import json
from resource_path import resource_path
from word_template import WordTemplate, get_output_folder
from datetime import datetime

sys.path.append(str(Path(__file__).resolve().parent.parent))
from excel_reader import load_excel


class MainWindow:

    def __init__(self):

        self.rows = []
        self.fields = []

        self.current_widget = None

        self.root = tk.Tk()

        try:
            icon_path = resource_path("icon.ico")

            self.root.iconbitmap(icon_path)

        except Exception as ex:
            print(f"Không thể tải icon cửa sổ: {ex}")

        self.root.title("Phần mềm tự động gửi email thông báo")

        self.root.geometry("1400x1100")

        self.build_ui()

    def open_config(self):

        ConfigWindow(self.root)

    def load_config(self):

        with open(resource_path("config/config.json"), "r",
                  encoding="utf-8") as f:

            return json.load(f)

    def check_data(self):

        if not self.rows:

            self.logger.warning("Chưa chọn file Excel")

            return

        self.logger.info("Bắt đầu kiểm tra dữ liệu...")

        total = len(self.rows)

        missing_email = 0
        invalid_email = 0

        for row in self.rows:

            email = str(row.get("EMAIL", "")).strip()

            if not email:

                missing_email += 1

                continue

            if "@" not in email:

                invalid_email += 1

        # ==========================
        # Kiểm tra placeholder
        # ==========================

        pdf_filename = self.pdf_filename_entry.get("1.0", "end")

        subject = self.subject_entry.get("1.0", "end")

        body = self.body_text.get("1.0", "end")

        template_text = (pdf_filename + "\n" + subject + "\n" + body)

        placeholders = set(re.findall(r"\{\{(.*?)\}\}", template_text))

        excel_fields = set(self.fields)

        # TIMESTAMP là field hệ thống,
        # không cần tồn tại trong Excel.

        excel_fields.add("TIMESTAMP")

        invalid_fields = []

        for field in placeholders:

            if field not in excel_fields:

                invalid_fields.append(field)

        # ==========================
        # Log
        # ==========================

        self.logger.success(f"Tổng KH: {total}")

        self.logger.success(f"Có email: "
                            f"{total - missing_email}")

        if missing_email:

            self.logger.warning(f"Thiếu email: "
                                f"{missing_email}")

        if invalid_email:

            self.logger.warning(f"Email sai định dạng: "
                                f"{invalid_email}")

        if invalid_fields:

            self.logger.warning("Key không tồn tại:")

            for field in invalid_fields:

                self.logger.warning(f"  - {field}")

        else:

            self.logger.success("Tất cả key hợp lệ")

        self.summary_label.config(text=(f"Tổng KH: {total} | "
                                        f"Đã gửi: 0 | "
                                        f"Lỗi: 0 | "
                                        f"Thiếu email: {missing_email}"))

        self.logger.success("Kiểm tra hoàn tất")

    def run_export_pdf_thread(self):

        thread = threading.Thread(target=self.export_pdf_worker, daemon=True)

        thread.start()

    def export_pdf_worker(self):

        self.root.after(0, lambda: self.progress.configure(value=0))

        self.root.after(
            0, lambda: self.btn_export_pdf.config(state="disabled",
                                                  text="Đang xuất PDF..."))

        if not self.rows:

            self.logger.warning("Chưa chọn file Excel")

            return

        engine = WordTemplate(logger=self.logger)

        try:

            output_folder = get_output_folder()

            total = len(self.rows)

            self.logger.info(f"Bắt đầu xuất PDF ({total} khách hàng)")

            for index, row in enumerate(self.rows, start=1):

                ma_kh = str(row.get("MA_KHANG", row.get("MA_KHTT",
                                                        "UNKNOWN"))).strip()

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

                row_data = dict(row)

                row_data["TIMESTAMP"] = timestamp

                pdf_name = replace_fields(
                    self.pdf_filename_entry.get("1.0", "end").strip(),
                    row_data)

                if not pdf_name.lower().endswith(".pdf"):

                    pdf_name += ".pdf"

                pdf_name = re.sub(r'[<>:"/\\|?*]', "_", pdf_name)

                pdf_file = output_folder / pdf_name

                if pdf_file.exists():

                    stem = pdf_file.stem
                    suffix = pdf_file.suffix

                    counter = 1

                    while True:

                        candidate = (output_folder /
                                     f"{stem}({counter}){suffix}")

                        if not candidate.exists():

                            pdf_file = candidate
                            break

                        counter += 1

                docx_file = (output_folder / f"{ma_kh}_{timestamp}.docx")

                try:

                    self.logger.info(f"[{index}/{total}] "
                                     f"Đang tạo PDF {ma_kh}")

                    engine.generate_docx(self.template_var.get(), row_data,
                                         docx_file)

                    engine.docx_to_pdf(docx_file, pdf_file)

                    Path(docx_file).unlink(missing_ok=True)

                    percent = int(index * 100 / total)

                    self.root.after(
                        0, lambda p=percent: self.progress.configure(value=p))

                except Exception as ex:

                    self.logger.error(f"Lỗi {ma_kh}: {ex}")

            self.root.after(
                0, lambda: self.btn_export_pdf.config(state="normal",
                                                      text="Xuất PDF offline"))

            self.root.after(0, lambda: self.progress.configure(value=100))

            self.logger.success("Xuất PDF hoàn tất")

        finally:

            engine.close()

    def set_current_widget(self, widget):

        self.current_widget = widget

    def save_email_template(self):

        filename = self.pdf_filename_entry.get("1.0", "end").strip()

        subject = self.subject_entry.get("1.0", "end").strip()

        body = self.body_text.get("1.0", "end").strip()

        save_templates(subject, body, filename)

        self.logger.success("Đã lưu mẫu")

    def load_excel_file(self):

        file_path = filedialog.askopenfilename(title="Chọn file Excel dữ liệu",
                                               filetypes=[("Excel Files",
                                                           "*.xlsx")])

        if not file_path:
            return

        try:

            rows = load_excel(file_path)

        except Exception as ex:

            self.file_var.set("")

            self.rows = []
            self.fields = []

            self.load_field_list()

            self.logger.error(f"Lỗi file Excel: {ex}")

            messagebox.showerror("Lỗi file Excel", str(ex))

            return

        if not rows:

            self.file_var.set("")

            self.rows = []
            self.fields = []

            self.load_field_list()

            self.logger.error("File Excel không có dữ liệu")

            messagebox.showerror("Lỗi file Excel",
                                 "File Excel không có dữ liệu.")

            return

        # ==========================================
        # CHỈ GÁN FILE SAU KHI ĐỌC THÀNH CÔNG
        # ==========================================

        self.file_var.set(file_path)

        self.rows = rows

        self.fields = list(self.rows[0].keys())

        self.load_field_list()

        self.logger.success(f"Đã tải {len(self.fields)} trường dữ liệu")

    def load_template_file(self):

        file_path = filedialog.askopenfilename(filetypes=[("Word Document",
                                                           "*.docx")])

        if not file_path:
            return

        self.template_var.set(file_path)

        self.logger.success("Đã chọn file Word")

    def load_field_list(self):

        self.field_listbox.delete(0, tk.END)

        for field in self.fields:

            self.field_listbox.insert(tk.END, f"{{{{{field}}}}}")

        self.field_listbox.insert(tk.END, "{{TIMESTAMP}}")

    def filter_fields(self, *args):

        keyword = (self.search_var.get().strip().upper())

        self.field_listbox.delete(0, tk.END)

        for field in self.fields:

            if keyword in field.upper():

                self.field_listbox.insert(tk.END, f"{{{{{field}}}}}")

    def insert_selected_field(self):

        if not self.current_widget:
            return

        selected = (self.field_listbox.curselection())

        if not selected:
            return

        value = self.field_listbox.get(selected[0])

        self.current_widget.insert("insert", value)

    def insert_tag(self, tag):

        if not self.current_widget:
            return

        insert_html_tag(self.current_widget, tag)

    def preview_email(self):

        if not self.rows:

            self.logger.warning("Chưa chọn file Excel")

            return

        subject_template = self.subject_entry.get("1.0", "end").strip()

        body_template = self.body_text.get("1.0", "end").strip()

        pdf_filename_template = self.pdf_filename_entry.get("1.0",
                                                            "end").strip()

        PreviewWindow(self.root, self.rows, subject_template, body_template,
                      pdf_filename_template)

    def open_output(self):

        open_output_folder()

        self.logger.success("Đã mở thư mục output")

    def update_progress(self, percent):

        self.progress["value"] = percent

        self.root.update_idletasks()

    def send_email(self):

        if not self.rows:

            self.logger.warning("Chưa chọn file Excel")

            return

        confirm = messagebox.askyesno(
            "Xác nhận", f"Bạn có chắc muốn gửi "
            f"{len(self.rows)} email?")

        if not confirm:

            return

        subject_template = (self.subject_entry.get("1.0", "end").strip())

        body_template = (self.body_text.get("1.0", "end").strip())

        self.progress["value"] = 0

        thread = threading.Thread(target=self.run_send_job,
                                  args=(subject_template, body_template),
                                  daemon=True)

        thread.start()

    def run_send_job(self, subject_template, body_template):

        config = self.load_config()

        sender_email = config["sender_email"]

        app_password = config["app_password"]

        pdf_filename_template = self.pdf_filename_entry.get("1.0",
                                                            "end").strip()

        job = SendJob(rows=self.rows,
                      subject_template=subject_template,
                      body_template=body_template,
                      pdf_filename_template=pdf_filename_template,
                      sender_email=sender_email,
                      app_password=app_password,
                      template_file=self.template_var.get(),
                      logger=self.logger,
                      progress_callback=self.update_progress)

        job.run()

        self.summary_label.config(text=(f"Tổng KH: {len(self.rows)} | "
                                        f"Đã gửi: {job.sent_count} | "
                                        f"Lỗi: {job.error_count}"))

    def build_ui(self):

        # ==================================================
        # FILE
        # ==================================================

        file_frame = ttk.LabelFrame(self.root, text="File excel dữ liệu")

        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_var = tk.StringVar()

        ttk.Entry(file_frame, textvariable=self.file_var).pack(side="left",
                                                               fill="x",
                                                               expand=True,
                                                               padx=5,
                                                               pady=5)

        # ==================================================
        # WORD TEMPLATE
        # ==================================================

        template_frame = ttk.LabelFrame(self.root, text="File mẫu Word")

        template_frame.pack(fill="x", padx=10, pady=5)

        self.template_var = tk.StringVar()

        ttk.Entry(template_frame,
                  textvariable=self.template_var).pack(side="left",
                                                       fill="x",
                                                       expand=True,
                                                       padx=5,
                                                       pady=5)

        ttk.Button(template_frame,
                   text="Chọn file",
                   command=self.load_template_file).pack(side="right",
                                                         padx=5,
                                                         pady=5)

        ttk.Button(file_frame, text="Chọn file",
                   command=self.load_excel_file).pack(side="right",
                                                      padx=5,
                                                      pady=5)

        # ==================================================
        # MAIN
        # ==================================================

        main_frame = ttk.Frame(self.root)

        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # LEFT

        left_frame = ttk.Frame(main_frame)

        left_frame.pack(side="left", fill="both", expand=True)

        # PDF FILENAME

        ttk.Label(left_frame, text="Tên file PDF").pack(anchor="w")

        self.pdf_filename_entry = tk.Text(left_frame, height=2)

        self.pdf_filename_entry.bind(
            "<FocusIn>",
            lambda e: self.set_current_widget(self.pdf_filename_entry))

        self.pdf_filename_entry.pack(fill="x", pady=5)

        # SUBJECT

        ttk.Label(left_frame, text="Subject (Tiêu đề email)").pack(anchor="w")

        self.subject_entry = tk.Text(left_frame, height=3)

        self.subject_entry.bind(
            "<FocusIn>", lambda e: self.set_current_widget(self.subject_entry))

        self.subject_entry.pack(fill="x", pady=5)

        # BODY

        ttk.Label(left_frame, text="Body (Nội dung email)").pack(anchor="w")

        self.body_text = tk.Text(left_frame)

        self.body_text.bind("<FocusIn>",
                            lambda e: self.set_current_widget(self.body_text))

        self.body_text.pack(fill="both", expand=True)

        # Load templates

        self.pdf_filename_entry.insert("1.0", load_filename())

        self.subject_entry.insert("1.0", load_subject())

        self.body_text.insert("1.0", load_body())

        # RIGHT

        right_frame = ttk.Frame(main_frame, width=350)

        right_frame.pack(side="right", fill="y", padx=10)

        ttk.Label(right_frame, text="Tìm kiếm trường").pack(anchor="w")

        self.search_var = tk.StringVar()

        self.search_var.trace_add("write", self.filter_fields)

        ttk.Entry(right_frame, textvariable=self.search_var).pack(fill="x",
                                                                  pady=5)

        ttk.Label(right_frame, text="Trường dữ liệu").pack(anchor="w")

        list_frame = ttk.Frame(right_frame)

        list_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame)

        scrollbar.pack(side="right", fill="y")

        self.field_listbox = tk.Listbox(list_frame,
                                        yscrollcommand=scrollbar.set)

        self.field_listbox.pack(side="left", fill="both", expand=True)

        self.field_listbox.bind("<Double-Button-1>",
                                lambda e: self.insert_selected_field())

        scrollbar.config(command=self.field_listbox.yview)

        ttk.Button(right_frame,
                   text="Chèn trường",
                   command=self.insert_selected_field).pack(pady=5)

        # ==================================================
        # TOOLBAR
        # ==================================================

        toolbar_frame = ttk.Frame(self.root)

        toolbar_frame.pack(fill="x", padx=10, pady=5)

        for tag in ["B", "I", "U", "BR", "P"]:

            ttk.Button(toolbar_frame,
                       text=tag,
                       command=lambda t=tag: self.insert_tag(t)).pack(
                           side="left", padx=2)

        # ==================================================
        # ACTION
        # ==================================================

        action_frame = ttk.Frame(self.root)

        action_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(action_frame,
                   text="Lưu mẫu email",
                   command=self.save_email_template).pack(side="left", padx=5)

        ttk.Button(action_frame, text="Xem trước",
                   command=self.preview_email).pack(side="left", padx=5)

        ttk.Button(action_frame,
                   text="Kiểm tra dữ liệu",
                   command=self.check_data).pack(side="left", padx=5)

        self.btn_export_pdf = ttk.Button(action_frame,
                                         text="Xuất PDF offline",
                                         command=self.run_export_pdf_thread)

        self.btn_export_pdf.pack(side="left", padx=5)

        ttk.Button(action_frame,
                   text="Cấu hình email",
                   command=self.open_config).pack(side="left", padx=5)

        ttk.Button(action_frame, text="Gửi email",
                   command=self.send_email).pack(side="left", padx=5)

        ttk.Button(action_frame, text="Mở Output",
                   command=self.open_output).pack(side="left", padx=5)

        # ==================================================
        # SUMMARY
        # ==================================================

        self.summary_label = ttk.Label(self.root,
                                       text=("Tổng KH: 0 | "
                                             "Đã gửi: 0 | "
                                             "Lỗi: 0 | "
                                             "Thiếu email: 0"))

        self.summary_label.pack(fill="x", padx=10, pady=5)

        # ==================================================
        # PROGRESS
        # ==================================================

        self.progress = ttk.Progressbar(self.root, mode="determinate")

        self.progress.pack(fill="x", padx=10, pady=5)

        # ==================================================
        # LOG
        # ==================================================

        log_frame = ttk.LabelFrame(self.root, text="Log hoạt động")

        log_frame.pack(fill="both", expand=False, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, height=12)

        self.log_text.pack(fill="both", expand=True)

        self.logger = GuiLogger(self.log_text)

    def run(self):

        self.root.mainloop()
