import tkinter as tk
import json

from tkhtmlview import HTMLLabel
from datetime import datetime

from ui_helpers import replace_fields
from resource_path import resource_path


class PreviewWindow:

    def __init__(self, parent, rows, subject_template, body_template,
                 pdf_filename_template):

        self.rows = rows
        self.index = 0

        self.subject_template = subject_template

        self.body_template = body_template

        self.pdf_filename_template = pdf_filename_template

        self.preview_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

        self.window = tk.Toplevel(parent)

        self.window.title("Xem trước email")

        self.window.geometry("900x700")

        self.build_ui()

        self.render_customer()

    def load_sender_email(self):

        try:

            with open(resource_path("config/config.json"),
                      "r",
                      encoding="utf-8") as f:

                config = json.load(f)

            return config.get("sender_email", "")

        except Exception:

            return ""

    def build_ui(self):

        nav_frame = tk.Frame(self.window)

        nav_frame.pack(fill="x", padx=10, pady=5)

        tk.Button(nav_frame, text="◀ Trước",
                  command=self.prev_customer).pack(side="left")

        tk.Button(nav_frame, text="Sau ▶",
                  command=self.next_customer).pack(side="right")

        self.info_label = tk.Label(nav_frame, text="")

        self.info_label.pack()

        # ==================================================
        # EMAIL HEADER
        # ==================================================

        header_frame = tk.Frame(self.window, padx=10, pady=5)

        header_frame.pack(fill="x", padx=10)

        self.from_label = tk.Label(header_frame,
                                   text="",
                                   anchor="w",
                                   justify="left",
                                   font=("Segoe UI", 10))

        self.from_label.pack(anchor="w")

        self.to_label = tk.Label(header_frame,
                                 text="",
                                 anchor="w",
                                 justify="left",
                                 font=("Segoe UI", 10))

        self.to_label.pack(anchor="w")

        tk.Frame(header_frame, height=1, bg="#d0d0d0").pack(fill="x", pady=6)

        subject_frame = tk.Frame(self.window)

        subject_frame.pack(fill="x", padx=10, pady=10)

        self.subject_label = tk.Label(subject_frame,
                                      text="",
                                      wraplength=850,
                                      justify="left",
                                      font=("Segoe UI", 16, "bold"))

        self.subject_label.pack(anchor="w")

        tk.Frame(subject_frame, height=1, bg="#d0d0d0").pack(fill="x", pady=5)

        body_frame = tk.Frame(self.window)

        body_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.html_view = HTMLLabel(body_frame, html="")

        self.html_view.pack(fill="both", expand=True)

        # ==================================================
        # FILE PDF
        # ==================================================

        pdf_frame = tk.Frame(self.window,
                             bd=1,
                             relief="solid",
                             padx=10,
                             pady=8)

        pdf_frame.pack(fill="x", padx=10, pady=(0, 10))

        tk.Label(pdf_frame, text="📎", font=("Segoe UI", 18)).pack(side="left",
                                                                  padx=(0, 10))

        pdf_info_frame = tk.Frame(pdf_frame)

        pdf_info_frame.pack(side="left", fill="x", expand=True)

        self.pdf_filename_label = tk.Label(pdf_info_frame,
                                           text="",
                                           anchor="w",
                                           justify="left",
                                           font=("Segoe UI", 10, "bold"))

        self.pdf_filename_label.pack(anchor="w")

        self.pdf_type_label = tk.Label(pdf_info_frame,
                                       text="PDF",
                                       anchor="w",
                                       font=("Segoe UI", 9))

        self.pdf_type_label.pack(anchor="w")

    def render_customer(self):

        row = self.rows[self.index]

        sender_email = self.load_sender_email()

        receiver_email = str(row.get("EMAIL", "")).strip()

        self.from_label.config(text=f"Từ: {sender_email}")

        self.to_label.config(text=f"Đến: {receiver_email}")

        subject = replace_fields(self.subject_template, row)

        body = replace_fields(self.body_template, row)

        # TIMESTAMP dùng cho preview
        preview_row = dict(row)

        preview_row["TIMESTAMP"] = self.preview_timestamp

        pdf_filename = replace_fields(self.pdf_filename_template, preview_row)

        body = body.replace("\n", "<br>")

        self.subject_label.config(text=subject)

        self.pdf_filename_label.config(text=pdf_filename)

        self.html_view.set_html(body)

        customer_name = row.get("TEN_KH", "")

        customer_code = row.get("MA_KHANG", "")

        self.info_label.config(text=(f"STT {self.index + 1}/{len(self.rows)}"
                                     f" | {customer_code}"
                                     f" | {customer_name}"))

    def next_customer(self):

        if self.index < len(self.rows) - 1:

            self.index += 1

            self.render_customer()

    def prev_customer(self):

        if self.index > 0:

            self.index -= 1

            self.render_customer()
