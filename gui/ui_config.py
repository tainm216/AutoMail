import tkinter as tk
from tkinter import ttk, messagebox
import json
from resource_path import resource_path


class ConfigWindow:

    def __init__(self, parent):

        self.parent = parent

        self.window = tk.Toplevel(parent)

        self.window.title("Cấu hình email")

        self.window.geometry("500x250")

        self.window.resizable(False, False)

        self.build_ui()

        self.load_config()

        self.window.transient(parent)

        self.window.grab_set()

    def build_ui(self):

        main_frame = ttk.Frame(self.window)

        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # ==========================
        # EMAIL
        # ==========================

        ttk.Label(main_frame, text="Email gửi:").pack(anchor="w")

        self.sender_email_var = tk.StringVar()

        ttk.Entry(main_frame,
                  textvariable=self.sender_email_var).pack(fill="x",
                                                           pady=(5, 15))

        # ==========================
        # APP PASSWORD
        # ==========================

        ttk.Label(main_frame, text="App Password:").pack(anchor="w")

        self.app_password_var = tk.StringVar()

        ttk.Entry(main_frame, textvariable=self.app_password_var,
                  show="*").pack(fill="x", pady=5)

        # ==========================
        # BUTTON
        # ==========================

        button_frame = ttk.Frame(main_frame)

        button_frame.pack(fill="x", pady=(20, 0))

        ttk.Button(button_frame, text="Lưu cấu hình",
                   command=self.save_config).pack(side="right")

        ttk.Button(button_frame, text="Hủy",
                   command=self.window.destroy).pack(side="right", padx=5)

    def load_config(self):

        try:

            with open(resource_path("config/config.json"),
                      "r",
                      encoding="utf-8") as f:

                config = json.load(f)

            self.sender_email_var.set(config.get("sender_email", ""))

            self.app_password_var.set(config.get("app_password", ""))

        except Exception as ex:

            messagebox.showerror("Lỗi",
                                 f"Không thể đọc file cấu hình:\n{ex}",
                                 parent=self.window)

    def save_config(self):

        sender_email = (self.sender_email_var.get().strip())

        app_password = (self.app_password_var.get().strip())

        if not sender_email:

            messagebox.showwarning("Thiếu thông tin",
                                   "Vui lòng nhập email gửi.",
                                   parent=self.window)

            return

        if not app_password:

            messagebox.showwarning("Thiếu thông tin",
                                   "Vui lòng nhập App Password.",
                                   parent=self.window)

            return

        config = {"sender_email": sender_email, "app_password": app_password}

        try:

            with open(resource_path("config/config.json"),
                      "w",
                      encoding="utf-8") as f:

                json.dump(config, f, ensure_ascii=False, indent=4)

            messagebox.showinfo("Thành công",
                                "Đã lưu cấu hình email.",
                                parent=self.window)

            self.window.destroy()

        except Exception as ex:

            messagebox.showerror("Lỗi",
                                 f"Không thể lưu cấu hình:\n{ex}",
                                 parent=self.window)
