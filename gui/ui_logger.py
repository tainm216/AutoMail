import tkinter as tk
from datetime import datetime


class GuiLogger:

    def __init__(self, text_widget):

        self.text_widget = text_widget

        self.text_widget.tag_config("INFO", foreground="black")

        self.text_widget.tag_config("SUCCESS", foreground="green")

        self.text_widget.tag_config("WARNING", foreground="orange")

        self.text_widget.tag_config("ERROR", foreground="red")

    def _write(self, level, message):

        now = datetime.now().strftime("%H:%M:%S")

        text = (f"[{now}] "
                f"[{level}] "
                f"{message}\n")

        self.text_widget.insert(tk.END, text, level)

        self.text_widget.see(tk.END)

    def info(self, message):
        self._write("INFO", message)

    def success(self, message):
        self._write("SUCCESS", message)

    def warning(self, message):
        self._write("WARNING", message)

    def error(self, message):
        self._write("ERROR", message)
