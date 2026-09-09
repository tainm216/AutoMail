import tkinter as tk
from pathlib import Path
import os
import re

CONFIG_DIR = Path("config")

SUBJECT_FILE = (CONFIG_DIR / "subject.html")

BODY_FILE = (CONFIG_DIR / "body.html")

FILENAME_FILE = (CONFIG_DIR / "filename.html")


def load_subject():

    if not SUBJECT_FILE.exists():
        return ""

    return SUBJECT_FILE.read_text(encoding="utf-8")


def save_subject(content):

    SUBJECT_FILE.write_text(content, encoding="utf-8")


def load_body():

    if not BODY_FILE.exists():
        return ""

    return BODY_FILE.read_text(encoding="utf-8")


def save_body(content):

    BODY_FILE.write_text(content, encoding="utf-8")


def load_filename():

    if not FILENAME_FILE.exists():
        return "{{MA_KHANG}}_{{TIMESTAMP}}"

    return FILENAME_FILE.read_text(encoding="utf-8")


def save_filename(content):

    FILENAME_FILE.write_text(content, encoding="utf-8")


def save_templates(subject, body, filename):

    save_subject(subject)

    save_body(body)

    save_filename(filename)


def open_output_folder():

    output_dir = Path("output")

    output_dir.mkdir(exist_ok=True)

    os.startfile(output_dir.resolve())


def insert_text(text_widget, value):

    text_widget.insert("insert", value)


def insert_field(text_widget, field_name):

    placeholder = (f"{{{{{field_name}}}}}")

    insert_text(text_widget, placeholder)


def insert_html_tag(text_widget, tag):

    tag_map = {"B": "b", "I": "i", "U": "u"}

    try:

        selected_text = text_widget.get("sel.first", "sel.last")

        html_tag = tag_map[tag]

        replacement = (f"<{html_tag}>"
                       f"{selected_text}"
                       f"</{html_tag}>")

        text_widget.delete("sel.first", "sel.last")

        text_widget.insert("insert", replacement)

    except tk.TclError:

        if tag == "BR":

            text_widget.insert("insert", "<br>")

        elif tag == "P":

            text_widget.insert("insert", "<p></p>")

        else:

            html_tag = tag_map[tag]

            text_widget.insert("insert", f"<{html_tag}></{html_tag}>")


def replace_fields(text, row):

    for key, value in row.items():

        text = text.replace(f"{{{{{key}}}}}", str(value))

    return text
