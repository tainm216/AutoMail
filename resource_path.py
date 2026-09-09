import sys
from pathlib import Path


def resource_path(relative_path):

    # Khi chạy file .py bằng VS Code
    if not getattr(sys, "frozen", False):

        return str(Path(__file__).resolve().parent / relative_path)

    # Khi chạy file .exe
    # Lấy thư mục nằm cạnh AutoMail.exe
    return str(Path(sys.executable).resolve().parent / relative_path)
