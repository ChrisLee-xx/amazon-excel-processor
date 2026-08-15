"""Cross-platform build script using PyInstaller"""

import os
import platform
import subprocess
import sys

APP_NAME = "amazon-excel-processor"

# Fix Windows CI encoding (cp1252 can't handle CJK/emoji)
if sys.stdout.encoding and sys.stdout.encoding.lower().startswith("cp"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def build():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", APP_NAME,
        "--clean",
        "--noconfirm",
        "--hidden-import", "openpyxl",
        "--hidden-import", "openpyxl.cell",
        "--hidden-import", "openpyxl.worksheet",
        "--hidden-import", "openpyxl.reader",
        "--hidden-import", "openpyxl.writer",
        "--hidden-import", "openpyxl.packaging",
        "--hidden-import", "openpyxl.utils",
        "--hidden-import", "openpyxl.styles",
        "--hidden-import", "openpyxl.xml",
        "--hidden-import", "openpyxl.xml.functions",
        "--hidden-import", "et_xmlfile",
        "--paths", "src",
        "src/amazon_excel_processor/gui_entry.py",
    ]

    # 两个平台都需要控制台窗口（用于显示处理进度和等待用户输入）
    cmd.append("--console")

    print(f"Building for {platform.system()}...")
    print(f"Command: {' '.join(cmd)}\n")

    result = subprocess.run(cmd)
    if result.returncode == 0:
        ext = ".exe" if platform.system() == "Windows" else ""
        print(f"\nBuild OK! Output: dist/{APP_NAME}{ext}")
    else:
        print(f"\nBuild FAILED (exit code: {result.returncode})")
        sys.exit(1)


if __name__ == "__main__":
    build()
