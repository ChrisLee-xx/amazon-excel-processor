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
    if result.returncode != 0:
        print(f"\nBuild FAILED (exit code: {result.returncode})")
        sys.exit(1)

    ext = ".exe" if platform.system() == "Windows" else ""
    binary_path = f"dist/{APP_NAME}{ext}"
    print(f"\nBuild OK! Output: {binary_path}")

    # Mac: 生成 .command 启动脚本, 双击即可在终端中运行二进制
    # (无扩展名的二进制直接双击会被 TextEdit 当作文本打开并报编码错误)
    if platform.system() == "Darwin":
        command_path = f"dist/run-{APP_NAME}.command"
        script = (
            "#!/bin/bash\n"
            "cd \"$(dirname \"$0\")\"\n"
            f"./{APP_NAME} \"$@\"\n"
        )
        with open(command_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(script)
        os.chmod(command_path, 0o755)
        print(f"Mac launcher: {command_path} (双击此文件运行)")


if __name__ == "__main__":
    build()
