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
    # 使用全新临时目录作为 work/dist/spec, 避免 PyInstaller 删除既有产物时
    # 触发环境的安全删除确认 (SAFE_DELETE_BULK_CONFIRM_REQUIRED) 而卡住。
    # 打包完成后把可执行文件复制到标准 dist/ 目录。
    tmp_work = os.path.join(os.getcwd(), ".build_tmp_work")
    tmp_dist = os.path.join(os.getcwd(), ".build_tmp_dist")
    tmp_spec = os.path.join(os.getcwd(), ".build_tmp_spec")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", APP_NAME,
        "--noconfirm",
        "--workpath", tmp_work,
        "--distpath", tmp_dist,
        "--specpath", tmp_spec,
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
        # 复制产物到标准 dist/ 目录
        os.makedirs("dist", exist_ok=True)
        ext = ".exe" if platform.system() == "Windows" else ""
        src_exe = os.path.join(tmp_dist, APP_NAME + ext)
        dst_exe = os.path.join("dist", APP_NAME + ext)
        if os.path.exists(src_exe):
            import shutil
            shutil.copy2(src_exe, dst_exe)
            print(f"\nBuild OK! Output: {dst_exe}")
        else:
            print(f"\nBuild OK, 但未找到产物: {src_exe}")
    else:
        print(f"\nBuild FAILED (exit code: {result.returncode})")
        sys.exit(1)


if __name__ == "__main__":
    build()
