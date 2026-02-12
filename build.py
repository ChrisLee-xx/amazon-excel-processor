"""跨平台打包脚本 — 使用 PyInstaller 生成独立可执行文件"""

import platform
import subprocess
import sys

APP_NAME = "amazon-excel-processor"

def build():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", APP_NAME,
        "--clean",
        "--noconfirm",
        # 确保 openpyxl 完整打包
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
        # 源码路径
        "--paths", "src",
        # 入口
        "src/amazon_excel_processor/gui_entry.py",
    ]

    # Windows 下加 console 模式（保留命令行窗口显示进度）
    if platform.system() == "Windows":
        cmd.append("--console")

    print(f"正在打包 ({platform.system()})...")
    print(f"命令: {' '.join(cmd)}\n")

    result = subprocess.run(cmd)
    if result.returncode == 0:
        ext = ".exe" if platform.system() == "Windows" else ""
        print(f"\n✅ 打包成功！")
        print(f"   输出: dist/{APP_NAME}{ext}")
    else:
        print(f"\n❌ 打包失败 (exit code: {result.returncode})")
        sys.exit(1)


if __name__ == "__main__":
    build()
