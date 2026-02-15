"""GUI 友好入口 — 支持拖拽文件或双击运行（无需命令行）"""

import logging
import sys
import re
import traceback
from pathlib import Path

# Windows 控制台编码修复：PyInstaller exe 在 cmd/PowerShell 中默认使用 cp936/cp1252，
# 无法输出中文和特殊字符，导致 UnicodeEncodeError 使程序静默崩溃
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

VERSION = "1.1.0"


def _setup_file_logger(log_dir: Path) -> logging.Logger:
    """在输入文件同目录创建日志文件，记录详细处理过程。"""
    log_path = log_dir / "amazon-excel-processor.log"
    file_logger = logging.getLogger("aep")
    file_logger.setLevel(logging.DEBUG)
    # 避免重复添加 handler
    if not file_logger.handlers:
        fh = logging.FileHandler(str(log_path), mode="w", encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
        ))
        file_logger.addHandler(fh)
    return file_logger


def pause_exit(code: int = 0):
    """等待用户按回车后退出（双击运行时窗口不会立刻关闭）。"""
    print()
    input("按回车键退出...")
    sys.exit(code)


def main():
    # 如果有命令行参数，直接当文件路径用
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # 没有参数，提示用户输入
        print("=" * 50)
        print(f"  亚马逊 Excel 模板批量处理工具 v{VERSION}")
        print("=" * 50)
        print()
        print("用法：将 .xlsm / .xlsx 文件拖到本程序图标上")
        print("  或在下方粘贴文件路径：")
        print()
        input_file = input("文件路径: ").strip().strip('"').strip("'")
        input_file = re.sub(r'\\(?=[^/\\:\w])', '', input_file)
        if not input_file:
            print("未输入文件路径")
            pause_exit(1)

    input_path = Path(input_file)
    if not input_path.exists():
        print(f"ERROR: 文件不存在: {input_path}")
        pause_exit(1)

    if input_path.suffix.lower() not in (".xlsx", ".xlsm"):
        print(f"ERROR: 不支持的文件格式: {input_path.suffix}，仅支持 .xlsx 和 .xlsm")
        pause_exit(1)

    # 初始化文件日志（写到输入文件同目录）
    flog = _setup_file_logger(input_path.parent)
    flog.info("版本: %s, 平台: %s, Python: %s", VERSION, sys.platform, sys.version)
    flog.info("输入文件: %s", input_path)

    # 延迟导入，让上面的基本检查更快
    from amazon_excel_processor.excel_io import load_workbook, locate_columns, group_rows, save_workbook
    from amazon_excel_processor.name_normalizer import normalize_group
    from amazon_excel_processor.field_filler import detect_ratio_type, fill_group

    def log(msg: str):
        print(msg, flush=True)
        flog.info(msg.strip())

    try:
        log(f"\n>> 读取文件: {input_path.name} ...")
        wb, ws, template_name = load_workbook(input_path)
        flog.info("sheet='%s', max_row=%d, max_column=%d", template_name, ws.max_row, ws.max_column)
        log(">> 文件加载完成")

        col_map = locate_columns(ws)
        product_name_col = col_map["Product Name"]
        found_cols = sorted(col_map.items(), key=lambda x: x[1])
        col_info = ', '.join(f'{name}(列{idx})' for name, idx in found_cols)
        log(f">> 列定位完成: {col_info}")

        groups = group_rows(ws)
        flog.info("数据行范围: row %d - %d, 分组: %d 组", 4, 4 + len(groups) * 11 - 1 if groups else 3, len(groups))
        if not groups:
            log("[!] 没有可处理的数据")
            output_path = save_workbook(ws, input_path, template_name)
            log(f"输出文件: {output_path}")
            pause_exit(0)

        total_rows = len(groups) * 11
        log(f">> 共 {len(groups)} 个产品组, {total_rows} 行数据\n")

        for idx, rows in enumerate(groups, 1):
            ratio_type = detect_ratio_type(ws, rows, product_name_col)
            log(f"  [{idx}/{len(groups)}] 行{rows[0]}-{rows[-1]} 比例: {ratio_type}")
            normalize_group(ws, rows, product_name_col, ratio_type)
            fill_group(ws, rows, col_map, ratio_type)

        log("\n>> 保存文件...")
        output_path = save_workbook(ws, input_path, template_name)
        flog.info("输出文件: %s", output_path)

        log("")
        log("=" * 50)
        log("  [OK] 处理完成")
        log("=" * 50)
        log(f"  产品组数: {len(groups)}")
        log(f"  总行数:   {total_rows}")
        log(f"  输出文件: {output_path}")
        log("=" * 50)

        pause_exit(0)

    except Exception as e:
        flog.exception("处理失败")
        print(f"\n[ERROR] 处理失败: {e}")
        traceback.print_exc()
        print(f"\n详细日志已保存到: {input_path.parent / 'amazon-excel-processor.log'}")
        pause_exit(1)


if __name__ == "__main__":
    main()
