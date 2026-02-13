"""GUI 友好入口 — 支持拖拽文件或双击运行（无需命令行）"""

import sys
import os
import re
import traceback
from pathlib import Path


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
        print("  亚马逊 Excel 模板批量处理工具")
        print("=" * 50)
        print()
        print("用法：将 .xlsm / .xlsx 文件拖到本程序图标上")
        print("  或在下方粘贴文件路径：")
        print()
        input_file = input("文件路径: ").strip().strip('"').strip("'")
        # 清理 shell 转义符：macOS zsh 粘贴路径时会把特殊字符转义
        # 如 file\[1\].xlsm → 实际文件名是 file[1].xlsm
        # 只移除"反斜杠+非路径字符"的组合，保留 Windows 路径分隔符 \
        # \[ \] \( \) \  \! \# \$ \& \' \~ \{ \} 等都是 shell 转义
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

    # 延迟导入，让上面的基本检查更快
    from amazon_excel_processor.excel_io import load_workbook, locate_columns, group_rows, save_workbook
    from amazon_excel_processor.name_normalizer import normalize_group
    from amazon_excel_processor.field_filler import detect_ratio_type, fill_group

    def log(msg: str):
        print(msg, flush=True)

    try:
        log(f"\n📂 读取文件: {input_path.name} ...")
        wb, ws, template_name = load_workbook(input_path)
        log("✅ 文件加载完成")

        col_map = locate_columns(ws)
        product_name_col = col_map["Product Name"]
        log(f"✅ 列定位完成")

        groups = group_rows(ws)
        if not groups:
            log("⚠️ 没有可处理的数据")
            output_path = save_workbook(ws, input_path, template_name)
            log(f"输出文件: {output_path}")
            pause_exit(0)

        total_rows = len(groups) * 11
        log(f"📊 共 {len(groups)} 个产品组, {total_rows} 行数据\n")

        for idx, rows in enumerate(groups, 1):
            ratio_type = detect_ratio_type(ws, rows, product_name_col)
            log(f"  [{idx}/{len(groups)}] 比例: {ratio_type}")
            normalize_group(ws, rows, product_name_col, ratio_type)
            fill_group(ws, rows, col_map, ratio_type)

        log("\n💾 保存文件...")
        output_path = save_workbook(ws, input_path, template_name)

        log("")
        log("=" * 50)
        log("  ✅ 处理完成")
        log("=" * 50)
        log(f"  产品组数: {len(groups)}")
        log(f"  总行数:   {total_rows}")
        log(f"  输出文件: {output_path}")
        log("=" * 50)

        pause_exit(0)

    except Exception as e:
        print(f"\n❌ 处理失败: {e}")
        traceback.print_exc()
        pause_exit(1)


if __name__ == "__main__":
    main()
