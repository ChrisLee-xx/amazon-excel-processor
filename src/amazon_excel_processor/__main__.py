"""CLI 入口 — 串联 excel_io → name_normalizer → field_filler → save"""

import argparse
import logging
import sys
from pathlib import Path

from .excel_io import load_workbook, locate_columns, group_rows, save_workbook
from .name_normalizer import normalize_group
from .field_filler import detect_ratio_type, fill_group

logger = logging.getLogger("amazon_excel_processor")


def main():
    parser = argparse.ArgumentParser(
        description="亚马逊上架商品 Excel 模板批量规范化处理工具"
    )
    parser.add_argument("input_file", help="输入 Excel 文件路径 (.xlsx 或 .xlsm)")
    parser.add_argument("-o", "--output", help="输出文件路径（默认: {input}_processed.{ext}）")
    parser.add_argument("-v", "--verbose", action="store_true", help="显示详细日志")
    args = parser.parse_args()

    log_level = logging.DEBUG if args.verbose else logging.INFO
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    handler.setLevel(log_level)
    logging.root.addHandler(handler)
    logging.root.setLevel(log_level)

    def log_print(msg: str):
        print(msg, flush=True)

    input_path = Path(args.input_file)
    if not input_path.exists():
        log_print(f"ERROR: 文件不存在: {input_path}")
        sys.exit(1)

    try:
        log_print(f">> 读取文件: {input_path} ...")
        wb, ws, template_name = load_workbook(input_path)
        log_print(">> 文件加载完成")

        col_map = locate_columns(ws)
        product_name_col = col_map["Product Name"]
        log_print(f">> 列定位完成: {', '.join(col_map.keys())}")

        groups = group_rows(ws)
        if not groups:
            log_print("[!] 没有可处理的数据")
            output_path = save_workbook(ws, input_path, template_name, args.output)
            log_print(f"输出文件: {output_path}")
            return

        total_rows = len(groups) * 11
        log_print(f">> 共 {len(groups)} 个产品组, {total_rows} 行数据")
        log_print("")

        for idx, rows in enumerate(groups, 1):
            ratio_type = detect_ratio_type(ws, rows, product_name_col)
            log_print(f"  [{idx}/{len(groups)}] 比例: {ratio_type}")

            normalize_group(ws, rows, product_name_col, ratio_type)
            fill_group(ws, rows, col_map, ratio_type)

        log_print("")
        log_print(">> 保存文件...")
        output_path = save_workbook(ws, input_path, template_name, args.output)

        log_print("")
        log_print("=" * 50)
        log_print("  [OK] 处理完成")
        log_print("=" * 50)
        log_print(f"  产品组数: {len(groups)}")
        log_print(f"  总行数:   {total_rows}")
        log_print(f"  输出文件: {output_path}")
        log_print("=" * 50)

    except ValueError as e:
        logger.error("处理失败: %s", e)
        sys.exit(1)
    except Exception as e:
        logger.error("未预期错误: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()
