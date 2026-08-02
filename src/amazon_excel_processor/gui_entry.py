"""GUI 友好入口 — 支持拖拽文件或双击运行（无需命令行）

模式:
  1) 单文件处理 — 单个 .xlsm/.xlsx 走原 normalize + fill 流程
  2) 三文件合并 — 普文件(主) + 木框文件 + 金框文件, 输出 21 行/组 的新文件

CLI 行为:
  - 1 个参数 → 单文件模式
  - 3 个参数 → 合并模式 (顺序: 主 木 金)
  - 0 个参数 → 交互菜单
"""

import argparse
import logging
import sys
import re
import traceback
from pathlib import Path

# Windows 控制台编码修复
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

VERSION = "1.2.0"


def _setup_file_logger(log_dir: Path) -> logging.Logger:
    log_path = log_dir / "amazon-excel-processor.log"
    file_logger = logging.getLogger("aep")
    file_logger.setLevel(logging.DEBUG)
    if not file_logger.handlers:
        fh = logging.FileHandler(str(log_path), mode="w", encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"
        ))
        file_logger.addHandler(fh)
    return file_logger


def pause_exit(code: int = 0):
    print()
    input("按回车键退出...")
    sys.exit(code)


def _clean_path(s: str) -> str:
    s = s.strip().strip('"').strip("'")
    s = re.sub(r'\\(?=[^/\\:\w])', '', s)
    return s


def _prompt_path(prompt: str) -> str:
    return _clean_path(input(prompt))


def _prompt_choice(prompt: str, choices: list) -> str:
    while True:
        v = input(prompt).strip()
        if v in choices:
            return v
        print(f"  请输入 {'/'.join(choices)} 之一")


def _run_single(input_path: Path, flog: logging.Logger):
    from amazon_excel_processor.excel_io import load_workbook, locate_columns, group_rows, save_workbook
    from amazon_excel_processor.name_normalizer import normalize_group
    from amazon_excel_processor.field_filler import detect_ratio_type, fill_group

    def log(msg: str):
        print(msg, flush=True)
        flog.info(msg.strip())

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
    if not groups:
        log("[!] 没有可处理的数据")
        output_path = save_workbook(ws, input_path, template_name)
        log(f"输出文件: {output_path}")
        return

    total_rows = len(groups) * 11
    log(f">> 共 {len(groups)} 个产品组, {total_rows} 行数据\n")

    for idx, rows in enumerate(groups, 1):
        ratio_type = detect_ratio_type(ws, rows, col_map)
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


def _run_merge(main_path: Path, wood_path, gold_path, flog: logging.Logger):
    """合并流程 (主必填, 木/金可选)。

    wood_path / gold_path 可为 Path 或 None (None 表示该文件未提供)。
    """
    from amazon_excel_processor.merger import merge_files

    def log(msg: str):
        print(msg, flush=True)
        flog.info(msg.strip())

    wood_disp = wood_path if wood_path else "(未提供, 跳过)"
    gold_disp = gold_path if gold_path else "(未提供, 跳过)"
    log("")
    log("=" * 50)
    log("  合并模式 (木/金可选)")
    log("=" * 50)
    log(f"  普文件 (主): {main_path}")
    log(f"  木框文件:    {wood_disp}")
    log(f"  金框文件:    {gold_disp}")
    log("")

    # 第 1 步: 选择上架类型
    log("请选择上架类型:")
    log("  1) 新品上架      (所有 SKU 按新命名规则重新编号)")
    log("  2) 老品补充变体  (普文件原 SKU 保留, 仅金+木 SKU 按新规则编号)")
    list_type = _prompt_choice("  输入 [1/2]: ", ["1", "2"])
    mode = "new" if list_type == "1" else "old_variant"
    mode_label = "新品上架" if mode == "new" else "老品补充变体"
    log(f"  → {mode_label}")
    log("")

    # 老品补充变体必须有变体文件
    if mode == "old_variant" and not (wood_path or gold_path):
        print("ERROR: 老品补充变体模式需要至少一个木框或金框文件")
        pause_exit(1)

    # 第 2 步: 输入 SKU 前缀
    log("请输入 SKU 命名 (推荐格式: 店铺名+日期+主题, 如 HM725):")
    sku_prefix = input("  SKU 命名: ").strip()
    if not sku_prefix:
        print("ERROR: SKU 命名不能为空")
        pause_exit(1)
    log(f"  → SKU 前缀: {sku_prefix}")
    log("")

    log(">> 开始合并 ...")
    output_path = merge_files(
        main_path=main_path,
        wood_path=wood_path,
        gold_path=gold_path,
        sku_prefix=sku_prefix,
        mode=mode,
    )
    flog.info("合并输出: %s (mode=%s)", output_path, mode)

    log("")
    log("=" * 50)
    log(f"  [OK] 合并完成 ({mode_label})")
    log("=" * 50)
    log(f"  输出文件: {output_path}")
    log("=" * 50)


def main():
    parser = argparse.ArgumentParser(description=f"亚马逊 Excel 模板批量处理工具 v{VERSION}")
    parser.add_argument("files", nargs="*",
                        help="1 个=单文件; 3 个=合并 (主 木 金); 木/金可选时用 1 个普文件 + --wood/--gold")
    parser.add_argument("--mode", choices=["single", "merge"], help="强制模式 (默认按文件数自动)")
    parser.add_argument("--wood", help="木框文件路径 (合并模式, 可选)")
    parser.add_argument("--gold", help="金框文件路径 (合并模式, 可选)")
    args = parser.parse_args()
    interactive = not args.files  # 无命令行参数 = 交互式 GUI 模式

    flog = None
    try:
        if not args.files:
            print("=" * 50)
            print(f"  亚马逊 Excel 模板批量处理工具 v{VERSION}")
            print("=" * 50)
            print()
            print("  请选择模式:")
            print("    1) 单文件处理")
            print("    2) 三文件合并 (普 + 木 + 金)")
            print()
            choice = _prompt_choice("  输入 [1/2]: ", ["1", "2"])

            if choice == "1":
                print()
                print("  请将 .xlsm / .xlsx 文件拖到此处, 或粘贴路径:")
                raw = _prompt_path("  文件路径: ")
                if not raw:
                    print("未输入文件路径")
                    pause_exit(1)
                p = Path(raw)
                if not p.exists():
                    print(f"ERROR: 文件不存在: {p}")
                    pause_exit(1)
                if p.suffix.lower() not in (".xlsx", ".xlsm"):
                    print(f"ERROR: 不支持的文件格式: {p.suffix}")
                    pause_exit(1)
                flog = _setup_file_logger(p.parent)
                flog.info("版本: %s, 模式: single", VERSION)
                _run_single(p, flog)
                pause_exit(0)
            else:
                print()
                print("  请依次输入文件路径 (普文件必填, 木框/金框可留空跳过):")
                raw_main = _prompt_path("  1. 普文件 (主文件, 含 Frame+Unframe): ")
                if not raw_main:
                    print("必须输入普文件路径")
                    pause_exit(1)
                raw_wood = _prompt_path("  2. 木框文件 (Vintage Wood Grain, 留空跳过): ")
                raw_gold = _prompt_path("  3. 金框文件 (Vintage Ornate Gold, 留空跳过): ")
                p_main = Path(raw_main)
                p_wood = Path(raw_wood) if raw_wood else None
                p_gold = Path(raw_gold) if raw_gold else None
                for pp in (p_main, p_wood, p_gold):
                    if pp is not None and not pp.exists():
                        print(f"ERROR: 文件不存在: {pp}")
                        pause_exit(1)
                flog = _setup_file_logger(p_main.parent)
                flog.info("版本: %s, 模式: merge (wood=%s, gold=%s)", VERSION,
                          bool(p_wood), bool(p_gold))
                _run_merge(p_main, p_wood, p_gold, flog)
                pause_exit(0)
        else:
            # CLI 模式
            has_wood_flag = bool(args.wood)
            has_gold_flag = bool(args.gold)
            # 判定合并模式: 强制 merge / 3 位置文件 / 带 --wood 或 --gold
            is_merge = (args.mode == "merge"
                        or len(args.files) == 3
                        or has_wood_flag or has_gold_flag)
            if args.mode == "single":
                is_merge = False
            if is_merge:
                if len(args.files) == 3:
                    p_main = Path(_clean_path(args.files[0]))
                    p_wood = Path(_clean_path(args.files[1]))
                    p_gold = Path(_clean_path(args.files[2]))
                    if has_wood_flag or has_gold_flag:
                        print("提示: 已提供 3 个位置文件, 忽略 --wood/--gold")
                elif len(args.files) == 1:
                    p_main = Path(_clean_path(args.files[0]))
                    p_wood = Path(_clean_path(args.wood)) if has_wood_flag else None
                    p_gold = Path(_clean_path(args.gold)) if has_gold_flag else None
                else:
                    print("ERROR: 合并模式需要 1 个普文件 (可用 --wood/--gold 补充木/金) 或 3 个文件 (主 木 金)")
                    print("       2 个位置文件无法区分木/金, 请用 --wood / --gold 分别指定")
                    sys.exit(1)
                for pp in (p_main, p_wood, p_gold):
                    if pp is not None and not pp.exists():
                        print(f"ERROR: 文件不存在: {pp}")
                        sys.exit(1)
                flog = _setup_file_logger(p_main.parent)
                flog.info("版本: %s, 模式: merge (CLI, wood=%s, gold=%s)", VERSION,
                          bool(p_wood), bool(p_gold))
                _run_merge(p_main, p_wood, p_gold, flog)
            else:
                if len(args.files) != 1:
                    print("ERROR: 单文件模式只接受 1 个文件 (合并: 3 个文件 或 1 个普文件 + --wood/--gold)")
                    sys.exit(1)
                p = Path(_clean_path(args.files[0]))
                if not p.exists():
                    print(f"ERROR: 文件不存在: {p}")
                    sys.exit(1)
                if p.suffix.lower() not in (".xlsx", ".xlsm"):
                    print(f"ERROR: 不支持的文件格式: {p.suffix}")
                    sys.exit(1)
                flog = _setup_file_logger(p.parent)
                flog.info("版本: %s, 模式: single (CLI)", VERSION)
                _run_single(p, flog)
    except ValueError as e:
        # 业务错误 (如同名产品重复, 文件类型不符): 给用户清晰提示, traceback 只进 log
        if flog:
            flog.exception("处理失败 (ValueError)")
        print(f"\n[ERROR] {e}")
        if flog:
            print(f"\n详细日志已保存到: {flog.handlers[0].baseFilename}")
        if interactive:
            pause_exit(1)
        sys.exit(1)
    except Exception as e:
        # 未预期错误: 用户看简短信息, traceback 进 log
        if flog:
            flog.exception("处理失败 (未预期错误)")
        print(f"\n[ERROR] 处理失败: {e}")
        if flog:
            print(f"\n详细日志已保存到: {flog.handlers[0].baseFilename}")
        else:
            print("\n(无日志)")
        if interactive:
            pause_exit(1)
        sys.exit(1)


if __name__ == "__main__":
    main()
