"""两文件合并模块"""

import logging
import re
from collections import defaultdict
from pathlib import Path
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet

from .excel_io import (
    DATA_START_ROW,
    group_rows,
    load_workbook,
    locate_columns,
    save_workbook,
)
from .field_filler import fill_group_21, fill_list_price
from .name_normalizer import (
    SIZES_32,
    VARIANT_LABELS_21,
    extract_base_title,
    remove_numeric_suffix,
    replace_hyphens,
    replace_underscores,
    deduplicate_words,
    collapse_spaces,
)

logger = logging.getLogger(__name__)

MERGED_GROUP_SIZE = 21
GOLD_GROUP_SIZE = 6
MAIN_GROUP_SIZE = 11

WOOD_STYLE = "Vintage Wood Grain Frame-style"
GOLD_STYLE = "Vintage Ornate Gold Frame-style"

COL_SELLER_SKU = 2
COL_PRODUCT_NAME = 9
COL_YOUR_PRICE = 13
COL_RELATIONSHIP_TYPE = 24
COL_PACKAGE_LEVEL = 25
COL_VARIATION_THEME = 26
COL_PARENT_SKU = 27
COL_PARENTAGE = 30
COL_COLOR = 38
COL_SIZE = 41
COL_SIZE_MAP = 55
COL_LENGTH = 62
COL_WIDTH = 63
COL_WEIGHT = 69
COL_LIST_PRICE = 145


def build_sku_prefix(shop, date, theme):
    return f"{shop.strip()}{date.strip()}{theme.strip()}"


def identify_main_file(groups):
    if not groups:
        return False, False
    first = groups[0]
    if len(first) == MAIN_GROUP_SIZE:
        return True, False
    if len(first) == GOLD_GROUP_SIZE:
        return False, True
    return False, False


def _normalize_name_for_compare(name):
    if not name:
        return ""
    s = str(name)
    s = extract_base_title(s)
    s = remove_numeric_suffix(s)
    # 主文件有 "Henri Matisse - Harmony in Red" 格式, 金文件无 "-".
    # 配对前统一替换 "- " (连字符+空格) 为单空格.
    s = re.sub(r"\s*-\s*", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def _group_base_name(ws, group, name_col=COL_PRODUCT_NAME):
    parent_row = group[0]
    v = ws.cell(row=parent_row, column=name_col).value
    return _normalize_name_for_compare(str(v) if v else "")


def pair_gold_groups(ws, groups, name_col=COL_PRODUCT_NAME):
    by_name = defaultdict(list)
    for g in groups:
        name = _group_base_name(ws, g, name_col)
        by_name[name].append(g)
    pairs = []
    for name in sorted(by_name.keys()):
        gs = by_name[name]
        if len(gs) != 2:
            raise ValueError(
                f"base name '{name}' 有 {len(gs)} 个 group, 期望 2 个 (Wood + Gold)"
            )
        pairs.append((gs[0], gs[1]))
    return pairs


def _copy_row_data(src_ws, src_row, dst_ws, dst_row, max_col):
    for c in range(1, max_col + 1):
        dst_ws.cell(row=dst_row, column=c).value = src_ws.cell(row=src_row, column=c).value


def _snapshot_row(ws, row, max_col):
    """快照一行数据, 返回 {col: value} dict (避免 ws 后续修改污染)."""
    return {c: ws.cell(row=row, column=c).value for c in range(1, max_col + 1)}


def _write_row(dst_ws, dst_row, snapshot, max_col):
    """把快照写回一行."""
    for c in range(1, max_col + 1):
        dst_ws.cell(row=dst_row, column=c).value = snapshot.get(c)


def _col_letter(col_idx):
    result = ""
    n = col_idx
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _fill_meta_columns(ws, rows, col_map):
    parent_row = rows[0]
    child_rows = rows[1:]
    var_theme = "color-size"
    if "Variation Theme" in col_map:
        vt_col = col_map["Variation Theme"]
        for r in rows:
            ws.cell(row=r, column=vt_col).value = var_theme
    if "Package Level" in col_map:
        pl_col = col_map["Package Level"]
        for r in rows:
            ws.cell(row=r, column=pl_col).value = "unit"
    if "Parentage" in col_map:
        par_col = col_map["Parentage"]
        ws.cell(row=parent_row, column=par_col).value = "Parent"
        for r in child_rows:
            ws.cell(row=r, column=par_col).value = "Child"
    if "Relationship Type" in col_map:
        rt_col = col_map["Relationship Type"]
        ws.cell(row=parent_row, column=rt_col).value = None
        for r in child_rows:
            ws.cell(row=r, column=rt_col).value = "Variation"


def merge_one_painting(
    main_snapshots,
    gold_pair,
    output_start_row,
    output_ws,
    col_map,
    gold_wood_ws=None,
    gold_gold_ws=None,
    max_col=None,
    ratio_type="3:2",
):
    """合并 1 画到 21 行结构.

    Args:
        main_snapshots: 11 元素 list, 每个元素是 {col: value} dict (main group 行的快照)
                       必须在 merge_files 入口对 main_groups 全部行做一次性快照,
                       否则 main_ws 后续被覆盖会污染后续 group 的源数据.
        gold_pair: (wood_group_6行, gold_group_6行) - 来源 ws
        output_start_row: 写到 output_ws 的起始行
        output_ws: 目标 worksheet
        col_map: 目标 ws 的列映射
        max_col: 列数, 需 main/gold/output ws 中最大值
    """
    wood_group, gold_group = gold_pair
    assert len(main_snapshots) == MAIN_GROUP_SIZE
    assert len(wood_group) == GOLD_GROUP_SIZE
    assert len(gold_group) == GOLD_GROUP_SIZE

    if max_col is None:
        max_col = output_ws.max_column

    wood_snapshots = [_snapshot_row(gold_wood_ws, r, max_col) for r in wood_group]
    gold_snapshots = [_snapshot_row(gold_gold_ws, r, max_col) for r in gold_group]

    merged_rows = []

    parent_row = output_start_row
    _write_row(output_ws, parent_row, main_snapshots[0], max_col)
    merged_rows.append(parent_row)

    for i, snap in enumerate(main_snapshots[1:]):
        dst = output_start_row + 1 + i
        _write_row(output_ws, dst, snap, max_col)
        merged_rows.append(dst)

    for i, snap in enumerate(wood_snapshots[1:]):
        dst = output_start_row + 11 + i
        _write_row(output_ws, dst, snap, max_col)
        merged_rows.append(dst)

    for i, snap in enumerate(gold_snapshots[1:]):
        dst = output_start_row + 16 + i
        _write_row(output_ws, dst, snap, max_col)
        merged_rows.append(dst)

    normalize_group_21(output_ws, merged_rows, COL_PRODUCT_NAME, ratio_type)
    fill_group_21(output_ws, merged_rows, col_map, ratio_type)
    _fill_meta_columns(output_ws, merged_rows, col_map)

    return merged_rows


def normalize_group_21(ws, rows, name_col, ratio_type="3:2"):
    sizes = SIZES_32
    parent_cell = ws.cell(row=rows[0], column=name_col)
    parent_value = parent_cell.value
    if parent_value is None:
        return
    base_title = extract_base_title(str(parent_value))
    base_title = remove_numeric_suffix(base_title)
    base_title = collapse_spaces(base_title)

    for i, row in enumerate(rows):
        cell = ws.cell(row=row, column=name_col)
        value = cell.value
        if value is None:
            continue
        if i == 0:
            name = base_title
        else:
            label = VARIANT_LABELS_21[i]
            size_idx = (i - 1) % 5
            size = sizes[size_idx]
            name = f"{base_title} {label} {size}"
        name = collapse_spaces(name)
        name = remove_numeric_suffix(name)
        name = replace_hyphens(name)
        name = replace_underscores(name)
        name = deduplicate_words(name)
        name = collapse_spaces(name)
        cell.value = name


def rewrite_sku(ws, groups, prefix, sku_col=COL_SELLER_SKU):
    counter = 1
    for group in groups:
        for row in group:
            ws.cell(row=row, column=sku_col).value = f"{prefix}-{counter}"
            counter += 1


def write_parent_sku_formulas(ws, groups, parent_sku_col=COL_PARENT_SKU, seller_sku_col=COL_SELLER_SKU):
    for group in groups:
        if len(group) < 2:
            continue
        parent_row = group[0]
        first_child = group[1]
        ws.cell(row=parent_row, column=parent_sku_col).value = None
        seller_letter = _col_letter(seller_sku_col)
        ws.cell(row=first_child, column=parent_sku_col).value = f"={seller_letter}{parent_row}"
        parent_sku_letter = _col_letter(parent_sku_col)
        for i in range(2, len(group)):
            prev_row = group[i - 1]
            ws.cell(row=group[i], column=parent_sku_col).value = f"={parent_sku_letter}{prev_row}"


def fill_list_price_synced(ws, rows, col_map):
    fill_list_price(ws, rows, col_map)


def merge_files(
    main_path,
    gold_path,
    shop,
    date,
    theme="",
    output_path=None,
):
    main_path = Path(main_path)
    gold_path = Path(gold_path)
    prefix = build_sku_prefix(shop, date, theme)
    logger.info("合并开始: main=%s, gold=%s, prefix=%s",
                main_path.name, gold_path.name, prefix)

    main_wb, main_ws, main_sheet = load_workbook(main_path)
    gold_wb, gold_ws, _ = load_workbook(gold_path)

    main_groups = group_rows(main_ws, group_size=MAIN_GROUP_SIZE)
    gold_groups = group_rows(gold_ws, group_size=GOLD_GROUP_SIZE)

    is_main, _ = identify_main_file(main_groups)
    _, is_gold = identify_main_file(gold_groups)
    if not (is_main and is_gold):
        raise ValueError(
            f"文件类型识别失败: main={len(main_groups[0]) if main_groups else 0}行/组, "
            f"gold={len(gold_groups[0]) if gold_groups else 0}行/组. "
            f"期望 main=11, gold=6."
        )

    gold_pairs = pair_gold_groups(gold_ws, gold_groups)

    main_by_name = {}
    for g in main_groups:
        main_by_name[_group_base_name(main_ws, g)] = g
    gold_by_name = {}
    for w, g in gold_pairs:
        gold_by_name[_group_base_name(gold_ws, w)] = (w, g)

    common_names = sorted(set(main_by_name.keys()) & set(gold_by_name.keys()))
    only_main = set(main_by_name.keys()) - set(gold_by_name.keys())
    only_gold = set(gold_by_name.keys()) - set(main_by_name.keys())
    if only_main:
        logger.warning("普独有 base: %s", only_main)
    if only_gold:
        logger.warning("金独有 base: %s", only_gold)

    # 不需要清空 main 原数据: merge_one_painting 会覆盖原 11 行
    # (parent + main children 写到 output_start_row..+10)
    col_map = locate_columns(main_ws)

    new_groups = []
    out_row = DATA_START_ROW
    # 关键: 在循环之前一次性快照所有 main 行, 避免后续 group 的写入污染源数据
    max_col_for_snapshot = max(main_ws.max_column, gold_ws.max_column)
    main_all_snapshots = {
        r: _snapshot_row(main_ws, r, max_col_for_snapshot)
        for g in main_groups
        for r in g
    }
    # 提前算 base name (此时 main_ws 还没被合并覆盖)
    main_base_names = {id(g): _group_base_name(main_ws, g) for g in main_groups}
    # 按 main_groups 原顺序遍历 (普文件原顺序就是最终顺序, 不按字母)
    for main_g in main_groups:
        name = main_base_names[id(main_g)]
        if name not in gold_by_name:
            logger.warning("跳过 main group (无 gold 配对): %s", name)
            continue
        wood_g, gold_g = gold_by_name[name]
        main_snapshots = [main_all_snapshots[r] for r in main_g]
        merged = merge_one_painting(
            main_snapshots=main_snapshots,
            gold_pair=(wood_g, gold_g),
            output_start_row=out_row,
            output_ws=main_ws,
            col_map=col_map,
            gold_wood_ws=gold_ws,
            gold_gold_ws=gold_ws,
            max_col=max_col_for_snapshot,
        )
        new_groups.append(merged)
        out_row += MERGED_GROUP_SIZE

    rewrite_sku(main_ws, new_groups, prefix)
    write_parent_sku_formulas(main_ws, new_groups)

    out = save_workbook(
        main_ws,
        main_path,
        main_sheet,
        output_path=str(output_path) if output_path else None,
    )
    logger.info("合并完成: 输出 %s, %d 画 × 21 行", out, len(new_groups))
    return out
