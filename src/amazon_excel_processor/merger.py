"""三文件合并模块

输入 3 个文件:
  - main_path: 普文件 (11 行/组, 含 Frame+Unframe 2 个 style)
  - wood_path: 木框文件 (6 行/组, 每画 1 个 group, 对应 Vintage Wood Grain Frame-style)
  - gold_path: 金框文件 (6 行/组, 每画 1 个 group, 对应 Vintage Ornate Gold Frame-style)

合并输出 21 行/组(1 parent + 4 style × 5 size) 的新 Excel, 顺序固定:
  1. Frame-style (5 尺寸, 来自 main)
  2. Unframe-style (5 尺寸, 来自 main)
  3. Vintage Wood Grain Frame-style (5 尺寸, 来自 wood)
  4. Vintage Ornate Gold Frame-style (5 尺寸, 来自 gold)

用户在 GUI 中按 [主, 木, 金] 顺序指定 3 个文件, 不再依赖"第 1 次/第 2 次"假设.
"""

import logging
import re
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
VARIANT_GROUP_SIZE = 6   # 木/金文件每画 1 个 group, 6 行
MAIN_GROUP_SIZE = 11     # 普文件每画 1 个 group, 11 行

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


def identify_file_role(groups):
    """根据 group 长度识别文件角色.

    Returns:
        ("main", None)  — 11 行/组, 普文件
        ("variant", None) — 6 行/组, 木或金文件 (具体由用户在 GUI 指定)
        ("unknown", None) — 无法识别
    """
    if not groups:
        return "unknown", None
    first = groups[0]
    if len(first) == MAIN_GROUP_SIZE:
        return "main", None
    if len(first) == VARIANT_GROUP_SIZE:
        return "variant", None
    return "unknown", None


# 保留旧名向后兼容
def identify_main_file(groups):
    role, _ = identify_file_role(groups)
    return (role == "main", role == "variant")


def _normalize_name_for_compare(name):
    if not name:
        return ""
    s = str(name)
    s = extract_base_title(s)
    s = remove_numeric_suffix(s)
    # 主文件有 "Henri Matisse - Harmony in Red" 格式, 木/金文件无 "-".
    # 配对前统一替换 "- " (连字符+空格) 为单空格.
    s = re.sub(r"\s*-\s*", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


def _group_base_name(ws, group, name_col=COL_PRODUCT_NAME):
    parent_row = group[0]
    v = ws.cell(row=parent_row, column=name_col).value
    return _normalize_name_for_compare(str(v) if v else "")


def index_groups_by_name(ws, groups, name_col=COL_PRODUCT_NAME):
    """把 groups 按 base name 索引, 返回 {base_name: group}.

    要求每个 base name 只出现 1 次 (每画 1 个 group).
    """
    by_name = {}
    for g in groups:
        name = _group_base_name(ws, g, name_col)
        if name in by_name:
            raise ValueError(
                f"base name '{name}' 在文件中出现多次, 期望每画 1 个 group"
            )
        by_name[name] = g
    return by_name


# 保留旧 API 向后兼容 (内部不再使用, 测试可能引用)
def pair_gold_groups(ws, groups, name_col=COL_PRODUCT_NAME):
    """[Deprecated] 旧 2 文件 API, 保留兼容. 推荐用 index_groups_by_name."""
    by_name = {}
    for g in groups:
        name = _group_base_name(ws, g, name_col)
        by_name.setdefault(name, []).append(g)
    pairs = []
    for name in sorted(by_name.keys()):
        gs = by_name[name]
        if len(gs) != 2:
            raise ValueError(
                f"base name '{name}' 有 {len(gs)} 个 group, 期望 2 个 (Wood + Gold)"
            )
        pairs.append((gs[0], gs[1]))
    return pairs


def _snapshot_row(ws, row, max_col):
    """快照一行数据, 返回 {col: value} dict (避免 ws 后续修改污染)."""
    return {c: ws.cell(row=row, column=c).value for c in range(1, max_col + 1)}


def _write_row(dst_ws, dst_row, snapshot, max_col):
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
    wood_group,
    gold_group,
    output_start_row,
    output_ws,
    col_map,
    wood_ws=None,
    gold_ws=None,
    max_col=None,
    ratio_type="3:2",
):
    """合并 1 画到 21 行结构.

    Args:
        main_snapshots: 11 元素 list (main group 行的快照, 必须提前快照避免被覆盖)
        wood_group: 木框文件的 6 行 group (来源 wood_ws)
        gold_group: 金框文件的 6 行 group (来源 gold_ws)
        output_start_row: 写到 output_ws 的起始行
        output_ws: 目标 worksheet
        col_map: 目标 ws 的列映射
        wood_ws: 木框文件 worksheet (用于读 wood_group 数据)
        gold_ws: 金框文件 worksheet (用于读 gold_group 数据)
        max_col: 列数
    """
    assert len(main_snapshots) == MAIN_GROUP_SIZE
    assert len(wood_group) == VARIANT_GROUP_SIZE
    assert len(gold_group) == VARIANT_GROUP_SIZE

    if max_col is None:
        max_col = output_ws.max_column

    wood_snapshots = [_snapshot_row(wood_ws, r, max_col) for r in wood_group]
    gold_snapshots = [_snapshot_row(gold_ws, r, max_col) for r in gold_group]

    merged_rows = []

    # parent (来自 main)
    parent_row = output_start_row
    _write_row(output_ws, parent_row, main_snapshots[0], max_col)
    merged_rows.append(parent_row)

    # main children: Frame×5 + Unframe×5 → output rows 1-10
    for i, snap in enumerate(main_snapshots[1:]):
        dst = output_start_row + 1 + i
        _write_row(output_ws, dst, snap, max_col)
        merged_rows.append(dst)

    # wood children: 5 → output rows 11-15
    for i, snap in enumerate(wood_snapshots[1:]):
        dst = output_start_row + 11 + i
        _write_row(output_ws, dst, snap, max_col)
        merged_rows.append(dst)

    # gold children: 5 → output rows 16-20
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
    wood_path,
    gold_path,
    shop,
    date,
    theme="",
    output_path=None,
):
    """三文件合并主入口.

    Args:
        main_path: 普文件 (11 行/组, Frame+Unframe)
        wood_path: 木框文件 (6 行/组, 每画 1 个 group)
        gold_path: 金框文件 (6 行/组, 每画 1 个 group)
        shop: 店铺缩写
        date: 日期
        theme: 主题缩写 (可空)
        output_path: 输出路径 (默认: {main_stem}_processed.xlsm)

    Returns:
        实际输出文件路径
    """
    main_path = Path(main_path)
    wood_path = Path(wood_path)
    gold_path = Path(gold_path)
    prefix = build_sku_prefix(shop, date, theme)
    logger.info("合并开始: main=%s, wood=%s, gold=%s, prefix=%s",
                main_path.name, wood_path.name, gold_path.name, prefix)

    main_wb, main_ws, main_sheet = load_workbook(main_path)
    wood_wb, wood_ws, _ = load_workbook(wood_path)
    gold_wb, gold_ws, _ = load_workbook(gold_path)

    main_groups = group_rows(main_ws, group_size=MAIN_GROUP_SIZE)
    wood_groups = group_rows(wood_ws, group_size=VARIANT_GROUP_SIZE)
    gold_groups = group_rows(gold_ws, group_size=VARIANT_GROUP_SIZE)

    main_role, _ = identify_file_role(main_groups)
    wood_role, _ = identify_file_role(wood_groups)
    gold_role, _ = identify_file_role(gold_groups)
    if main_role != "main":
        raise ValueError(
            f"主文件类型错误: {main_path.name} 是 {main_role}, 期望 main (11 行/组)"
        )
    if wood_role != "variant":
        raise ValueError(
            f"木框文件类型错误: {wood_path.name} 是 {wood_role}, 期望 variant (6 行/组)"
        )
    if gold_role != "variant":
        raise ValueError(
            f"金框文件类型错误: {gold_path.name} 是 {gold_role}, 期望 variant (6 行/组)"
        )

    main_by_name = index_groups_by_name(main_ws, main_groups)
    wood_by_name = index_groups_by_name(wood_ws, wood_groups)
    gold_by_name = index_groups_by_name(gold_ws, gold_groups)

    only_main = set(main_by_name.keys()) - set(wood_by_name.keys()) - set(gold_by_name.keys())
    only_wood = set(wood_by_name.keys()) - set(main_by_name.keys())
    only_gold = set(gold_by_name.keys()) - set(main_by_name.keys())
    if only_main:
        logger.warning("普独有 base (无木/金配对): %s", only_main)
    if only_wood:
        logger.warning("木独有 base (无普配对): %s", only_wood)
    if only_gold:
        logger.warning("金独有 base (无普配对): %s", only_gold)

    col_map = locate_columns(main_ws)

    # 关键: 在合并前一次性快照所有 main 行 + 提前算 base name
    max_col_for_snapshot = max(main_ws.max_column, wood_ws.max_column, gold_ws.max_column)
    main_all_snapshots = {
        r: _snapshot_row(main_ws, r, max_col_for_snapshot)
        for g in main_groups
        for r in g
    }
    main_base_names = {id(g): _group_base_name(main_ws, g) for g in main_groups}

    new_groups = []
    out_row = DATA_START_ROW
    # 按 main_groups 原顺序遍历 (普文件原顺序就是最终顺序)
    for main_g in main_groups:
        name = main_base_names[id(main_g)]
        if name not in wood_by_name:
            logger.warning("跳过 main group (无木框配对): %s", name)
            continue
        if name not in gold_by_name:
            logger.warning("跳过 main group (无金框配对): %s", name)
            continue
        wood_g = wood_by_name[name]
        gold_g = gold_by_name[name]
        main_snapshots = [main_all_snapshots[r] for r in main_g]
        merged = merge_one_painting(
            main_snapshots=main_snapshots,
            wood_group=wood_g,
            gold_group=gold_g,
            output_start_row=out_row,
            output_ws=main_ws,
            col_map=col_map,
            wood_ws=wood_ws,
            gold_ws=gold_ws,
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
