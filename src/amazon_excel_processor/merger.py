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


def build_sku_prefix(sku):
    """SKU 前缀直接用用户输入的字符串 (推荐格式: 店铺名+日期+主题, 如 HM725)."""
    return sku.strip()


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
    """归一化 Product Name 用于 base name 配对.

    处理常见差异:
    - 去 Frame-/Unframe- 之后的内容
    - 去 -1, -2 等数字后缀
    - 去文件扩展名 (.jpg, .png 等)
    - 去括号及内容 (如 (1), (copy))
    - 去所有标点符号 (只保留字母、数字、中文、空格)
    - 合并空格、转小写
    """
    if not name:
        return ""
    s = str(name)
    s = extract_base_title(s)
    s = remove_numeric_suffix(s)
    # 去文件扩展名
    s = re.sub(r"\.(jpg|jpeg|png|gif|bmp|webp|tiff?)\b", "", s, flags=re.IGNORECASE)
    # 去括号及内容
    s = re.sub(r"\([^)]*\)", "", s)
    # 去所有非字母数字中文空格字符 (标点符号等)
    s = re.sub(r"[^a-z0-9\u4e00-\u9fff\s]", " ", s.lower())
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _find_close_matches(target, candidates, cutoff=0.6):
    """用 difflib 找最接近的候选, 返回 [(name, ratio), ...]."""
    from difflib import SequenceMatcher
    scored = []
    for c in candidates:
        ratio = SequenceMatcher(None, target, c).ratio()
        if ratio >= cutoff:
            scored.append((c, ratio))
    scored.sort(key=lambda x: -x[1])
    return scored


def _group_base_name(ws, group, name_col=COL_PRODUCT_NAME):
    parent_row = group[0]
    v = ws.cell(row=parent_row, column=name_col).value
    return _normalize_name_for_compare(str(v) if v else "")


def index_groups_by_name(ws, groups, name_col=COL_PRODUCT_NAME, file_label=""):
    """把 groups 按 base name 索引, 返回 {base_name: group}.

    要求每个 base name 只出现 1 次 (每画 1 个 group).
    重复时报错, 给出具体 Product Name 和行号.
    """
    by_name = {}
    label_prefix = f"[{file_label}] " if file_label else ""
    for g in groups:
        name = _group_base_name(ws, g, name_col)
        if name in by_name:
            prev_g = by_name[name]
            raw_name = ws.cell(row=g[0], column=name_col).value
            prev_raw = ws.cell(row=prev_g[0], column=name_col).value
            raise ValueError(
                f"{label_prefix}检测到同名产品重复:\n"
                f"  Product Name: {raw_name}\n"
                f"  重复位置: 行 {prev_g[0]} 和 行 {g[0]}\n"
                f"  请检查 Excel 文件, 确保每个产品名只出现 1 次"
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
    mode="new",
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
        mode: "new" = 新品上架 (全部 21 行 normalize + fill + meta)
              "old_variant" = 老品补充变体 (普文件原 11 行保留不动, 仅 Wood/Gold 行处理)
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

    if mode == "new":
        # 新品上架: 全部 21 行 normalize + fill + meta
        normalize_group_21(output_ws, merged_rows, COL_PRODUCT_NAME, ratio_type)
        fill_group_21(output_ws, merged_rows, col_map, ratio_type)
        _fill_meta_columns(output_ws, merged_rows, col_map)
    elif mode == "old_variant":
        # 老品补充变体: 普文件原 11 行 (rows[0:11]) 完全不动
        # 仅对 Wood/Gold 行 (rows[11:21]) 做 normalize + fill + meta
        variant_rows = merged_rows[11:]
        _normalize_variant_names(output_ws, merged_rows, variant_rows, COL_PRODUCT_NAME, ratio_type)
        _fill_variant_fields(output_ws, variant_rows, col_map, ratio_type)
        _fill_meta_columns_variant(output_ws, variant_rows, col_map)
    else:
        raise ValueError(f"未知 mode: {mode}")

    return merged_rows


def _normalize_variant_names(ws, all_rows, variant_rows, name_col, ratio_type="3:2"):
    """老品补充模式: 只对 Wood/Gold 行 (variant_rows) 做 Product Name 规范化.

    base_title 从 parent 行 (all_rows[0]) 提取, 但不修改 parent 行.
    """
    sizes = SIZES_32
    parent_cell = ws.cell(row=all_rows[0], column=name_col)
    parent_value = parent_cell.value
    if parent_value is None:
        return
    base_title = extract_base_title(str(parent_value))
    base_title = remove_numeric_suffix(base_title)
    base_title = collapse_spaces(base_title)

    # variant_rows 对应 VARIANT_LABELS_21[11:21], size_idx 从 0 开始
    for i, row in enumerate(variant_rows):
        cell = ws.cell(row=row, column=name_col)
        value = cell.value
        if value is None:
            continue
        label = VARIANT_LABELS_21[11 + i]
        size_idx = i % 5
        size = sizes[size_idx]
        name = f"{base_title} {label} {size}"
        name = collapse_spaces(name)
        name = remove_numeric_suffix(name)
        name = replace_hyphens(name)
        name = replace_underscores(name)
        name = deduplicate_words(name)
        name = collapse_spaces(name)
        cell.value = name


def _fill_variant_fields(ws, variant_rows, col_map, ratio_type="3:2"):
    """老品补充模式: 只对 Wood/Gold 行 (variant_rows) 填充字段.

    使用 COLOR_SEQUENCE_21[11:21] / PRICE_SEQUENCE_21[11:21] 等.
    """
    from .field_filler import (
        COLOR_SEQUENCE_21, SIZE_MAP_SEQUENCE_21, SIZE_32_21,
        LENGTH_32_21, WIDTH_32_21, WEIGHT_SEQUENCE_21, PRICE_SEQUENCE_21,
    )

    # Wood/Gold 行在 21 元素序列中的偏移量是 11
    offset = 11

    if "Color" in col_map:
        col = col_map["Color"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = COLOR_SEQUENCE_21[offset + i]

    if "Size" in col_map and ratio_type != "square":
        col = col_map["Size"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = SIZE_32_21[offset + i]

    if "Size Map" in col_map:
        col = col_map["Size Map"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = SIZE_MAP_SEQUENCE_21[offset + i]

    if "Length" in col_map:
        col = col_map["Length"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = LENGTH_32_21[offset + i]

    if "Width" in col_map:
        col = col_map["Width"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = WIDTH_32_21[offset + i]

    if "Weight" in col_map:
        col = col_map["Weight"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = WEIGHT_SEQUENCE_21[offset + i]

    if "Your Price" in col_map:
        col = col_map["Your Price"]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = PRICE_SEQUENCE_21[offset + i]

    # List Price = Your Price
    if "List Price" in col_map and "Your Price" in col_map:
        lp_col = col_map["List Price"]
        yp_col = col_map["Your Price"]
        for row in variant_rows:
            ws.cell(row=row, column=lp_col).value = ws.cell(row=row, column=yp_col).value

    # Variation Theme / Paint Type / Color Map
    simple_fills = {"Variation Theme": "color-size", "Paint Type": "Oil", "Color Map": "Multi"}
    for field, val in simple_fills.items():
        if field in col_map:
            col = col_map[field]
            for row in variant_rows:
                ws.cell(row=row, column=col).value = val

    # Item Length Longer Edge: [12, 18, 24, 30, 36] × 2 (Wood + Gold)
    if "Item Length Longer Edge" in col_map:
        col = col_map["Item Length Longer Edge"]
        edge_values = [12, 18, 24, 30, 36] * 2
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = edge_values[i]

    # Search Terms: 替换下划线
    if "Search Terms" in col_map:
        col = col_map["Search Terms"]
        for row in variant_rows:
            v = ws.cell(row=row, column=col).value
            if v is not None and isinstance(v, str) and "_" in v:
                ws.cell(row=row, column=col).value = v.replace("_", " ")


def _fill_meta_columns_variant(ws, variant_rows, col_map):
    """老品补充模式: 只对 Wood/Gold 行设 Parentage/Relationship Type/Variation Theme/Package Level."""
    if "Variation Theme" in col_map:
        col = col_map["Variation Theme"]
        for r in variant_rows:
            ws.cell(row=r, column=col).value = "color-size"
    if "Package Level" in col_map:
        col = col_map["Package Level"]
        for r in variant_rows:
            ws.cell(row=r, column=col).value = "unit"
    if "Parentage" in col_map:
        col = col_map["Parentage"]
        for r in variant_rows:
            ws.cell(row=r, column=col).value = "Child"
    if "Relationship Type" in col_map:
        col = col_map["Relationship Type"]
        for r in variant_rows:
            ws.cell(row=r, column=col).value = "Variation"


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


def rewrite_sku(ws, groups, prefix, sku_col=COL_SELLER_SKU, mode="new"):
    """重写 Seller SKU.

    Args:
        ws: 目标 worksheet
        groups: 多个 21 行 group
        prefix: SKU 前缀 (如 HM725)
        sku_col: Seller SKU 列
        mode: "new" = 新品上架 (全部 21 行 × N 画从 prefix-1 连续编号)
              "old_variant" = 老品补充变体 (普文件原 11 行 SKU 保留,
                              新增 Wood/Gold 10 行 × N 画从 prefix-1 连续编号)
    """
    if mode == "new":
        counter = 1
        for group in groups:
            for row in group:
                ws.cell(row=row, column=sku_col).value = f"{prefix}-{counter}"
                counter += 1
    elif mode == "old_variant":
        # 普文件原 11 行 (group[0:11]) SKU 保留, 新增 Wood/Gold 行 (group[11:21]) 重写
        counter = 1
        for group in groups:
            for row in group[11:]:  # Wood×5 + Gold×5 = 10 行
                ws.cell(row=row, column=sku_col).value = f"{prefix}-{counter}"
                counter += 1
    else:
        raise ValueError(f"未知 mode: {mode}, 期望 'new' 或 'old_variant'")


def write_parent_sku_formulas(ws, groups, parent_sku_col=COL_PARENT_SKU, seller_sku_col=COL_SELLER_SKU, mode="new"):
    """写 Parent SKU 公式.

    Args:
        ws: 目标 worksheet
        groups: 多个 21 行 group
        parent_sku_col: Parent SKU 列
        seller_sku_col: Seller SKU 列
        mode: "new" = 新品上架 (全部 21 行 parent SKU 公式重写)
              "old_variant" = 老品补充变体 (普文件原 11 行保留,
                              新增 Wood/Gold 行从 =AA{prev_unframe_last} 开始链式引用)
    """
    seller_letter = _col_letter(seller_sku_col)
    parent_sku_letter = _col_letter(parent_sku_col)

    for group in groups:
        if len(group) < 2:
            continue

        if mode == "new":
            # 全部重写: parent 清空, 第 1 child =B{parent}, 后续 =AA{prev}
            parent_row = group[0]
            first_child = group[1]
            ws.cell(row=parent_row, column=parent_sku_col).value = None
            ws.cell(row=first_child, column=parent_sku_col).value = f"={seller_letter}{parent_row}"
            for i in range(2, len(group)):
                prev_row = group[i - 1]
                ws.cell(row=group[i], column=parent_sku_col).value = f"={parent_sku_letter}{prev_row}"

        elif mode == "old_variant":
            # 普文件原 11 行 (group[0:11]) parent SKU 公式保留
            # 新增 Wood/Gold 行 (group[11:21]) 从 =AA{group[10]} 开始链式
            # group[10] = Unframe 最后一个, group[11] = Wood 第 1 个
            prev_row = group[10]  # Unframe 最后一个
            for i in range(11, len(group)):
                ws.cell(row=group[i], column=parent_sku_col).value = f"={parent_sku_letter}{prev_row}"
                prev_row = group[i]

        else:
            raise ValueError(f"未知 mode: {mode}, 期望 'new' 或 'old_variant'")


def fill_list_price_synced(ws, rows, col_map):
    fill_list_price(ws, rows, col_map)


def merge_files(
    main_path,
    wood_path,
    gold_path,
    sku_prefix,
    mode="new",
    output_path=None,
):
    """三文件合并主入口.

    Args:
        main_path: 普文件 (11 行/组, Frame+Unframe)
        wood_path: 木框文件 (6 行/组, 每画 1 个 group)
        gold_path: 金框文件 (6 行/组, 每画 1 个 group)
        sku_prefix: SKU 前缀 (如 HM725, 推荐格式 店铺名+日期+主题)
        mode: "new" = 新品上架 (全部 SKU 重写)
              "old_variant" = 老品补充变体 (普文件原 SKU 保留, 仅 Wood/Gold 重写)
        output_path: 输出路径 (默认: {main_stem}_processed.xlsm)

    Returns:
        实际输出文件路径
    """
    main_path = Path(main_path)
    wood_path = Path(wood_path)
    gold_path = Path(gold_path)
    prefix = build_sku_prefix(sku_prefix)
    logger.info("合并开始: main=%s, wood=%s, gold=%s, prefix=%s, mode=%s",
                main_path.name, wood_path.name, gold_path.name, prefix, mode)

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

    main_by_name = index_groups_by_name(main_ws, main_groups, file_label="普文件")
    wood_by_name = index_groups_by_name(wood_ws, wood_groups, file_label="木框文件")
    gold_by_name = index_groups_by_name(gold_ws, gold_groups, file_label="金框文件")

    # 检查配对完整性, 配不上时报错并给出模糊匹配候选
    _check_pairing(main_by_name, wood_by_name, gold_by_name,
                   main_ws, wood_ws, gold_ws)

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
            mode=mode,
        )
        new_groups.append(merged)
        out_row += MERGED_GROUP_SIZE

    rewrite_sku(main_ws, new_groups, prefix, mode=mode)
    write_parent_sku_formulas(main_ws, new_groups, mode=mode)

    out = save_workbook(
        main_ws,
        main_path,
        main_sheet,
        output_path=str(output_path) if output_path else None,
    )
    logger.info("合并完成: 输出 %s, %d 画 × 21 行", out, len(new_groups))
    return out


def _get_raw_name(ws, group, name_col):
    """获取 group parent 行的原始 Product Name."""
    v = ws.cell(row=group[0], column=name_col).value
    return str(v) if v else ""


def _check_pairing(main_by_name, wood_by_name, gold_by_name,
                   main_ws, wood_ws, gold_ws):
    """检查普/木/金文件的 base name 配对完整性.

    配不上时用模糊匹配找候选, 报错时列出最接近的候选和相似度.
    不自动配对 (避免配错), 只给用户参考信息.
    """
    main_keys = set(main_by_name.keys())
    wood_keys = set(wood_by_name.keys())
    gold_keys = set(gold_by_name.keys())

    main_no_wood = main_keys - wood_keys
    main_no_gold = main_keys - gold_keys

    if not main_no_wood and not main_no_gold:
        return  # 全部配对成功

    errors = []
    for name in sorted(main_no_wood):
        main_raw = _get_raw_name(main_ws, main_by_name[name], COL_PRODUCT_NAME)
        wood_candidates = _find_close_matches(name, list(wood_keys))
        msg = f"  普文件产品找不到木框配对:\n"
        msg += f"    普文件 Product Name: {main_raw}\n"
        msg += f"    (归一化后: '{name}')\n"
        if wood_candidates:
            msg += f"    最接近的木框候选:\n"
            for cname, ratio in wood_candidates[:3]:
                wood_raw = _get_raw_name(wood_ws, wood_by_name[cname], COL_PRODUCT_NAME)
                msg += f"      [{ratio:.0%}] {wood_raw}\n"
        else:
            msg += f"    木框文件中无相似产品\n"
        errors.append(msg)

    for name in sorted(main_no_gold):
        main_raw = _get_raw_name(main_ws, main_by_name[name], COL_PRODUCT_NAME)
        gold_candidates = _find_close_matches(name, list(gold_keys))
        msg = f"  普文件产品找不到金框配对:\n"
        msg += f"    普文件 Product Name: {main_raw}\n"
        msg += f"    (归一化后: '{name}')\n"
        if gold_candidates:
            msg += f"    最接近的金框候选:\n"
            for cname, ratio in gold_candidates[:3]:
                gold_raw = _get_raw_name(gold_ws, gold_by_name[cname], COL_PRODUCT_NAME)
                msg += f"      [{ratio:.0%}] {gold_raw}\n"
        else:
            msg += f"    金框文件中无相似产品\n"
        errors.append(msg)

    raise ValueError(
        "产品配对失败, 以下普文件产品在木/金文件中找不到匹配:\n\n"
        + "\n".join(errors)
        + "\n请检查 Product Name 是否一致 (允许标点/扩展名/括号差异), "
        "或手动修改后重试"
    )
