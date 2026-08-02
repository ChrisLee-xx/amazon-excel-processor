"""三文件合并模块

输入文件 (木/金可选):
  - main_path: 普文件 (11 行/组, 含 Frame+Unframe 2 个 style), 必填
  - wood_path: 木框文件 (6 行/组, 每画 1 个 group, 对应 Vintage Wood Grain Frame-style), 可选
  - gold_path: 金框文件 (6 行/组, 每画 1 个 group, 对应 Vintage Ornate Gold Frame-style), 可选

合并输出每组行数 = 1 + 5×(2 + 有木 + 有金): 11 / 16 / 21, 顺序固定:
  1. Frame-style (5 尺寸, 来自 main)
  2. Unframe-style (5 尺寸, 来自 main)
  3. Vintage Wood Grain Frame-style (5 尺寸, 来自 wood, 若提供)
  4. Vintage Ornate Gold Frame-style (5 尺寸, 来自 gold, 若提供)

用户在 GUI 中按 [主, 木, 金] 顺序指定文件 (木/金可留空跳过), 不再依赖"第 1 次/第 2 次"假设.
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
from .field_filler import (
    fill_group_merged,
    fill_list_price,
    build_active_styles,
    _build_sequences,
    STYLE_SPECS,
)
from .name_normalizer import (
    SIZES_32,
    extract_base_title,
    remove_numeric_suffix,
    replace_hyphens,
    replace_underscores,
    deduplicate_words,
    collapse_spaces,
)

logger = logging.getLogger(__name__)

MERGED_GROUP_SIZE = 21  # 最大: parent + 4 style × 5 size (wood+gold 都在)
VARIANT_GROUP_SIZE = 6   # 木/金文件每画 1 个 group, 6 行
MAIN_GROUP_SIZE = 11     # 普文件每画 1 个 group, 11 行


def merged_group_size(has_wood: bool, has_gold: bool) -> int:
    """合并输出每组行数 = 1 (parent) + 5 × (2 + 有木 + 有金)。

    木/金皆无 → 11; 只有其一 → 16; 都在 → 21。
    """
    return 1 + 5 * (2 + bool(has_wood) + bool(has_gold))

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
    """把 groups 按 base name 索引, 返回 {base_name: [group, ...]}.

    同名产品允许出现多次 (按文件中的出现顺序排列).
    配对时用 pair_counter 按顺序匹配 (普第 N 次 ↔ 木第 N 次 ↔ 金第 N 次).
    """
    from collections import defaultdict
    by_name = defaultdict(list)
    for g in groups:
        name = _group_base_name(ws, g, name_col)
        by_name[name].append(g)
    return dict(by_name)


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
    output_start_row,
    output_ws,
    col_map,
    wood_group=None,
    gold_group=None,
    wood_ws=None,
    gold_ws=None,
    max_col=None,
    ratio_type="3:2",
    mode="new",
):
    """合并 1 画到动态行数结构 (木/金可选).

    Args:
        main_snapshots: 11 元素 list (main group 行的快照, 必须提前快照避免被覆盖)
        wood_group: 木框文件的 6 行 group (来源 wood_ws); None 表示无木框
        gold_group: 金框文件的 6 行 group (来源 gold_ws); None 表示无金框
        output_start_row: 写到 output_ws 的起始行
        output_ws: 目标 worksheet
        col_map: 目标 ws 的列映射
        wood_ws: 木框文件 worksheet (用于读 wood_group 数据)
        gold_ws: 金框文件 worksheet (用于读 gold_group 数据)
        max_col: 列数
        mode: "new" = 新品上架 (全部行 normalize + fill + meta)
              "old_variant" = 老品补充变体 (普文件原 11 行保留不动, 仅变体行处理)

    输出行数 = 1 + 5×(2 + 有木 + 有金): 11 / 16 / 21。
    普文件恒为前 11 行 (parent + Frame×5 + Unframe×5), 变体行紧随其后。
    """
    assert len(main_snapshots) == MAIN_GROUP_SIZE
    has_wood = wood_group is not None
    has_gold = gold_group is not None
    if has_wood:
        assert wood_ws is not None and len(wood_group) == VARIANT_GROUP_SIZE
    if has_gold:
        assert gold_ws is not None and len(gold_group) == VARIANT_GROUP_SIZE
    # old_variant 模式必须有变体行可处理
    if mode == "old_variant" and not (has_wood or has_gold):
        raise ValueError("老品补充变体模式需要至少一个木框或金框文件")

    if max_col is None:
        max_col = output_ws.max_column

    active_styles = build_active_styles(has_wood, has_gold)

    merged_rows = []

    # parent (来自 main)
    parent_row = output_start_row
    _write_row(output_ws, parent_row, main_snapshots[0], max_col)
    merged_rows.append(parent_row)

    # main children: Frame×5 + Unframe×5 → output rows 1-10 (恒定 10 行)
    for i, snap in enumerate(main_snapshots[1:]):
        dst = output_start_row + 1 + i
        _write_row(output_ws, dst, snap, max_col)
        merged_rows.append(dst)

    # 变体 children 紧跟在 main 之后 (从 output row 11 起), 按动态紧凑偏移写入
    next_offset = 11  # main 占 11 行 (1 parent + 10 children)
    if has_wood:
        wood_snapshots = [_snapshot_row(wood_ws, r, max_col) for r in wood_group]
        for i, snap in enumerate(wood_snapshots[1:]):
            dst = output_start_row + next_offset + i
            _write_row(output_ws, dst, snap, max_col)
            merged_rows.append(dst)
        next_offset += 5
    if has_gold:
        gold_snapshots = [_snapshot_row(gold_ws, r, max_col) for r in gold_group]
        for i, snap in enumerate(gold_snapshots[1:]):
            dst = output_start_row + next_offset + i
            _write_row(output_ws, dst, snap, max_col)
            merged_rows.append(dst)
        next_offset += 5

    if mode == "new":
        # 新品上架: 全部行 normalize + fill + meta
        normalize_group_merged(output_ws, merged_rows, COL_PRODUCT_NAME, ratio_type, active_styles)
        fill_group_merged(output_ws, merged_rows, col_map, ratio_type, active_styles)
        _fill_meta_columns(output_ws, merged_rows, col_map)
    elif mode == "old_variant":
        # 老品补充变体: 普文件原 11 行 (rows[0:11]) 完全不动
        # 仅对变体行 (rows[11:]) 做 normalize + fill + meta
        variant_rows = merged_rows[11:]
        variant_styles = active_styles[2:]  # 去掉 frame/unframe, 只剩变体 style
        _normalize_variant_names(output_ws, merged_rows, variant_rows, COL_PRODUCT_NAME,
                                 ratio_type, variant_styles)
        _fill_variant_fields(output_ws, variant_rows, col_map, ratio_type, variant_styles)
        _fill_meta_columns_variant(output_ws, variant_rows, col_map)
    else:
        raise ValueError(f"未知 mode: {mode}")

    return merged_rows


def _normalize_variant_names(ws, all_rows, variant_rows, name_col, ratio_type="3:2",
                             variant_styles=None):
    """老品补充模式: 只对变体行 (variant_rows) 做 Product Name 规范化.

    base_title 从 parent 行 (all_rows[0]) 提取, 但不修改 parent 行。
    variant_styles 决定各行 label (如 ["wood"] 或 ["gold"] 或 ["wood","gold"])。
    """
    if variant_styles is None:
        variant_styles = ["wood", "gold"]
    sizes = SIZES_32
    parent_cell = ws.cell(row=all_rows[0], column=name_col)
    parent_value = parent_cell.value
    if parent_value is None:
        return
    base_title = extract_base_title(str(parent_value))
    base_title = remove_numeric_suffix(base_title)
    base_title = collapse_spaces(base_title)

    # 变体 labels (不含 parent 占位): 每个 style × 5
    var_labels = []
    for key in variant_styles:
        var_labels.extend([STYLE_SPECS[key]["label"]] * 5)

    for i, row in enumerate(variant_rows):
        cell = ws.cell(row=row, column=name_col)
        value = cell.value
        if value is None:
            continue
        label = var_labels[i]
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


def _fill_variant_fields(ws, variant_rows, col_map, ratio_type="3:2", variant_styles=None):
    """老品补充模式: 只对变体行 (variant_rows) 填充字段.

    variant_styles 决定 color/price/weight 等序列 (如 ["wood"] / ["gold"] / ["wood","gold"])。
    序列从 index 0 应用到 variant_rows (variant_rows 不含 parent)。
    """
    if variant_styles is None:
        variant_styles = ["wood", "gold"]
    # 构建仅含变体 style 的逐行序列 (去掉 parent 占位 index 0)
    full = _build_sequences(variant_styles)
    vseq = {k: v[1:] for k, v in full.items()}

    def _set(field_name, seq_key):
        if field_name not in col_map:
            return
        col = col_map[field_name]
        for i, row in enumerate(variant_rows):
            ws.cell(row=row, column=col).value = vseq[seq_key][i]

    _set("Color", "color")
    if ratio_type != "square":
        _set("Size", "size_32")
    _set("Size Map", "size_map")
    _set("Length", "length")
    _set("Width", "width")
    _set("Weight", "weight")
    _set("Your Price", "price")
    _set("Item Length Longer Edge", "edge")

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


def normalize_group_merged(ws, rows, name_col, ratio_type="3:2", active_styles=None):
    """对合并产品组 (动态行数) 执行 Product Name 规范化。

    active_styles 决定各行 label (含 parent 占位)。始终用 SIZES_32 构造尺寸名
    (历史行为, 不随 ratio_type 切换)。
    """
    if active_styles is None:
        active_styles = ["frame", "unframe", "wood", "gold"]
    labels = _build_sequences(active_styles)["labels"]
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
            label = labels[i]
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
    wood_path=None,
    gold_path=None,
    sku_prefix="",
    mode="new",
    output_path=None,
):
    """合并主入口 (木/金可选).

    Args:
        main_path: 普文件 (11 行/组, Frame+Unframe), 必填
        wood_path: 木框文件 (6 行/组, 每画 1 个 group), 可选; None 表示无木框
        gold_path: 金框文件 (6 行/组, 每画 1 个 group), 可选; None 表示无金框
        sku_prefix: SKU 前缀 (如 HM725, 推荐格式 店铺名+日期+主题)
        mode: "new" = 新品上架 (全部 SKU 重写)
              "old_variant" = 老品补充变体 (普文件原 SKU 保留, 仅变体重写);
                              此模式需要至少一个木/金文件
        output_path: 输出路径 (默认: {main_stem}_processed.xlsm)

    输出每组行数 = 1 + 5×(2 + 有木 + 有金): 11 / 16 / 21。

    Returns:
        实际输出文件路径
    """
    main_path = Path(main_path)
    has_wood = wood_path is not None
    has_gold = gold_path is not None
    if mode == "old_variant" and not (has_wood or has_gold):
        raise ValueError("老品补充变体模式需要至少一个木框或金框文件")

    prefix = build_sku_prefix(sku_prefix)
    logger.info("合并开始: main=%s, wood=%s, gold=%s, prefix=%s, mode=%s",
                main_path.name,
                wood_path.name if has_wood else "无",
                gold_path.name if has_gold else "无",
                prefix, mode)

    main_wb, main_ws, main_sheet = load_workbook(main_path)
    wood_ws = gold_ws = None
    if has_wood:
        _, wood_ws, _ = load_workbook(Path(wood_path))
    if has_gold:
        _, gold_ws, _ = load_workbook(Path(gold_path))

    main_groups = group_rows(main_ws, group_size=MAIN_GROUP_SIZE)
    wood_groups = group_rows(wood_ws, group_size=VARIANT_GROUP_SIZE) if has_wood else []
    gold_groups = group_rows(gold_ws, group_size=VARIANT_GROUP_SIZE) if has_gold else []

    main_role, _ = identify_file_role(main_groups)
    if main_role != "main":
        raise ValueError(
            f"主文件类型错误: {main_path.name} 是 {main_role}, 期望 main (11 行/组)"
        )
    if has_wood:
        wood_role, _ = identify_file_role(wood_groups)
        if wood_role != "variant":
            raise ValueError(
                f"木框文件类型错误: {wood_path.name} 是 {wood_role}, 期望 variant (6 行/组)"
            )
    if has_gold:
        gold_role, _ = identify_file_role(gold_groups)
        if gold_role != "variant":
            raise ValueError(
                f"金框文件类型错误: {gold_path.name} 是 {gold_role}, 期望 variant (6 行/组)"
            )

    main_by_name = index_groups_by_name(main_ws, main_groups, file_label="普文件")
    wood_by_name = index_groups_by_name(wood_ws, wood_groups, file_label="木框文件") if has_wood else {}
    gold_by_name = index_groups_by_name(gold_ws, gold_groups, file_label="金框文件") if has_gold else {}

    # 检查配对完整性 (no-op 桩, 实际配对在主循环 pair_counter 处理)
    _check_pairing(main_by_name, wood_by_name, gold_by_name,
                   main_ws, wood_ws, gold_ws)

    col_map = locate_columns(main_ws)

    # 关键: 在合并前一次性快照所有 main 行 + 提前算 base name
    snap_cols = [main_ws.max_column]
    if has_wood:
        snap_cols.append(wood_ws.max_column)
    if has_gold:
        snap_cols.append(gold_ws.max_column)
    max_col_for_snapshot = max(snap_cols)
    main_all_snapshots = {
        r: _snapshot_row(main_ws, r, max_col_for_snapshot)
        for g in main_groups
        for r in g
    }
    main_base_names = {id(g): _group_base_name(main_ws, g) for g in main_groups}

    group_size = merged_group_size(has_wood, has_gold)

    # 追踪每个 base name 已配对次数 (支持同名多 group 按顺序配对)
    # 普/木/金文件的产品顺序一致, 同名产品按出现顺序配对
    pair_counter = {}
    skipped = []  # 记录配不上的 main group

    new_groups = []
    out_row = DATA_START_ROW
    for main_g in main_groups:
        name = main_base_names[id(main_g)]
        idx = pair_counter.get(name, 0)
        pair_counter[name] = idx + 1

        wood_list = wood_by_name.get(name, []) if has_wood else None
        gold_list = gold_by_name.get(name, []) if has_gold else None
        # 文件整体缺失不报错; 只有"文件存在但缺该画"才进 skipped
        if (has_wood and idx >= len(wood_list)) or (has_gold and idx >= len(gold_list)):
            main_raw = _get_raw_name(main_ws, main_g, COL_PRODUCT_NAME)
            skipped.append((name, main_raw, idx,
                            len(wood_list) if has_wood else None,
                            len(gold_list) if has_gold else None))
            continue
        wood_g = wood_list[idx] if has_wood else None
        gold_g = gold_list[idx] if has_gold else None
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
        out_row += group_size

    # 配不上的报错 (用模糊匹配给候选)
    if skipped:
        _raise_pairing_error(skipped, wood_by_name, gold_by_name, wood_ws, gold_ws,
                             has_wood, has_gold)

    rewrite_sku(main_ws, new_groups, prefix, mode=mode)
    write_parent_sku_formulas(main_ws, new_groups, mode=mode)

    out = save_workbook(
        main_ws,
        main_path,
        main_sheet,
        output_path=str(output_path) if output_path else None,
    )
    logger.info("合并完成: 输出 %s, %d 画 × %d 行/组", out, len(new_groups), group_size)
    return out


def _get_raw_name(ws, group, name_col):
    """获取 group parent 行的原始 Product Name."""
    v = ws.cell(row=group[0], column=name_col).value
    return str(v) if v else ""


def _check_pairing(main_by_name, wood_by_name, gold_by_name,
                   main_ws, wood_ws, gold_ws):
    """预检查: 普文件的 base name 是否都能在木/金文件中找到 (允许同名多 group).

    只在 base name 完全不存在时才警告 (不是每次配对都检查).
    同名多 group 的按顺序配对在 merge_files 主循环中处理.
    """
    # 此函数保留为 no-op, 实际配对检查在 merge_files 主循环中通过 pair_counter 处理
    pass


def _raise_pairing_error(skipped, wood_by_name, gold_by_name, wood_ws, gold_ws,
                         has_wood=True, has_gold=True):
    """配不上时用模糊匹配找候选, 报错列出。

    木/金文件整体缺失 (has_wood/has_gold=False) 时, 对应文件报"未提供"且不列候选。
    只有"文件存在但缺该画"才会到达这里 (文件整体缺失在主循环不会进 skipped)。
    """
    errors = []
    for name, main_raw, idx, wood_count, gold_count in skipped:
        msg = f"  普文件产品找不到配对:\n"
        msg += f"    Product Name: {main_raw}\n"
        msg += f"    (归一化后: '{name}', 第 {idx+1} 次出现)\n"
        wood_disp = f"{wood_count} 个" if has_wood else "未提供"
        gold_disp = f"{gold_count} 个" if has_gold else "未提供"
        msg += f"    木框文件中该名 {wood_disp}, 金框文件中该名 {gold_disp}\n"
        # 模糊匹配给候选 (仅对实际提供的文件)
        if has_wood:
            all_wood_keys = list(wood_by_name.keys())
            wood_candidates = _find_close_matches(name, all_wood_keys)
            if wood_candidates:
                msg += f"    最接近的木框候选:\n"
                for cname, ratio in wood_candidates[:3]:
                    wood_raw = _get_raw_name(wood_ws, wood_by_name[cname][0], COL_PRODUCT_NAME)
                    msg += f"      [{ratio:.0%}] {wood_raw}\n"
        if has_gold:
            all_gold_keys = list(gold_by_name.keys())
            gold_candidates = _find_close_matches(name, all_gold_keys)
            if gold_candidates:
                msg += f"    最接近的金框候选:\n"
                for cname, ratio in gold_candidates[:3]:
                    gold_raw = _get_raw_name(gold_ws, gold_by_name[cname][0], COL_PRODUCT_NAME)
                    msg += f"      [{ratio:.0%}] {gold_raw}\n"
        errors.append(msg)

    raise ValueError(
        f"产品配对失败, {len(skipped)} 个普文件产品在木/金文件中找不到匹配:\n\n"
        + "\n".join(errors)
        + "\n请检查 Product Name 是否一致 (允许标点/扩展名/括号差异), "
        "或手动修改后重试"
    )
