"""变体字段填充模块"""

import logging
import re

from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

COLOR_SEQUENCE = [
    "",
    "Frame-style", "Frame-style", "Frame-style", "Frame-style", "Frame-style",
    "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style",
]

SIZE_MAP_SEQUENCE = [
    "",
    "X-Small", "Small", "Medium", "Large", "X-Large",
    "X-Small", "Small", "Medium", "Large", "X-Large",
]

SIZE_32 = [
    "",
    "12L''x08W''", "18L''x12W''", "24L''x16W''", "30L''x20W''", "36L''x24W''",
    "12L''x08W''", "18L''x12W''", "24L''x16W''", "30L''x20W''", "36L''x24W''",
]

SIZE_SQUARE = [
    "",
    "12L''x12W''", "16L''x16W''", "20L''x20W''", "24L''x24W''", "28L''x28W''",
    "12L''x12W''", "16L''x16W''", "20L''x20W''", "24L''x24W''", "28L''x28W''",
]

LENGTH_32 = ["", 20, 30, 40, 50, 60, 20, 30, 40, 50, 60]
LENGTH_SQUARE = ["", 30, 40, 50, 60, 70, 30, 40, 50, 60, 70]

WIDTH_32 = ["", 30, 45, 60, 75, 90, 30, 45, 60, 75, 90]
# 正方形：两边相等，Width 与 Length 一致
WIDTH_SQUARE = ["", 30, 40, 50, 60, 70, 30, 40, 50, 60, 70]

WEIGHT_SEQUENCE = ["", 0.18, 0.28, 0.48, 0.68, 0.88, 0.02, 0.04, 0.07, 0.15, 0.25]

# 固定价格表（3:2 和正方形通用）：parent 空，5 Frame + 5 Unframe
PRICE_SEQUENCE = ["", 19.9, 29.9, 45, 75, 99, 11.9, 14.9, 19.9, 24.9, 34.9]

# ===== 合并模式 21 元素序列 =====
# 结构: [parent, Frame×5, Unframe×5, VintageWood×5, VintageOrnate×5]
# Wood 和 Gold 的 Size/SizeMap/Length/Width/Weight 与 Frame×5 一致
# Wood 和 Gold 的 Price 为: 26.9 / 39.9 / 59.9 / 99.9 / 129.9

COLOR_SEQUENCE_21 = [
    "",
    "Frame-style", "Frame-style", "Frame-style", "Frame-style", "Frame-style",
    "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style",
    "Vintage Wood Grain Frame-style", "Vintage Wood Grain Frame-style",
    "Vintage Wood Grain Frame-style", "Vintage Wood Grain Frame-style", "Vintage Wood Grain Frame-style",
    "Vintage Ornate Gold Frame-style", "Vintage Ornate Gold Frame-style",
    "Vintage Ornate Gold Frame-style", "Vintage Ornate Gold Frame-style", "Vintage Ornate Gold Frame-style",
]

SIZE_MAP_SEQUENCE_21 = [
    "",
    "X-Small", "Small", "Medium", "Large", "X-Large",
    "X-Small", "Small", "Medium", "Large", "X-Large",
    "X-Small", "Small", "Medium", "Large", "X-Large",
    "X-Small", "Small", "Medium", "Large", "X-Large",
]

SIZE_32_21 = [
    "",
    "12L''x08W''", "18L''x12W''", "24L''x16W''", "30L''x20W''", "36L''x24W''",
    "12L''x08W''", "18L''x12W''", "24L''x16W''", "30L''x20W''", "36L''x24W''",
    "12L''x08W''", "18L''x12W''", "24L''x16W''", "30L''x20W''", "36L''x24W''",
    "12L''x08W''", "18L''x12W''", "24L''x16W''", "30L''x20W''", "36L''x24W''",
]

LENGTH_32_21 = [""] + [20, 30, 40, 50, 60] * 4  # parent + 4 styles × 5 sizes = 21

WIDTH_32_21 = [""] + [30, 45, 60, 75, 90] * 4

# Weight: parent=空, Frame 和 Wood/Gold 均为 0.18-0.88, Unframe 为 0.02-0.25
WEIGHT_SEQUENCE_21 = [
    "",
    0.18, 0.28, 0.48, 0.68, 0.88,  # Frame×5
    0.02, 0.04, 0.07, 0.15, 0.25,  # Unframe×5
    0.18, 0.28, 0.48, 0.68, 0.88,  # Wood×5 (与 Frame×5 一致)
    0.18, 0.28, 0.48, 0.68, 0.88,  # Gold×5 (与 Frame×5 一致)
]

# Price: parent=空, Frame: 19.9/29.9/45/75/99, Unframe: 11.9/14.9/19.9/24.9/34.9,
#        Wood: 26.9/39.9/59.9/99.9/129.9, Gold: 26.9/39.9/59.9/99.9/129.9
PRICE_SEQUENCE_21 = [
    "",
    19.9, 29.9, 45, 75, 99,         # Frame×5
    11.9, 14.9, 19.9, 24.9, 34.9,   # Unframe×5
    26.9, 39.9, 59.9, 99.9, 129.9,  # Wood×5
    26.9, 39.9, 59.9, 99.9, 129.9,  # Gold×5
]

# List Price = Your Price (每行同步填)
# 这里直接指向 PRICE_SEQUENCE_21, merger 阶段可共用


def fill_list_price(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """List Price 列与 Your Price 列同步填相同值 (合并输出要求)。"""
    if "List Price" not in col_map or "Your Price" not in col_map:
        return
    list_col = col_map["List Price"]
    price_col = col_map["Your Price"]
    for row in rows:
        ws.cell(row=row, column=list_col).value = ws.cell(row=row, column=price_col).value


def fill_group_21(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
    ratio_type: str,
) -> None:
    """编排 21 行合并产品组的所有字段填充。

    与 fill_group() 类似, 但用 *_21 序列; 同时填 List Price。
    """
    fill_simple_fields_21(ws, rows, col_map)
    fill_color_21(ws, rows, col_map)
    fill_size_21(ws, rows, col_map, ratio_type)
    fill_size_map_21(ws, rows, col_map)
    fill_length_21(ws, rows, col_map, ratio_type)
    fill_width_21(ws, rows, col_map, ratio_type)
    fill_weight_21(ws, rows, col_map)
    fill_price_21(ws, rows, col_map)
    fill_list_price(ws, rows, col_map)
    clean_search_terms(ws, rows, col_map)
    fill_item_length_longer_edge_21(ws, rows, col_map)


def fill_simple_fields_21(ws, rows, col_map):
    simple_fills = {
        "Variation Theme": "color-size",
        "Paint Type": "Oil",
        "Color Map": "Multi",
    }
    for field_name, value in simple_fills.items():
        if field_name not in col_map:
            continue
        col_idx = col_map[field_name]
        for row in rows:
            ws.cell(row=row, column=col_idx).value = value


def fill_color_21(ws, rows, col_map):
    if "Color" not in col_map:
        return
    col_idx = col_map["Color"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = COLOR_SEQUENCE_21[i]


def fill_size_21(ws, rows, col_map, ratio_type):
    if "Size" not in col_map:
        return
    if ratio_type == "square":
        return
    col_idx = col_map["Size"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = SIZE_32_21[i]


def fill_size_map_21(ws, rows, col_map):
    if "Size Map" not in col_map:
        return
    col_idx = col_map["Size Map"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = SIZE_MAP_SEQUENCE_21[i]


def fill_length_21(ws, rows, col_map, ratio_type):
    if "Length" not in col_map:
        return
    col_idx = col_map["Length"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = LENGTH_32_21[i]


def fill_width_21(ws, rows, col_map, ratio_type):
    if "Width" not in col_map:
        return
    col_idx = col_map["Width"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = WIDTH_32_21[i]


def fill_weight_21(ws, rows, col_map):
    if "Weight" not in col_map:
        return
    col_idx = col_map["Weight"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = WEIGHT_SEQUENCE_21[i]


def fill_price_21(ws, rows, col_map):
    if "Your Price" not in col_map:
        return
    col_idx = col_map["Your Price"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = PRICE_SEQUENCE_21[i]


def fill_item_length_longer_edge_21(ws, rows, col_map):
    """parent 行填 1, child 行按尺寸填实际英寸值 (5 种尺寸循环 4 次)。"""
    if "Item Length Longer Edge" not in col_map:
        return
    col_idx = col_map["Item Length Longer Edge"]
    # 21 元素: parent=1, 然后 [12,18,24,30,36]×4
    values = [1] + [12, 18, 24, 30, 36] * 4
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = values[i]


def detect_ratio_type(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> str:
    """检测产品组的比例类型。

    解析 Size 列预填值中的两个数字，L==W 为正方形，L!=W 为 3:2。
    Size 列为空时默认 3:2（由脚本后续填充）。
    返回 "square" 或 "3:2"。
    """
    if "Size" not in col_map:
        return "3:2"
    size_col = col_map["Size"]
    for i, row in enumerate(rows):
        if i == 0:  # 跳过 parent 行（本就为空）
            continue
        value = ws.cell(row=row, column=size_col).value
        if value is None or not str(value).strip():
            continue
        # 解析 Size 值中的数字，比较前两个判断长宽是否相等
        numbers = re.findall(r"\d+", str(value))
        if len(numbers) >= 2:
            return "square" if numbers[0] == numbers[1] else "3:2"
        # 格式无法解析但有值，保守判断为正方形
        return "square"
    return "3:2"


def fill_simple_fields(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """批量填充简单字段：Variation Theme, Paint Type, Color Map。"""
    simple_fills = {
        "Variation Theme": "color-size",
        "Paint Type": "Oil",
        "Color Map": "Multi",
    }

    for field_name, value in simple_fills.items():
        if field_name not in col_map:
            continue
        col_idx = col_map[field_name]
        for row in rows:
            ws.cell(row=row, column=col_idx).value = value


def fill_color(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """按 11 行组填充 Color 列。"""
    if "Color" not in col_map:
        return
    col_idx = col_map["Color"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = COLOR_SEQUENCE[i]


def fill_size(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
    ratio_type: str,
) -> None:
    """按比例类型填充 Size 列。

    正方形组的 Size 列由用户预填，不覆盖；3:2 组填 SIZE_32。
    """
    if "Size" not in col_map:
        return
    if ratio_type == "square":
        return  # 正方形：保留用户预填值，不覆盖
    col_idx = col_map["Size"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = SIZE_32[i]


def fill_size_map(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """填充 Size Map 列。"""
    if "Size Map" not in col_map:
        return
    col_idx = col_map["Size Map"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = SIZE_MAP_SEQUENCE[i]


def fill_length(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
    ratio_type: str,
) -> None:
    """按比例类型填充 Length 列。"""
    if "Length" not in col_map:
        return
    col_idx = col_map["Length"]
    sequence = LENGTH_SQUARE if ratio_type == "square" else LENGTH_32
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = sequence[i]


def fill_width(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
    ratio_type: str,
) -> None:
    """按比例类型填充 Width 列。

    3:2：填宽度值；正方形：两边相等，Width 与 Length 一致。
    """
    if "Width" not in col_map:
        return
    col_idx = col_map["Width"]
    sequence = WIDTH_SQUARE if ratio_type == "square" else WIDTH_32
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = sequence[i]


def fill_weight(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """填充 Weight 列。"""
    if "Weight" not in col_map:
        return
    col_idx = col_map["Weight"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = WEIGHT_SEQUENCE[i]


def fill_price(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """填充 Your Price 列（固定价格表，3:2 和正方形通用）。"""
    if "Your Price" not in col_map:
        return
    col_idx = col_map["Your Price"]
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = PRICE_SEQUENCE[i]


def clean_search_terms(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """将 Search Terms 列中的下划线替换为空格。"""
    if "Search Terms" not in col_map:
        return
    col_idx = col_map["Search Terms"]
    for row in rows:
        value = ws.cell(row=row, column=col_idx).value
        if value is not None and isinstance(value, str) and "_" in value:
            ws.cell(row=row, column=col_idx).value = value.replace("_", " ")


def fill_item_length_longer_edge(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
) -> None:
    """只填充 Item Length Longer Edge 的 parent 行（第1行）为 1。"""
    if "Item Length Longer Edge" not in col_map:
        return
    col_idx = col_map["Item Length Longer Edge"]
    ws.cell(row=rows[0], column=col_idx).value = 1


def fill_group(
    ws: Worksheet,
    rows: list[int],
    col_map: dict[str, int],
    ratio_type: str,
) -> None:
    """编排单个产品组的所有字段填充。"""
    fill_simple_fields(ws, rows, col_map)
    fill_color(ws, rows, col_map)
    fill_size(ws, rows, col_map, ratio_type)
    fill_size_map(ws, rows, col_map)
    fill_length(ws, rows, col_map, ratio_type)
    fill_width(ws, rows, col_map, ratio_type)
    fill_weight(ws, rows, col_map)
    fill_price(ws, rows, col_map)
    clean_search_terms(ws, rows, col_map)
    fill_item_length_longer_edge(ws, rows, col_map)
