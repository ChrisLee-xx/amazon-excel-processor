"""变体字段填充模块"""

import logging

from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

SQUARE_KEYWORDS = {"12x12", "16x16", "20x20", "24x24", "28x28"}

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

WEIGHT_SEQUENCE = ["", 0.18, 0.28, 0.48, 0.68, 0.88, 0.02, 0.04, 0.07, 0.15, 0.25]


def detect_ratio_type(ws: Worksheet, rows: list[int], product_name_col: int) -> str:
    """检测产品组的比例类型。

    检查所有行的 Product Name 中是否包含正方形尺寸关键词。
    返回 "square" 或 "3:2"。
    """
    for row in rows:
        value = ws.cell(row=row, column=product_name_col).value
        if value is None:
            continue
        text = str(value)
        for kw in SQUARE_KEYWORDS:
            if kw in text:
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
    """按比例类型填充 Size 列。"""
    if "Size" not in col_map:
        return
    col_idx = col_map["Size"]
    sequence = SIZE_SQUARE if ratio_type == "square" else SIZE_32
    for i, row in enumerate(rows):
        ws.cell(row=row, column=col_idx).value = sequence[i]


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
    fill_weight(ws, rows, col_map)
    clean_search_terms(ws, rows, col_map)
