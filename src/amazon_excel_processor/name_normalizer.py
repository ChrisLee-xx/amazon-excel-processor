"""Product Name 规范化模块"""

import logging
import re
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

FRAME_SPLIT_PATTERN = re.compile(r"\s+(Frame|Unframe)-", re.IGNORECASE)
NUMERIC_SUFFIX_PATTERN = re.compile(r"-(\d+)(?=\s|$)")

# 固定尺寸顺序（与 md 文档一致）
SIZES_32 = [
    "08x12inch(20x30cm)",
    "12x18inch(30x45cm)",
    "16x24inch(40x60cm)",
    "20x30inch(50x75cm)",
    "24x36inch(60x90cm)",
]
SIZES_SQUARE = [
    "12x12inch(30x30cm)",
    "16x16inch(40x40cm)",
    "20x20inch(50x50cm)",
    "24x24inch(60x60cm)",
    "28x28inch(70x70cm)",
]

# 固定的 11 行结构：[parent, Frame×5, Unframe×5]
VARIANT_LABELS = [
    None,  # parent
    "Frame-style", "Frame-style", "Frame-style", "Frame-style", "Frame-style",
    "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style",
]

# 21 行结构（合并模式输出）：[parent, Frame×5, Unframe×5, Wood×5, Gold×5]
VARIANT_LABELS_21 = [
    None,  # parent
    "Frame-style", "Frame-style", "Frame-style", "Frame-style", "Frame-style",
    "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style", "Unframe-style",
    "Vintage Wood Grain Frame-style", "Vintage Wood Grain Frame-style",
    "Vintage Wood Grain Frame-style", "Vintage Wood Grain Frame-style", "Vintage Wood Grain Frame-style",
    "Vintage Ornate Gold Frame-style", "Vintage Ornate Gold Frame-style",
    "Vintage Ornate Gold Frame-style", "Vintage Ornate Gold Frame-style", "Vintage Ornate Gold Frame-style",
]

# Size 名 (用于 Product Name 拼接, 与 SIZES_32/SIZES_SQUARE 一致, 但展开为列表)
SIZES_32_LIST = [
    "08x12inch(20x30cm)",
    "12x18inch(30x45cm)",
    "16x24inch(40x60cm)",
    "20x30inch(50x75cm)",
    "24x36inch(60x90cm)",
]
SIZES_SQUARE_LIST = [
    "12x12inch(30x30cm)",
    "16x16inch(40x40cm)",
    "20x20inch(50x20cm)",
    "24x24inch(60x60cm)",
    "28x28inch(70x70cm)",
]


def collapse_spaces(text: str) -> str:
    """多空格合并为单空格，去首尾空白。"""
    return re.sub(r"\s{2,}", " ", text).strip()


def extract_base_title(name: str) -> str:
    """从 Product Name 中提取基础标题（去掉 Frame-/Unframe- 及之后的内容）。"""
    match = FRAME_SPLIT_PATTERN.search(name)
    if match:
        return name[:match.start()].strip()
    return name.strip()


def remove_numeric_suffix(text: str) -> str:
    """删除 -N 数字后缀（如 -1, -2）。"""
    return NUMERIC_SUFFIX_PATTERN.sub("", text)


def replace_hyphens(text: str) -> str:
    """连字符替换为空格，但保留 Frame-style 和 Unframe-style。"""
    # 必须先替换 Unframe-style，否则 Frame-style 会吃掉它的子串
    text = text.replace("Unframe-style", "UNFRAME__STYLE__")
    text = text.replace("Frame-style", "FRAME__STYLE__")
    text = text.replace("-", " ")
    text = text.replace("UNFRAME__STYLE__", "Unframe-style")
    text = text.replace("FRAME__STYLE__", "Frame-style")
    return text


def replace_underscores(text: str) -> str:
    """下划线替换为空格。"""
    return text.replace("_", " ")


def deduplicate_words(text: str) -> str:
    """单词去重：超过 2 次出现的删除第三次及之后。

    case-insensitive 比较，保留原始大小写。
    先去除单词两端的标点符号再比较。
    """
    import string
    words = text.split(" ")
    counts: dict[str, int] = {}
    result = []

    for word in words:
        if not word:
            result.append(word)
            continue
        key = word.lower().strip(string.punctuation)
        counts[key] = counts.get(key, 0) + 1
        if counts[key] <= 2:
            result.append(word)

    return " ".join(result)


def normalize_group(
    ws: Worksheet,
    rows: list[int],
    col_idx: int,
    ratio_type: str,
) -> None:
    """对一个 11 行产品组执行 Item Name 规范化。

    按固定位置直接构造：{标题} {Frame/Unframe}-style {尺寸}
    顺序固定：第1行parent，第2-6行Frame+尺寸，第7-11行Unframe+尺寸。
    """
    sizes = SIZES_SQUARE if ratio_type == "square" else SIZES_32

    base_title = _extract_base_from_rows(ws, rows, col_idx)
    if base_title is None:
        return

    for i, row in enumerate(rows):
        cell = ws.cell(row=row, column=col_idx)
        if i == 0:
            name = base_title
        else:
            label = VARIANT_LABELS[i]
            size_idx = (i - 1) % 5  # 0-4 循环
            size = sizes[size_idx]
            name = f"{base_title} {label} {size}"
        # 新格式: 基名保留原样 (不去连字符/标点), 只合并多余空格
        cell.value = collapse_spaces(name)


def normalize_group_21(
    ws: Worksheet,
    rows: list[int],
    col_idx: int,
    ratio_type: str = "3:2",
) -> None:
    """对一个 21 行产品组（合并模式输出）执行 Item Name 规范化。

    顺序固定：第1行parent，第2-6行Frame×5，第7-11行Unframe×5，
    第12-16行Vintage Wood Grain×5，第17-21行Vintage Ornate Gold×5。
    """
    sizes = SIZES_SQUARE if ratio_type == "square" else SIZES_32

    base_title = _extract_base_from_rows(ws, rows, col_idx)
    if base_title is None:
        return

    for i, row in enumerate(rows):
        cell = ws.cell(row=row, column=col_idx)
        if i == 0:
            name = base_title
        else:
            label = VARIANT_LABELS_21[i]
            size_idx = (i - 1) % 5
            size = sizes[size_idx]
            name = f"{base_title} {label} {size}"
        cell.value = _clean_name(name)


def normalize_variant_group(
    ws: Worksheet,
    rows: list[int],
    col_idx: int,
    style_label: str,
    ratio_type: str = "3:2",
) -> None:
    """对木/金文件的一个 6 行产品组执行 Item Name 规范化。

    只给该组的变体行（rows[1:]）贴指定的 style_label + 尺寸后缀。
    第 1 行是 parent（不贴 style）。
    """
    sizes = SIZES_SQUARE if ratio_type == "square" else SIZES_32

    base_title = _extract_base_from_rows(ws, rows, col_idx)
    if base_title is None:
        return

    for i, row in enumerate(rows):
        cell = ws.cell(row=row, column=col_idx)
        if i == 0:
            name = base_title
        else:
            size_idx = (i - 1) % 5
            size = sizes[size_idx]
            name = f"{base_title} {style_label} {size}"
        cell.value = _clean_name(name)


def _extract_base_from_rows(ws, rows, col_idx):
    """从 group 的 parent 行提取基础标题。"""
    parent_cell = ws.cell(row=rows[0], column=col_idx)
    parent_value = parent_cell.value
    if parent_value is None:
        return None
    return extract_base_title(str(parent_value))


def _clean_name(name: str) -> str:
    """Item Name 清理管道。"""
    name = collapse_spaces(name)
    name = remove_numeric_suffix(name)
    name = replace_hyphens(name)
    name = replace_underscores(name)
    name = deduplicate_words(name)
    name = collapse_spaces(name)
    return name
