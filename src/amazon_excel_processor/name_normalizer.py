"""Product Name 规范化模块"""

import logging
import re
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

FRAME_SPLIT_PATTERN = re.compile(r"\s+(Frame|Unframe)-", re.IGNORECASE)
NUMERIC_SUFFIX_PATTERN = re.compile(r"-(\d+)(?=\s|$)")

# 固定尺寸顺序（与 md 文档一致）—— {L}"L x {W}"W 格式
SIZES_32 = [
    '12"L x 8"W',
    '18"L x 12"W',
    '24"L x 16"W',
    '30"L x 20"W',
    '36"L x 24"W',
]
SIZES_SQUARE = [
    '12"L x 12"W',
    '16"L x 16"W',
    '20"L x 20"W',
    '24"L x 24"W',
    '28"L x 28"W',
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

    # 基名清理: 下划线→空格 + 多余空格合并
    base_title = replace_underscores(base_title)
    base_title = collapse_spaces(base_title)

    for i, row in enumerate(rows):
        cell = ws.cell(row=row, column=col_idx)
        if i == 0:
            name = base_title
        else:
            label = VARIANT_LABELS[i]
            size_idx = (i - 1) % 5  # 0-4 循环
            size = sizes[size_idx]
            name = f"{base_title} {label} {size}"
        cell.value = name


def _extract_base_from_rows(ws, rows, col_idx):
    """从 group 的 parent 行提取基础标题。"""
    parent_cell = ws.cell(row=rows[0], column=col_idx)
    parent_value = parent_cell.value
    if parent_value is None:
        return None
    return extract_base_title(str(parent_value))
