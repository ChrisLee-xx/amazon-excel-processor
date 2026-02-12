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
    """
    words = text.split(" ")
    counts: dict[str, int] = {}
    result = []

    for word in words:
        if not word:
            result.append(word)
            continue
        key = word.lower()
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
    """对一个 11 行产品组执行 Product Name 规范化。

    按固定位置直接构造：{标题} {Frame/Unframe}-style {尺寸}
    顺序固定：第1行parent，第2-6行Frame+尺寸，第7-11行Unframe+尺寸。
    """
    sizes = SIZES_SQUARE if ratio_type == "square" else SIZES_32

    # 先从 parent 行（第一行）提取基础标题
    parent_cell = ws.cell(row=rows[0], column=col_idx)
    parent_value = parent_cell.value
    if parent_value is None:
        return

    base_title = extract_base_title(str(parent_value))

    for i, row in enumerate(rows):
        cell = ws.cell(row=row, column=col_idx)
        value = cell.value
        if value is None:
            continue

        if i == 0:
            # parent 行：只做清理，不加 Frame/Unframe
            name = base_title
        else:
            # 变体行：按位置拼接
            label = VARIANT_LABELS[i]
            size_idx = (i - 1) % 5  # 0-4 循环
            size = sizes[size_idx]
            name = f"{base_title} {label} {size}"

        # 清理管道
        name = collapse_spaces(name)
        name = remove_numeric_suffix(name)
        name = replace_hyphens(name)
        name = replace_underscores(name)
        name = deduplicate_words(name)
        name = collapse_spaces(name)

        cell.value = name
