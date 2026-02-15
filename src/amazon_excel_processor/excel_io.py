"""Excel 文件读写模块"""

import logging
import os
import shutil
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook as _load_wb
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)

REQUIRED_COLUMNS = ["Product Name"]

OPTIONAL_COLUMNS = [
    "Variation Theme",
    "Paint Type",
    "Color Map",
    "Color",
    "Size",
    "Size Map",
    "Length",
    "Weight",
    "Search Terms",
]

HEADER_ROW = 2
DATA_START_ROW = 4
GROUP_SIZE = 11


def load_workbook(filepath: str | Path):
    """读取 Excel 文件，只保留 Template sheet 以加速处理。

    支持 xlsx 和 xlsm 格式。xlsm 使用 keep_vba=True 保留宏。
    返回 (workbook, worksheet, template_sheet_name)。
    """
    filepath = Path(filepath)

    if filepath.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(f"不支持的文件格式: {filepath.suffix}，仅支持 .xlsx 和 .xlsm")

    keep_vba = filepath.suffix.lower() == ".xlsm"
    wb = _load_wb(str(filepath), keep_vba=keep_vba)

    # 大小写不敏感匹配 template sheet
    sheet_name = None
    for name in wb.sheetnames:
        if name.lower() == "template":
            sheet_name = name
            break

    if sheet_name is None:
        available = ", ".join(wb.sheetnames)
        raise ValueError(f"找不到 'template' sheet。可用的 sheet: {available}")

    # 删除非 Template 的 sheet 以加速处理
    for name in wb.sheetnames:
        if name != sheet_name:
            del wb[name]

    ws = wb[sheet_name]
    return wb, ws, sheet_name


def locate_columns(ws: Worksheet, header_row: int = HEADER_ROW) -> dict[str, int]:
    """扫描表头行动态定位列索引。

    亚马逊模板中 Row 1 是元数据，Row 2 是列名，Row 3 是内部字段名，Row 4+ 是数据。
    返回 {列名: 列号(1-based)} 的映射。
    """
    col_map: dict[str, int] = {}

    for col_idx in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=header_row, column=col_idx).value
        if cell_value is None:
            continue
        header = str(cell_value).strip()
        all_columns = REQUIRED_COLUMNS + OPTIONAL_COLUMNS
        for expected in all_columns:
            if header.lower() == expected.lower():
                col_map[expected] = col_idx
                break

    for req in REQUIRED_COLUMNS:
        if req not in col_map:
            raise ValueError(f"必需列 '{req}' 在表头中未找到")

    found_optional = [c for c in OPTIONAL_COLUMNS if c in col_map]
    missing_optional = [c for c in OPTIONAL_COLUMNS if c not in col_map]
    if missing_optional:
        logger.info("可选列未找到（将跳过）: %s", ", ".join(missing_optional))
    if found_optional:
        logger.info("已定位列: %s", ", ".join([*REQUIRED_COLUMNS, *found_optional]))

    return col_map


def group_rows(ws: Worksheet) -> list[list[int]]:
    """将数据行按 11 行一组分组。

    返回 [[row_num, ...], ...] 列表，每组 11 个行号。
    不完整尾部组记录警告并跳过。
    """
    data_rows = list(range(DATA_START_ROW, ws.max_row + 1))

    if not data_rows:
        logger.warning("template sheet 没有数据行")
        return []

    total = len(data_rows)
    complete_groups = total // GROUP_SIZE
    remainder = total % GROUP_SIZE

    if remainder > 0:
        logger.warning(
            "数据行数 %d 不是 %d 的倍数，尾部 %d 行将被跳过",
            total, GROUP_SIZE, remainder,
        )

    groups = []
    for i in range(complete_groups):
        start = i * GROUP_SIZE
        group = data_rows[start: start + GROUP_SIZE]
        groups.append(group)

    return groups


def _can_write(path: Path) -> bool:
    """检测文件是否可写入（未被其他进程锁定）。

    对于只读权限的文件，先尝试删除再重建（跨平台安全）。
    对于被其他进程锁定的文件（如 Windows 上 Excel 打开），返回 False。
    """
    if not path.exists():
        return True

    # 尝试删除旧文件（copyfile 会重建）
    try:
        os.remove(str(path))
        return True
    except PermissionError:
        # Windows: 文件被其他进程锁定，无法删除
        return False
    except OSError:
        return False


def _resolve_output_path(input_path: Path, output_path: Optional[Path]) -> Path:
    """确定输出文件路径，若目标被占用则自动加序号。"""
    if output_path is None:
        base = input_path.parent / f"{input_path.stem}_processed{input_path.suffix}"
    else:
        base = output_path

    if _can_write(base):
        return base

    # 被占用，自动加序号
    for i in range(2, 100):
        candidate = base.parent / f"{base.stem}_{i}{base.suffix}"
        if _can_write(candidate):
            logger.warning("输出文件 %s 被占用，改用 %s", base.name, candidate.name)
            return candidate

    raise PermissionError(f"无法写入输出文件，请关闭占用 {base.name} 的程序后重试")


def save_workbook(
    processed_ws: Worksheet,
    input_path: str | Path,
    template_sheet_name: str,
    output_path: Optional[str | Path] = None,
) -> Path:
    """将处理后的 Template 写回原文件副本，保留所有其他 sheet。

    1. 复制原文件到输出路径（只复制内容，不复制权限）
    2. 打开副本，将处理后的单元格值写入 Template sheet
    3. 保存副本

    如果输出文件被其他程序占用，自动在文件名后加序号。
    """
    input_path = Path(input_path)
    out = _resolve_output_path(
        input_path,
        Path(output_path) if output_path is not None else None,
    )

    # 只复制文件内容，不复制权限元数据
    # 这样即使源文件是只读的（如浏览器下载），新文件也是可写的
    shutil.copyfile(str(input_path), str(out))

    # 打开副本，只更新 Template sheet 中被修改的列
    keep_vba = out.suffix.lower() == ".xlsm"
    out_wb = _load_wb(str(out), keep_vba=keep_vba)
    out_ws = out_wb[template_sheet_name]

    # 逐格复制处理后的数据（从 DATA_START_ROW 开始）
    for row in range(DATA_START_ROW, processed_ws.max_row + 1):
        for col in range(1, processed_ws.max_column + 1):
            src_val = processed_ws.cell(row=row, column=col).value
            out_ws.cell(row=row, column=col).value = src_val

    out_wb.save(str(out))
    return out
