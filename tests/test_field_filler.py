"""变体字段填充模块测试"""

import pytest
from openpyxl import Workbook

from amazon_excel_processor.field_filler import (
    detect_ratio_type,
    fill_color,
    fill_size,
    fill_size_map,
    fill_length,
    fill_weight,
    fill_simple_fields,
    COLOR_SEQUENCE,
    SIZE_32,
    SIZE_SQUARE,
    SIZE_MAP_SEQUENCE,
    LENGTH_32,
    LENGTH_SQUARE,
    WEIGHT_SEQUENCE,
)


def _create_test_ws(product_names: list[str]) -> tuple:
    """创建测试用 worksheet，返回 (ws, rows, col_map)。"""
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1).value = "Product Name"
    ws.cell(row=1, column=2).value = "Color"
    ws.cell(row=1, column=3).value = "Size"
    ws.cell(row=1, column=4).value = "Size Map"
    ws.cell(row=1, column=5).value = "Length"
    ws.cell(row=1, column=6).value = "Weight"
    ws.cell(row=1, column=7).value = "Variation Theme"
    ws.cell(row=1, column=8).value = "Paint Type"
    ws.cell(row=1, column=9).value = "Color Map"

    rows = []
    for i, name in enumerate(product_names):
        row = i + 2
        ws.cell(row=row, column=1).value = name
        rows.append(row)

    col_map = {
        "Product Name": 1, "Color": 2, "Size": 3,
        "Size Map": 4, "Length": 5, "Weight": 6,
        "Variation Theme": 7, "Paint Type": 8, "Color Map": 9,
    }
    return ws, rows, col_map


def _make_32_names():
    """生成 3:2 比例的 11 行产品名称。"""
    return [
        "Parent Title",
        "Title Frame-style 08x12inch(20x30cm)",
        "Title Frame-style 12x18inch(30x45cm)",
        "Title Frame-style 16x24inch(40x60cm)",
        "Title Frame-style 20x30inch(50x75cm)",
        "Title Frame-style 24x36inch(60x90cm)",
        "Title Unframe-style 08x12inch(20x30cm)",
        "Title Unframe-style 12x18inch(30x45cm)",
        "Title Unframe-style 16x24inch(40x60cm)",
        "Title Unframe-style 20x30inch(50x75cm)",
        "Title Unframe-style 24x36inch(60x90cm)",
    ]


def _make_square_names():
    """生成正方形比例的 11 行产品名称。"""
    return [
        "Parent Title",
        "Title Frame-style 12x12inch(30x30cm)",
        "Title Frame-style 16x16inch(40x40cm)",
        "Title Frame-style 20x20inch(50x50cm)",
        "Title Frame-style 24x24inch(60x60cm)",
        "Title Frame-style 28x28inch(70x70cm)",
        "Title Unframe-style 12x12inch(30x30cm)",
        "Title Unframe-style 16x16inch(40x40cm)",
        "Title Unframe-style 20x20inch(50x50cm)",
        "Title Unframe-style 24x24inch(60x60cm)",
        "Title Unframe-style 28x28inch(70x70cm)",
    ]


class TestDetectRatioType:
    def test_32_ratio(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        assert detect_ratio_type(ws, rows, col_map["Product Name"]) == "3:2"

    def test_square_ratio(self):
        ws, rows, col_map = _create_test_ws(_make_square_names())
        assert detect_ratio_type(ws, rows, col_map["Product Name"]) == "square"


class TestFillColor:
    def test_color_sequence(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_color(ws, rows, col_map)
        values = [ws.cell(row=r, column=col_map["Color"]).value for r in rows]
        assert values == COLOR_SEQUENCE


class TestFillSize:
    def test_32_size(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_size(ws, rows, col_map, "3:2")
        values = [ws.cell(row=r, column=col_map["Size"]).value for r in rows]
        assert values == SIZE_32

    def test_square_size(self):
        ws, rows, col_map = _create_test_ws(_make_square_names())
        fill_size(ws, rows, col_map, "square")
        values = [ws.cell(row=r, column=col_map["Size"]).value for r in rows]
        assert values == SIZE_SQUARE


class TestFillSizeMap:
    def test_size_map_sequence(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_size_map(ws, rows, col_map)
        values = [ws.cell(row=r, column=col_map["Size Map"]).value for r in rows]
        assert values == SIZE_MAP_SEQUENCE


class TestFillLength:
    def test_32_length(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_length(ws, rows, col_map, "3:2")
        values = [ws.cell(row=r, column=col_map["Length"]).value for r in rows]
        assert values == LENGTH_32

    def test_square_length(self):
        ws, rows, col_map = _create_test_ws(_make_square_names())
        fill_length(ws, rows, col_map, "square")
        values = [ws.cell(row=r, column=col_map["Length"]).value for r in rows]
        assert values == LENGTH_SQUARE


class TestFillWeight:
    def test_weight_sequence(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_weight(ws, rows, col_map)
        values = [ws.cell(row=r, column=col_map["Weight"]).value for r in rows]
        assert values == WEIGHT_SEQUENCE


class TestFillSimpleFields:
    def test_simple_fields(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_simple_fields(ws, rows, col_map)
        for r in rows:
            assert ws.cell(row=r, column=col_map["Variation Theme"]).value == "color-size"
            assert ws.cell(row=r, column=col_map["Paint Type"]).value == "Oil"
            assert ws.cell(row=r, column=col_map["Color Map"]).value == "Multi"
