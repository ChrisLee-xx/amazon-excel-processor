"""变体字段填充模块测试"""

import pytest
from openpyxl import Workbook

from amazon_excel_processor.field_filler import (
    detect_ratio_type,
    fill_color,
    fill_size,
    fill_size_map,
    fill_length,
    fill_width,
    fill_price,
    fill_weight,
    fill_simple_fields,
    COLOR_SEQUENCE,
    SIZE_32,
    SIZE_SQUARE,
    SIZE_MAP_SEQUENCE,
    LENGTH_32,
    LENGTH_SQUARE,
    WIDTH_32,
    WIDTH_SQUARE,
    WEIGHT_SEQUENCE,
    PRICE_SEQUENCE,
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
    ws.cell(row=1, column=6).value = "Width"
    ws.cell(row=1, column=7).value = "Weight"
    ws.cell(row=1, column=8).value = "Variation Theme"
    ws.cell(row=1, column=9).value = "Paint Type"
    ws.cell(row=1, column=10).value = "Color Map"
    ws.cell(row=1, column=11).value = "Your Price"

    rows = []
    for i, name in enumerate(product_names):
        row = i + 2
        ws.cell(row=row, column=1).value = name
        rows.append(row)

    col_map = {
        "Product Name": 1, "Color": 2, "Size": 3,
        "Size Map": 4, "Length": 5, "Width": 6,
        "Weight": 7, "Variation Theme": 8, "Paint Type": 9, "Color Map": 10,
        "Your Price": 11,
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
    def test_32_ratio_size_empty(self):
        """Size 列为空 → 3:2"""
        ws, rows, col_map = _create_test_ws(_make_32_names())
        assert detect_ratio_type(ws, rows, col_map) == "3:2"

    def test_32_ratio_size_prefilled_unequal(self):
        """Size 列预填值 L!=W（如 12L''x08W''）→ 3:2"""
        ws, rows, col_map = _create_test_ws(_make_32_names())
        size_col = col_map["Size"]
        for i, row in enumerate(rows):
            if i == 0:
                continue
            ws.cell(row=row, column=size_col).value = SIZE_32[i]
        assert detect_ratio_type(ws, rows, col_map) == "3:2"

    def test_square_ratio_size_prefilled(self):
        """Size 列预填 L==W（如 12L''x12W''）→ square"""
        ws, rows, col_map = _create_test_ws(_make_square_names())
        size_col = col_map["Size"]
        for i, row in enumerate(rows):
            if i == 0:
                continue
            ws.cell(row=row, column=size_col).value = SIZE_SQUARE[i]
        assert detect_ratio_type(ws, rows, col_map) == "square"


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

    def test_square_size_keeps_prefilled(self):
        """正方形：fill_size 不覆盖用户预填值"""
        ws, rows, col_map = _create_test_ws(_make_square_names())
        size_col = col_map["Size"]
        prefilled = [""] + [f"PRE{i}" for i in range(10)]
        for i, row in enumerate(rows):
            ws.cell(row=row, column=size_col).value = prefilled[i]
        fill_size(ws, rows, col_map, "square")
        values = [ws.cell(row=r, column=size_col).value for r in rows]
        assert values == prefilled


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


class TestFillWidth:
    def test_32_width(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_width(ws, rows, col_map, "3:2")
        values = [ws.cell(row=r, column=col_map["Width"]).value for r in rows]
        assert values == WIDTH_32

    def test_square_width(self):
        ws, rows, col_map = _create_test_ws(_make_square_names())
        fill_width(ws, rows, col_map, "square")
        values = [ws.cell(row=r, column=col_map["Width"]).value for r in rows]
        assert values == WIDTH_SQUARE


class TestFillWeight:
    def test_weight_sequence(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_weight(ws, rows, col_map)
        values = [ws.cell(row=r, column=col_map["Weight"]).value for r in rows]
        assert values == WEIGHT_SEQUENCE


class TestFillPrice:
    def test_price_sequence(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_price(ws, rows, col_map)
        values = [ws.cell(row=r, column=col_map["Your Price"]).value for r in rows]
        assert values == PRICE_SEQUENCE


class TestFillSimpleFields:
    def test_simple_fields(self):
        ws, rows, col_map = _create_test_ws(_make_32_names())
        fill_simple_fields(ws, rows, col_map)
        for r in rows:
            assert ws.cell(row=r, column=col_map["Variation Theme"]).value == "color-size"
            assert ws.cell(row=r, column=col_map["Paint Type"]).value == "Oil"
            assert ws.cell(row=r, column=col_map["Color Map"]).value == "Multi"
