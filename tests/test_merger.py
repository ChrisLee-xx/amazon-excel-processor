"""合并模块测试 (TDD)"""

import random
import pytest
from openpyxl import Workbook, load_workbook

from amazon_excel_processor.excel_io import DATA_START_ROW, HEADER_ROW, group_rows
from amazon_excel_processor.name_normalizer import VARIANT_LABELS_21
from amazon_excel_processor.merger import (
    identify_main_file,
    pair_gold_groups,
    merge_one_painting,
    rewrite_sku,
    write_parent_sku_formulas,
    fill_list_price_synced,
    build_sku_prefix,
)


# ===== 测试夹具 =====

def _create_main_workbook(paintings):
    """模拟"普文件": N 画各 11 行. 返回 (wb, col_map)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    headers = {
        1: "Product Type", 2: "Seller SKU", 9: "Product Name",
        13: "Your Price", 24: "Relationship Type", 25: "Package Level",
        26: "Variation Theme", 27: "Parent SKU", 30: "Parentage",
        38: "Color", 41: "Size", 55: "Size Map", 62: "Length", 63: "Width",
        69: "Weight", 145: "List Price",
    }
    for c, h in headers.items():
        ws.cell(row=HEADER_ROW, column=c).value = h
    row = DATA_START_ROW
    for title in paintings:
        ws.cell(row=row, column=9).value = title
        ws.cell(row=row, column=2).value = f"SKU-{row}"
        ws.cell(row=row, column=30).value = "Parent"
        row += 1
        for i in range(10):
            ws.cell(row=row, column=9).value = (
                f"{title} Frame-style 12x18inch" if i < 5
                else f"{title} Unframe-style 12x18inch"
            )
            ws.cell(row=row, column=2).value = f"SKU-{row}"
            ws.cell(row=row, column=30).value = "Child"
            row += 1
    col_map = {h: c for c, h in headers.items()}
    return wb, col_map


def _create_gold_workbook(paintings, shuffled=False):
    """模拟"金文件": N 画各 2 个 6 行 group (Wood + Gold).

    parent 行 Product Name 不带 style 关键字 (实际金文件就是这样).
    区分 Wood/Gold 靠出现顺序.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    for c, h in {2: "Seller SKU", 9: "Product Name"}.items():
        ws.cell(row=HEADER_ROW, column=c).value = h
    pairs = []
    for p in paintings:
        pairs.append((p, "Wood"))
        pairs.append((p, "Gold"))
    if shuffled:
        random.seed(42)
        random.shuffle(pairs)
    row = DATA_START_ROW
    for title, style in pairs:
        # parent 行: 真实金文件不带 style 关键字, 只用 base name
        ws.cell(row=row, column=9).value = title
        ws.cell(row=row, column=2).value = f"GOLD-{row}"
        row += 1
        # children: 用 style 后缀区分
        for i in range(5):
            ws.cell(row=row, column=9).value = f"{title} {style} size-{i}"
            ws.cell(row=row, column=2).value = f"GOLD-{row}"
            row += 1
    return wb


# ===== identify_main_file =====

class TestIdentifyMainFile:
    def test_11_rows_is_main(self):
        wb, col_map = _create_main_workbook(["A", "B"])
        groups = group_rows(wb.active, group_size=11)
        is_main, is_gold = identify_main_file(groups)
        assert is_main is True
        assert is_gold is False

    def test_6_rows_is_gold(self):
        wb = _create_gold_workbook(["A", "B"])
        groups = group_rows(wb.active, group_size=6)
        is_main, is_gold = identify_main_file(groups)
        assert is_main is False
        assert is_gold is True


# ===== pair_gold_groups =====

class TestPairGoldGroups:
    def test_pairs_8_groups_into_4(self):
        wb = _create_gold_workbook(["Art A", "Art B", "Art C", "Art D"])
        ws = wb.active
        groups = group_rows(ws, group_size=6)
        pairs = pair_gold_groups(ws, groups)
        assert len(pairs) == 4
        for w, g in pairs:
            assert len(w) == 6 and len(g) == 6
            # 同 base name
            wn = ws.cell(row=w[0], column=9).value
            gn = ws.cell(row=g[0], column=9).value
            assert wn == gn

    def test_shuffled_still_pairs(self):
        wb = _create_gold_workbook(["A", "B", "C", "D"], shuffled=True)
        ws = wb.active
        groups = group_rows(ws, group_size=6)
        pairs = pair_gold_groups(ws, groups)
        assert len(pairs) == 4


# ===== merge_one_painting =====

class TestMergeOnePainting:
    def test_output_21_rows(self):
        main_wb, main_col_map = _create_main_workbook(["Art A"])
        main_ws = main_wb.active
        main_groups = group_rows(main_ws, group_size=11)
        gold_ws = _create_gold_workbook(["Art A"]).active
        gold_groups = group_rows(gold_ws, group_size=6)
        gold_pairs = pair_gold_groups(gold_ws, gold_groups)
        # 快照 main group
        max_col = max(main_ws.max_column, gold_ws.max_column)
        main_snapshots = [{c: main_ws.cell(row=r, column=c).value for c in range(1, max_col+1)} for r in main_groups[0]]
        merged = merge_one_painting(
            main_snapshots=main_snapshots,
            gold_pair=gold_pairs[0],
            output_start_row=4,
            output_ws=main_ws,
            col_map=main_col_map,
            gold_wood_ws=gold_ws,
            gold_gold_ws=gold_ws,
            max_col=max_col,
        )
        assert len(merged) == 21
        assert merged[0] == 4
        assert merged[-1] == 24

    def test_output_color_sequence(self):
        main_wb, main_col_map = _create_main_workbook(["Art A"])
        main_ws = main_wb.active
        main_groups = group_rows(main_ws, group_size=11)
        gold_ws = _create_gold_workbook(["Art A"]).active
        gold_groups = group_rows(gold_ws, group_size=6)
        gold_pairs = pair_gold_groups(gold_ws, gold_groups)
        # 快照 main group
        max_col = max(main_ws.max_column, gold_ws.max_column)
        main_snapshots = [{c: main_ws.cell(row=r, column=c).value for c in range(1, max_col+1)} for r in main_groups[0]]
        merged = merge_one_painting(
            main_snapshots=main_snapshots,
            gold_pair=gold_pairs[0],
            output_start_row=4,
            output_ws=main_ws,
            col_map=main_col_map,
            gold_wood_ws=gold_ws,
            gold_gold_ws=gold_ws,
            max_col=max_col,
        )
        colors = [main_ws.cell(row=r, column=38).value for r in merged]
        # merged[0]=r4 (parent), merged[1-5]=Frame×5, merged[6-10]=Unframe×5,
        # merged[11-15]=Wood×5, merged[16-20]=Gold×5
        assert colors[0] in (None, "")  # parent
        assert colors[1] == "Frame-style"
        assert colors[5] == "Frame-style"
        assert colors[6] == "Unframe-style"
        assert colors[10] == "Unframe-style"
        assert colors[11] == "Vintage Wood Grain Frame-style"
        assert colors[15] == "Vintage Wood Grain Frame-style"
        assert colors[16] == "Vintage Ornate Gold Frame-style"
        assert colors[20] == "Vintage Ornate Gold Frame-style"

    def test_wood_gold_match_frame_size(self):
        main_wb, main_col_map = _create_main_workbook(["Art A"])
        main_ws = main_wb.active
        main_groups = group_rows(main_ws, group_size=11)
        gold_ws = _create_gold_workbook(["Art A"]).active
        gold_groups = group_rows(gold_ws, group_size=6)
        gold_pairs = pair_gold_groups(gold_ws, gold_groups)
        # 快照 main group
        max_col = max(main_ws.max_column, gold_ws.max_column)
        main_snapshots = [{c: main_ws.cell(row=r, column=c).value for c in range(1, max_col+1)} for r in main_groups[0]]
        merged = merge_one_painting(
            main_snapshots=main_snapshots,
            gold_pair=gold_pairs[0],
            output_start_row=4,
            output_ws=main_ws,
            col_map=main_col_map,
            gold_wood_ws=gold_ws,
            gold_gold_ws=gold_ws,
            max_col=max_col,
        )
        # Frame r5-r9, Wood r15-r19, Gold r20-r24
        for rf, rw, rg in [(5, 15, 20), (6, 16, 21), (7, 17, 22), (8, 18, 23), (9, 19, 24)]:
            assert main_ws.cell(row=rf, column=41).value == main_ws.cell(row=rw, column=41).value
            assert main_ws.cell(row=rf, column=41).value == main_ws.cell(row=rg, column=41).value
            assert main_ws.cell(row=rf, column=55).value == main_ws.cell(row=rw, column=55).value
            assert main_ws.cell(row=rf, column=62).value == main_ws.cell(row=rw, column=62).value
            assert main_ws.cell(row=rf, column=63).value == main_ws.cell(row=rw, column=63).value
            assert main_ws.cell(row=rf, column=69).value == main_ws.cell(row=rw, column=69).value
            assert main_ws.cell(row=rf, column=69).value == main_ws.cell(row=rg, column=69).value

    def test_parentage_relationship_type(self):
        main_wb, main_col_map = _create_main_workbook(["Art A"])
        main_ws = main_wb.active
        main_groups = group_rows(main_ws, group_size=11)
        gold_ws = _create_gold_workbook(["Art A"]).active
        gold_groups = group_rows(gold_ws, group_size=6)
        gold_pairs = pair_gold_groups(gold_ws, gold_groups)
        # 快照 main group
        max_col = max(main_ws.max_column, gold_ws.max_column)
        main_snapshots = [{c: main_ws.cell(row=r, column=c).value for c in range(1, max_col+1)} for r in main_groups[0]]
        merged = merge_one_painting(
            main_snapshots=main_snapshots,
            gold_pair=gold_pairs[0],
            output_start_row=4,
            output_ws=main_ws,
            col_map=main_col_map,
            gold_wood_ws=gold_ws,
            gold_gold_ws=gold_ws,
            max_col=max_col,
        )
        assert main_ws.cell(row=4, column=30).value == "Parent"
        assert main_ws.cell(row=5, column=30).value == "Child"
        assert main_ws.cell(row=24, column=30).value == "Child"
        assert main_ws.cell(row=4, column=24).value in (None, "")
        assert main_ws.cell(row=5, column=24).value == "Variation"


# ===== rewrite_sku =====

class TestRewriteSku:
    def test_continuous_numbering(self):
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25)), list(range(25, 46))]
        rewrite_sku(ws, groups, prefix="HM725")
        assert ws.cell(row=4, column=2).value == "HM725-1"
        assert ws.cell(row=5, column=2).value == "HM725-2"
        assert ws.cell(row=24, column=2).value == "HM725-21"
        assert ws.cell(row=25, column=2).value == "HM725-22"
        assert ws.cell(row=45, column=2).value == "HM725-42"

    def test_single_group(self):
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25))]
        rewrite_sku(ws, groups, prefix="AB")
        assert ws.cell(row=4, column=2).value == "AB-1"
        assert ws.cell(row=24, column=2).value == "AB-21"


# ===== write_parent_sku_formulas =====

class TestWriteParentSkuFormulas:
    def test_first_child_uses_b_parent(self):
        wb = Workbook()
        ws = wb.active
        rows = list(range(4, 25))
        write_parent_sku_formulas(ws, [rows])
        assert ws.cell(row=4, column=27).value in (None, "")
        assert ws.cell(row=5, column=27).value == "=B4"
        assert ws.cell(row=6, column=27).value == "=AA5"
        assert ws.cell(row=7, column=27).value == "=AA6"
        assert ws.cell(row=24, column=27).value == "=AA23"

    def test_multiple_groups(self):
        wb = Workbook()
        ws = wb.active
        write_parent_sku_formulas(ws, [list(range(4, 25)), list(range(25, 46))])
        assert ws.cell(row=5, column=27).value == "=B4"
        assert ws.cell(row=6, column=27).value == "=AA5"
        assert ws.cell(row=26, column=27).value == "=B25"
        assert ws.cell(row=27, column=27).value == "=AA26"


# ===== fill_list_price_synced =====

class TestFillListPriceSynced:
    def test_list_price_matches_your_price(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=HEADER_ROW, column=13).value = "Your Price"
        ws.cell(row=HEADER_ROW, column=145).value = "List Price"
        rows = [4, 5, 6]
        ws.cell(row=4, column=13).value = None
        ws.cell(row=5, column=13).value = 19.9
        ws.cell(row=6, column=13).value = 29.9
        fill_list_price_synced(ws, rows, col_map={"Your Price": 13, "List Price": 145})
        assert ws.cell(row=4, column=145).value is None
        assert ws.cell(row=5, column=145).value == 19.9
        assert ws.cell(row=6, column=145).value == 29.9

    def test_missing_columns_no_op(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=HEADER_ROW, column=13).value = "Your Price"
        rows = [4, 5]
        ws.cell(row=5, column=13).value = 19.9
        # 无 List Price 列
        fill_list_price_synced(ws, rows, col_map={"Your Price": 13})
        assert ws.cell(row=5, column=13).value == 19.9  # 不影响


# ===== build_sku_prefix =====

class TestBuildSkuPrefix:
    def test_concat_three_parts(self):
        assert build_sku_prefix("HM", "725", "AB") == "HM725AB"

    def test_empty_theme(self):
        """主题缩写在 3 部分模式下允许为空字符串"""
        assert build_sku_prefix("HM", "725", "") == "HM725"

    def test_strip_whitespace(self):
        assert build_sku_prefix(" HM ", " 725 ", " AB ") == "HM725AB"
