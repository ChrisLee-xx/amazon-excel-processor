"""合并模块测试 (3 文件 API: 主+木+金)"""

import random
import pytest
from openpyxl import Workbook, load_workbook

from amazon_excel_processor.excel_io import DATA_START_ROW, HEADER_ROW, group_rows
from amazon_excel_processor.name_normalizer import VARIANT_LABELS_21
from amazon_excel_processor.merger import (
    identify_file_role,
    identify_main_file,  # 向后兼容
    index_groups_by_name,
    merge_one_painting,
    rewrite_sku,
    write_parent_sku_formulas,
    fill_list_price_synced,
    build_sku_prefix,
    MAIN_GROUP_SIZE,
    VARIANT_GROUP_SIZE,
)


# ===== 测试夹具 =====

def _create_main_workbook(paintings):
    """模拟"普文件": N 画各 11 行 (1 parent + 5 Frame + 5 Unframe). 返回 (wb, col_map)."""
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


def _create_variant_workbook(paintings, role="wood", shuffled=False):
    """模拟"木框/金框文件": N 画各 1 个 6 行 group.

    parent 行 Product Name 不带 style 关键字 (与真实文件一致).
    children 用 {role} 后缀区分 (便于测试时验证来源).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    for c, h in {2: "Seller SKU", 9: "Product Name"}.items():
        ws.cell(row=HEADER_ROW, column=c).value = h
    items = list(paintings)
    if shuffled:
        random.seed(42)
        random.shuffle(items)
    row = DATA_START_ROW
    for title in items:
        ws.cell(row=row, column=9).value = title
        ws.cell(row=row, column=2).value = f"{role.upper()}-{row}"
        row += 1
        for i in range(5):
            ws.cell(row=row, column=9).value = f"{title} {role} size-{i}"
            ws.cell(row=row, column=2).value = f"{role.upper()}-{row}"
            row += 1
    return wb


# ===== identify_file_role =====

class TestIdentifyFileRole:
    def test_11_rows_is_main(self):
        wb, _ = _create_main_workbook(["A", "B"])
        groups = group_rows(wb.active, group_size=MAIN_GROUP_SIZE)
        role, _ = identify_file_role(groups)
        assert role == "main"

    def test_6_rows_is_variant(self):
        wb = _create_variant_workbook(["A", "B"], role="wood")
        groups = group_rows(wb.active, group_size=VARIANT_GROUP_SIZE)
        role, _ = identify_file_role(groups)
        assert role == "variant"

    def test_backward_compat_identify_main_file(self):
        """旧 API identify_main_file 仍可用"""
        wb, _ = _create_main_workbook(["A"])
        groups = group_rows(wb.active, group_size=MAIN_GROUP_SIZE)
        is_main, is_gold = identify_main_file(groups)
        assert is_main is True
        assert is_gold is False


# ===== index_groups_by_name =====

class TestIndexGroupsByName:
    def test_index_4_paintings(self):
        wb = _create_variant_workbook(["Art A", "Art B", "Art C", "Art D"], role="wood")
        ws = wb.active
        groups = group_rows(ws, group_size=VARIANT_GROUP_SIZE)
        by_name = index_groups_by_name(ws, groups)
        assert len(by_name) == 4

    def test_duplicate_raises(self):
        """同 base name 出现 2 次应报错"""
        wb = _create_variant_workbook(["Art A", "Art A"], role="wood")
        ws = wb.active
        groups = group_rows(ws, group_size=VARIANT_GROUP_SIZE)
        with pytest.raises(ValueError, match="出现多次"):
            index_groups_by_name(ws, groups)


# ===== merge_one_painting =====

def _setup_merge_one_painting(painting="Art A"):
    """构造单画的 main + wood + gold 测试数据, 返回所需参数."""
    main_wb, main_col_map = _create_main_workbook([painting])
    main_ws = main_wb.active
    main_groups = group_rows(main_ws, group_size=MAIN_GROUP_SIZE)

    wood_ws = _create_variant_workbook([painting], role="wood").active
    wood_groups = group_rows(wood_ws, group_size=VARIANT_GROUP_SIZE)

    gold_ws = _create_variant_workbook([painting], role="gold").active
    gold_groups = group_rows(gold_ws, group_size=VARIANT_GROUP_SIZE)

    max_col = max(main_ws.max_column, wood_ws.max_column, gold_ws.max_column)
    main_snapshots = [
        {c: main_ws.cell(row=r, column=c).value for c in range(1, max_col + 1)}
        for r in main_groups[0]
    ]
    return {
        "main_ws": main_ws,
        "main_col_map": main_col_map,
        "main_snapshots": main_snapshots,
        "wood_group": wood_groups[0],
        "gold_group": gold_groups[0],
        "wood_ws": wood_ws,
        "gold_ws": gold_ws,
        "max_col": max_col,
    }


class TestMergeOnePainting:
    def test_output_21_rows(self):
        s = _setup_merge_one_painting()
        merged = merge_one_painting(
            main_snapshots=s["main_snapshots"],
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
        )
        assert len(merged) == 21
        assert merged[0] == 4
        assert merged[-1] == 24

    def test_output_color_sequence(self):
        s = _setup_merge_one_painting()
        merged = merge_one_painting(
            main_snapshots=s["main_snapshots"],
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
        )
        colors = [s["main_ws"].cell(row=r, column=38).value for r in merged]
        # merged[0]=r4 parent, [1-5]=Frame, [6-10]=Unframe, [11-15]=Wood, [16-20]=Gold
        assert colors[0] in (None, "")
        assert colors[1] == "Frame-style"
        assert colors[5] == "Frame-style"
        assert colors[6] == "Unframe-style"
        assert colors[10] == "Unframe-style"
        assert colors[11] == "Vintage Wood Grain Frame-style"
        assert colors[15] == "Vintage Wood Grain Frame-style"
        assert colors[16] == "Vintage Ornate Gold Frame-style"
        assert colors[20] == "Vintage Ornate Gold Frame-style"

    def test_wood_gold_match_frame_size(self):
        s = _setup_merge_one_painting()
        merged = merge_one_painting(
            main_snapshots=s["main_snapshots"],
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
        )
        # Frame r5-r9, Wood r15-r19, Gold r20-r24
        for rf, rw, rg in [(5, 15, 20), (6, 16, 21), (7, 17, 22), (8, 18, 23), (9, 19, 24)]:
            ws = s["main_ws"]
            assert ws.cell(row=rf, column=41).value == ws.cell(row=rw, column=41).value
            assert ws.cell(row=rf, column=41).value == ws.cell(row=rg, column=41).value
            assert ws.cell(row=rf, column=55).value == ws.cell(row=rw, column=55).value
            assert ws.cell(row=rf, column=62).value == ws.cell(row=rw, column=62).value
            assert ws.cell(row=rf, column=63).value == ws.cell(row=rw, column=63).value
            assert ws.cell(row=rf, column=69).value == ws.cell(row=rw, column=69).value
            assert ws.cell(row=rf, column=69).value == ws.cell(row=rg, column=69).value

    def test_parentage_relationship_type(self):
        s = _setup_merge_one_painting()
        merge_one_painting(
            main_snapshots=s["main_snapshots"],
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
        )
        ws = s["main_ws"]
        assert ws.cell(row=4, column=30).value == "Parent"
        assert ws.cell(row=5, column=30).value == "Child"
        assert ws.cell(row=24, column=30).value == "Child"
        assert ws.cell(row=4, column=24).value in (None, "")
        assert ws.cell(row=5, column=24).value == "Variation"

    def test_wood_image_from_wood_file(self):
        """Wood 行的 Image URL 必须来自 wood 文件, Gold 行的来自 gold 文件.

        这是 3 文件方案的核心: 用户在 GUI 中明确指定 wood/gold, 不再依赖出现顺序.
        """
        s = _setup_merge_one_painting()
        merge_one_painting(
            main_snapshots=s["main_snapshots"],
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
        )
        ws = s["main_ws"]
        # Wood r15 (第 1 个 wood child) 的 SKU 应来自 wood 文件
        # 测试夹具里 wood SKU 是 "WOOD-{row}", gold SKU 是 "GOLD-{row}"
        # 但 merge_one_painting 后 SKU 没被重写 (rewrite_sku 是单独步骤)
        # 所以 r15 col2 应该是 wood 文件第 1 个 child 的 SKU
        wood_child_sku = ws.cell(row=15, column=2).value
        gold_child_sku = ws.cell(row=20, column=2).value
        # wood 文件第 1 个 group 的第 1 个 child 是 r5 (parent r4, child r5-r9)
        assert wood_child_sku == "WOOD-5"
        assert gold_child_sku == "GOLD-5"


# ===== rewrite_sku =====

class TestRewriteSku:
    def test_continuous_numbering_new_mode(self):
        """新品上架: 全部 21 行 × N 画从 prefix-1 连续编号"""
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25)), list(range(25, 46))]
        rewrite_sku(ws, groups, prefix="HM725", mode="new")
        assert ws.cell(row=4, column=2).value == "HM725-1"
        assert ws.cell(row=5, column=2).value == "HM725-2"
        assert ws.cell(row=24, column=2).value == "HM725-21"
        assert ws.cell(row=25, column=2).value == "HM725-22"
        assert ws.cell(row=45, column=2).value == "HM725-42"

    def test_single_group_new_mode(self):
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25))]
        rewrite_sku(ws, groups, prefix="AB", mode="new")
        assert ws.cell(row=4, column=2).value == "AB-1"
        assert ws.cell(row=24, column=2).value == "AB-21"

    def test_old_variant_mode_preserves_main_sku(self):
        """老品补充变体: group[0:11] (普文件原 11 行) SKU 保留, group[11:21] 重写"""
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25))]
        # 预设普文件原 11 行的 SKU (group[0:11] = r4-r14)
        for i, r in enumerate(groups[0][:11]):
            ws.cell(row=r, column=2).value = f"OLD-{i+1}"
        rewrite_sku(ws, groups, prefix="NEW", mode="old_variant")
        # r4-r14 保留原 SKU
        assert ws.cell(row=4, column=2).value == "OLD-1"
        assert ws.cell(row=14, column=2).value == "OLD-11"
        # r15-r24 (Wood+Gold) 用新前缀从 1 开始
        assert ws.cell(row=15, column=2).value == "NEW-1"
        assert ws.cell(row=16, column=2).value == "NEW-2"
        assert ws.cell(row=24, column=2).value == "NEW-10"

    def test_old_variant_mode_multi_groups_continuous(self):
        """老品补充变体: 多 group 时 Wood/Gold 跨 group 连续编号"""
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25)), list(range(25, 46))]
        rewrite_sku(ws, groups, prefix="NEW", mode="old_variant")
        # group 1: r15-r24 = NEW-1 到 NEW-10
        assert ws.cell(row=15, column=2).value == "NEW-1"
        assert ws.cell(row=24, column=2).value == "NEW-10"
        # group 2: r36-r45 = NEW-11 到 NEW-20
        assert ws.cell(row=36, column=2).value == "NEW-11"
        assert ws.cell(row=45, column=2).value == "NEW-20"


# ===== write_parent_sku_formulas =====

class TestWriteParentSkuFormulas:
    def test_first_child_uses_b_parent_new_mode(self):
        """新品上架: parent 清空, 第 1 child =B4, 后续 =AA{prev}"""
        wb = Workbook()
        ws = wb.active
        rows = list(range(4, 25))
        write_parent_sku_formulas(ws, [rows], mode="new")
        assert ws.cell(row=4, column=27).value in (None, "")
        assert ws.cell(row=5, column=27).value == "=B4"
        assert ws.cell(row=6, column=27).value == "=AA5"
        assert ws.cell(row=7, column=27).value == "=AA6"
        assert ws.cell(row=24, column=27).value == "=AA23"

    def test_multiple_groups_new_mode(self):
        wb = Workbook()
        ws = wb.active
        write_parent_sku_formulas(ws, [list(range(4, 25)), list(range(25, 46))], mode="new")
        assert ws.cell(row=5, column=27).value == "=B4"
        assert ws.cell(row=6, column=27).value == "=AA5"
        assert ws.cell(row=26, column=27).value == "=B25"
        assert ws.cell(row=27, column=27).value == "=AA26"

    def test_old_variant_mode_preserves_main_formulas(self):
        """老品补充变体: group[0:11] parent SKU 公式保留, group[11:21] 从 =AA{group[10]} 开始"""
        wb = Workbook()
        ws = wb.active
        rows = list(range(4, 25))
        # 预设普文件原 11 行的 parent SKU 公式
        ws.cell(row=4, column=27).value = None       # parent
        ws.cell(row=5, column=27).value = "=B4"      # Frame 1
        ws.cell(row=6, column=27).value = "=AA5"
        ws.cell(row=14, column=27).value = "=AA13"   # Unframe 5 (最后一个)
        write_parent_sku_formulas(ws, [rows], mode="old_variant")
        # r4-r14 保留
        assert ws.cell(row=4, column=27).value is None
        assert ws.cell(row=5, column=27).value == "=B4"
        assert ws.cell(row=14, column=27).value == "=AA13"
        # r15 (Wood 1) = =AA14 (引用 r14 Unframe 最后一个)
        assert ws.cell(row=15, column=27).value == "=AA14"
        # r16 = =AA15
        assert ws.cell(row=16, column=27).value == "=AA15"
        # r24 (Gold 5) = =AA23
        assert ws.cell(row=24, column=27).value == "=AA23"


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
        fill_list_price_synced(ws, rows, col_map={"Your Price": 13})
        assert ws.cell(row=5, column=13).value == 19.9


# ===== build_sku_prefix =====

class TestBuildSkuPrefix:
    def test_single_string(self):
        assert build_sku_prefix("HM725") == "HM725"

    def test_strip_whitespace(self):
        assert build_sku_prefix("  HM725  ") == "HM725"

    def test_complex_prefix(self):
        assert build_sku_prefix("AB2026风景") == "AB2026风景"
