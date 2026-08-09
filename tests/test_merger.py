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
    """模拟"普文件"(新格式): N 画各 11 行 (1 parent + 5 Frame + 5 Unframe). 返回 (wb, col_map)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    headers = {
        1: "SKU", 4: "Parentage Level", 5: "Parent SKU", 6: "Variation Theme Name",
        7: "Item Name", 46: "Style", 55: "Color", 56: "Size",
        124: "Item Length Longer Edge", 126: "Item Width Shorter Edge",
        147: "Item Weight", 154: "List Price",
    }
    for c, h in headers.items():
        ws.cell(row=HEADER_ROW, column=c).value = h
    row = DATA_START_ROW
    for title in paintings:
        ws.cell(row=row, column=7).value = title
        ws.cell(row=row, column=1).value = f"SKU-{row}"
        ws.cell(row=row, column=4).value = "Parent"
        row += 1
        for i in range(10):
            ws.cell(row=row, column=7).value = (
                f"{title} Frame-style 12x18inch" if i < 5
                else f"{title} Unframe-style 12x18inch"
            )
            ws.cell(row=row, column=1).value = f"SKU-{row}"
            ws.cell(row=row, column=4).value = "Child"
            row += 1
    col_map = {h: c for c, h in headers.items()}
    return wb, col_map


def _create_variant_workbook(paintings, role="wood", shuffled=False):
    """模拟"木框/金框文件"(新格式): N 画各 1 个 6 行 group.

    parent 行 Item Name 不带 style 关键字 (与真实文件一致).
    children 用 {role} 后缀区分 (便于测试时验证来源).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Template"
    for c, h in {1: "SKU", 4: "Parentage Level", 7: "Item Name"}.items():
        ws.cell(row=HEADER_ROW, column=c).value = h
    items = list(paintings)
    if shuffled:
        random.seed(42)
        random.shuffle(items)
    row = DATA_START_ROW
    for title in items:
        ws.cell(row=row, column=7).value = title
        ws.cell(row=row, column=1).value = f"{role.upper()}-{row}"
        ws.cell(row=row, column=4).value = "Parent"
        row += 1
        for i in range(5):
            ws.cell(row=row, column=7).value = f"{title} {role} size-{i}"
            ws.cell(row=row, column=1).value = f"{role.upper()}-{row}"
            ws.cell(row=row, column=4).value = "Child"
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

    def test_duplicate_allowed_by_order(self):
        """同 base name 出现多次允许, 按顺序配对 (普第 N 次 ↔ 木第 N 次)"""
        wb = _create_variant_workbook(["Art A", "Art A"], role="wood")
        ws = wb.active
        groups = group_rows(ws, group_size=VARIANT_GROUP_SIZE)
        by_name = index_groups_by_name(ws, groups)
        assert len(by_name) == 1
        assert len(by_name["art a"]) == 2  # 2 个 group, 按出现顺序


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
        colors = [s["main_ws"].cell(row=r, column=55).value for r in merged]
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
        # 注意: Color/Weight 不同 style 本来就不同, 不比较; 只比较 Size/Length/Width (物理尺寸一致)
        for rf, rw, rg in [(5, 15, 20), (6, 16, 21), (7, 17, 22), (8, 18, 23), (9, 19, 24)]:
            ws = s["main_ws"]
            assert ws.cell(row=rf, column=56).value == ws.cell(row=rw, column=56).value
            assert ws.cell(row=rf, column=56).value == ws.cell(row=rg, column=56).value
            assert ws.cell(row=rf, column=124).value == ws.cell(row=rw, column=124).value
            assert ws.cell(row=rf, column=124).value == ws.cell(row=rg, column=124).value
            assert ws.cell(row=rf, column=126).value == ws.cell(row=rw, column=126).value
            assert ws.cell(row=rf, column=126).value == ws.cell(row=rg, column=126).value

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
        assert ws.cell(row=4, column=4).value == "Parent"
        assert ws.cell(row=5, column=4).value == "Child"
        assert ws.cell(row=24, column=4).value == "Child"
        # 新格式: Variation Theme Name (col6) = "COLOR/SIZE"
        assert ws.cell(row=5, column=6).value == "COLOR/SIZE"
        assert ws.cell(row=24, column=6).value == "COLOR/SIZE"

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
        wood_child_sku = ws.cell(row=15, column=1).value
        gold_child_sku = ws.cell(row=20, column=1).value
        # wood 文件第 1 个 group 的第 1 个 child 是 r5 (parent r4, child r5-r9)
        assert wood_child_sku == "WOOD-9"
        assert gold_child_sku == "GOLD-9"

    def test_old_variant_mode_preserves_main_rows(self):
        """老品补充变体模式: 普文件原 11 行 (r4-r14) 完全不动, 仅 Wood/Gold 行处理"""
        s = _setup_merge_one_painting()
        main_ws = s["main_ws"]
        main_col_map = s["main_col_map"]
        main_snapshots = s["main_snapshots"]

        # 先记录普文件原 11 行的所有列值 (从快照)
        original_values = {}
        for i in range(11):
            for c, v in main_snapshots[i].items():
                original_values[(i, c)] = v

        # 用老品补充模式合并
        merge_one_painting(
            main_snapshots=main_snapshots,
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=main_ws,
            col_map=main_col_map,
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
            mode="old_variant",
        )

        # 验证 r4-r14 (普文件原 11 行) 的所有列值没变
        for i in range(11):
            r = 4 + i
            for c in range(1, s["max_col"] + 1):
                orig = original_values[(i, c)]
                now = main_ws.cell(row=r, column=c).value
                # 空字符串和 None 视为相同
                orig_e = orig if orig not in (None, "") else ""
                now_e = now if now not in (None, "") else ""
                assert orig_e == now_e, f"r{r} col{c}: 原值={orig} 现值={now} (老品补充不应修改普文件原行)"

    def test_old_variant_mode_wood_gold_color_price(self):
        """老品补充变体模式: Wood/Gold 行的 Color 和 Price 仍正确填充"""
        s = _setup_merge_one_painting()
        main_ws = s["main_ws"]
        merge_one_painting(
            main_snapshots=s["main_snapshots"],
            wood_group=s["wood_group"],
            gold_group=s["gold_group"],
            output_start_row=4,
            output_ws=main_ws,
            col_map=s["main_col_map"],
            wood_ws=s["wood_ws"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
            mode="old_variant",
        )
        # Wood r15-r19
        for i, r in enumerate(range(15, 20)):
            assert main_ws.cell(row=r, column=55).value == "Vintage Wood Grain Frame-style"
            assert main_ws.cell(row=r, column=154).value == [26.9, 39.9, 59.9, 99.9, 129.9][i]
        # Gold r20-r24
        for i, r in enumerate(range(20, 25)):
            assert main_ws.cell(row=r, column=55).value == "Vintage Ornate Gold Frame-style"
            assert main_ws.cell(row=r, column=154).value == [26.9, 39.9, 59.9, 99.9, 129.9][i]


# ===== rewrite_sku =====

class TestRewriteSku:
    def test_continuous_numbering_new_mode(self):
        """新品上架: 全部 21 行 × N 画从 prefix-1 连续编号"""
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25)), list(range(25, 46))]
        rewrite_sku(ws, groups, prefix="HM725", mode="new")
        assert ws.cell(row=4, column=1).value == "HM725-1"
        assert ws.cell(row=5, column=1).value == "HM725-2"
        assert ws.cell(row=24, column=1).value == "HM725-21"
        assert ws.cell(row=25, column=1).value == "HM725-22"
        assert ws.cell(row=45, column=1).value == "HM725-42"

    def test_single_group_new_mode(self):
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25))]
        rewrite_sku(ws, groups, prefix="AB", mode="new")
        assert ws.cell(row=4, column=1).value == "AB-1"
        assert ws.cell(row=24, column=1).value == "AB-21"

    def test_old_variant_mode_preserves_main_sku(self):
        """老品补充变体: group[0:11] (普文件原 11 行) SKU 保留, group[11:21] 重写"""
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25))]
        # 预设普文件原 11 行的 SKU (group[0:11] = r4-r14)
        for i, r in enumerate(groups[0][:11]):
            ws.cell(row=r, column=1).value = f"OLD-{i+1}"
        rewrite_sku(ws, groups, prefix="NEW", mode="old_variant")
        # r4-r14 保留原 SKU
        assert ws.cell(row=4, column=1).value == "OLD-1"
        assert ws.cell(row=14, column=1).value == "OLD-11"
        # r15-r24 (Wood+Gold) 用新前缀从 1 开始
        assert ws.cell(row=15, column=1).value == "NEW-1"
        assert ws.cell(row=16, column=1).value == "NEW-2"
        assert ws.cell(row=24, column=1).value == "NEW-10"

    def test_old_variant_mode_multi_groups_continuous(self):
        """老品补充变体: 多 group 时 Wood/Gold 跨 group 连续编号"""
        wb = Workbook()
        ws = wb.active
        groups = [list(range(4, 25)), list(range(25, 46))]
        rewrite_sku(ws, groups, prefix="NEW", mode="old_variant")
        # group 1: r15-r24 = NEW-1 到 NEW-10
        assert ws.cell(row=15, column=1).value == "NEW-1"
        assert ws.cell(row=24, column=1).value == "NEW-10"
        # group 2: r36-r45 = NEW-11 到 NEW-20
        assert ws.cell(row=36, column=1).value == "NEW-11"
        assert ws.cell(row=45, column=1).value == "NEW-20"


# ===== write_parent_sku_formulas =====

class TestWriteParentSkuFormulas:
    def test_first_child_uses_b_parent_new_mode(self):
        """新品上架: parent 清空, 第 1 child =B4, 后续 =AA{prev}"""
        wb = Workbook()
        ws = wb.active
        rows = list(range(4, 25))
        write_parent_sku_formulas(ws, [rows], mode="new")
        assert ws.cell(row=4, column=5).value in (None, "")
        assert ws.cell(row=5, column=5).value == "=A4"
        assert ws.cell(row=6, column=5).value == "=E5"
        assert ws.cell(row=7, column=5).value == "=E6"
        assert ws.cell(row=24, column=5).value == "=E23"

    def test_multiple_groups_new_mode(self):
        wb = Workbook()
        ws = wb.active
        write_parent_sku_formulas(ws, [list(range(4, 25)), list(range(25, 46))], mode="new")
        assert ws.cell(row=5, column=5).value == "=A4"
        assert ws.cell(row=6, column=5).value == "=E5"
        assert ws.cell(row=26, column=5).value == "=A25"
        assert ws.cell(row=27, column=5).value == "=E26"

    def test_old_variant_mode_preserves_main_formulas(self):
        """老品补充变体: group[0:11] parent SKU 公式保留, group[11:21] 从 =AA{group[10]} 开始"""
        wb = Workbook()
        ws = wb.active
        rows = list(range(4, 25))
        # 预设普文件原 11 行的 parent SKU 公式
        ws.cell(row=4, column=5).value = None       # parent
        ws.cell(row=5, column=5).value = "=A4"      # Frame 1
        ws.cell(row=6, column=5).value = "=E5"
        ws.cell(row=14, column=5).value = "=E13"   # Unframe 5 (最后一个)
        write_parent_sku_formulas(ws, [rows], mode="old_variant")
        # r4-r14 保留
        assert ws.cell(row=4, column=5).value is None
        assert ws.cell(row=5, column=5).value == "=A4"
        assert ws.cell(row=14, column=5).value == "=E13"
        # r15 (Wood 1) = =AA14 (引用 r14 Unframe 最后一个)
        assert ws.cell(row=15, column=5).value == "=E14"
        # r16 = =AA15
        assert ws.cell(row=16, column=5).value == "=E15"
        # r24 (Gold 5) = =AA23
        assert ws.cell(row=24, column=5).value == "=E23"


# ===== fill_list_price_synced =====

class TestFillListPriceSynced:
    def test_list_price_matches_your_price(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=HEADER_ROW, column=154).value = "Your Price"
        ws.cell(row=HEADER_ROW, column=154).value = "List Price"
        rows = [4, 5, 6]
        ws.cell(row=4, column=154).value = None
        ws.cell(row=5, column=154).value = 19.9
        ws.cell(row=6, column=154).value = 29.9
        fill_list_price_synced(ws, rows, col_map={"Your Price": 13, "List Price": 145})
        assert ws.cell(row=4, column=154).value is None
        assert ws.cell(row=5, column=154).value == 19.9
        assert ws.cell(row=6, column=154).value == 29.9

    def test_missing_columns_no_op(self):
        wb = Workbook()
        ws = wb.active
        ws.cell(row=HEADER_ROW, column=154).value = "Your Price"
        rows = [4, 5]
        ws.cell(row=5, column=154).value = 19.9
        fill_list_price_synced(ws, rows, col_map={"Your Price": 13})
        assert ws.cell(row=5, column=154).value == 19.9


# ===== build_sku_prefix =====

class TestBuildSkuPrefix:
    def test_single_string(self):
        assert build_sku_prefix("HM725") == "HM725"

    def test_strip_whitespace(self):
        assert build_sku_prefix("  HM725  ") == "HM725"

    def test_complex_prefix(self):
        assert build_sku_prefix("AB2026风景") == "AB2026风景"


# ===== 木/金可选 (动态行数 11/16/21) =====

class TestOptionalVariants:
    """木框/金框独立可选: 输出 11/16/21 行。"""

    def test_wood_only_16_rows(self):
        """只提供木框 (无金框) → 16 行, 木 color 在 index 11-15, 无金行。"""
        s = _setup_merge_one_painting()
        merged = merge_one_painting(
            main_snapshots=s["main_snapshots"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_group=s["wood_group"],
            wood_ws=s["wood_ws"],
            gold_group=None,
            gold_ws=None,
            max_col=s["max_col"],
        )
        assert len(merged) == 16
        assert merged[0] == 4
        assert merged[-1] == 19
        ws = s["main_ws"]
        colors = [ws.cell(row=r, column=55).value for r in merged]
        assert colors[0] in (None, "")
        assert colors[1] == "Frame-style"
        assert colors[10] == "Unframe-style"
        for i in range(11, 16):
            assert colors[i] == "Vintage Wood Grain Frame-style"
        prices = [ws.cell(row=r, column=154).value for r in merged]
        assert prices[11:16] == [26.9, 39.9, 59.9, 99.9, 129.9]

    def test_gold_only_16_rows(self):
        """只提供金框 (无木框) → 16 行, 金 color 紧随 main 占 index 11-15。"""
        s = _setup_merge_one_painting()
        merged = merge_one_painting(
            main_snapshots=s["main_snapshots"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_group=None,
            wood_ws=None,
            gold_group=s["gold_group"],
            gold_ws=s["gold_ws"],
            max_col=s["max_col"],
        )
        assert len(merged) == 16
        ws = s["main_ws"]
        colors = [ws.cell(row=r, column=55).value for r in merged]
        for i in range(11, 16):
            assert colors[i] == "Vintage Ornate Gold Frame-style"
        # 金行数据来自 gold 文件 (gold 第 1 个 child SKU = GOLD-9, 新格式 DATA_START_ROW=8)
        assert ws.cell(row=merged[11], column=1).value == "GOLD-9"

    def test_main_only_new_mode_11_rows(self):
        """无木无金, new 模式 → 11 行 (parent + Frame×5 + Unframe×5)。"""
        s = _setup_merge_one_painting()
        merged = merge_one_painting(
            main_snapshots=s["main_snapshots"],
            output_start_row=4,
            output_ws=s["main_ws"],
            col_map=s["main_col_map"],
            wood_group=None,
            wood_ws=None,
            gold_group=None,
            gold_ws=None,
            max_col=s["max_col"],
            mode="new",
        )
        assert len(merged) == 11
        assert merged[-1] == 14

    def test_old_variant_requires_variant(self):
        """old_variant 模式无变体 → 报错。"""
        s = _setup_merge_one_painting()
        with pytest.raises(ValueError, match="至少一个木框或金框"):
            merge_one_painting(
                main_snapshots=s["main_snapshots"],
                output_start_row=4,
                output_ws=s["main_ws"],
                col_map=s["main_col_map"],
                wood_group=None,
                gold_group=None,
                mode="old_variant",
            )

    def test_old_variant_wood_only_preserves_main(self):
        """old_variant + 只木 → main 11 行不动, 木 5 行处理, 无金行。"""
        s = _setup_merge_one_painting()
        main_ws = s["main_ws"]
        main_snapshots = s["main_snapshots"]
        original = {}
        for i in range(11):
            for c, v in main_snapshots[i].items():
                original[(i, c)] = v
        merge_one_painting(
            main_snapshots=main_snapshots,
            output_start_row=4,
            output_ws=main_ws,
            col_map=s["main_col_map"],
            wood_group=s["wood_group"],
            wood_ws=s["wood_ws"],
            gold_group=None,
            gold_ws=None,
            max_col=s["max_col"],
            mode="old_variant",
        )
        # main 11 行 (r4-r14) 完全不变
        for i in range(11):
            r = 4 + i
            for c in range(1, s["max_col"] + 1):
                orig = original[(i, c)]
                now = main_ws.cell(row=r, column=c).value
                orig_e = orig if orig not in (None, "") else ""
                now_e = now if now not in (None, "") else ""
                assert orig_e == now_e, f"r{r} col{c}: 原值={orig} 现值={now}"
        # 木行 r15-r19
        for i, r in enumerate(range(15, 20)):
            assert main_ws.cell(row=r, column=55).value == "Vintage Wood Grain Frame-style"
            assert main_ws.cell(row=r, column=154).value == [26.9, 39.9, 59.9, 99.9, 129.9][i]
        # 无金行
        assert main_ws.cell(row=20, column=55).value in (None, "")


# ===== merge_files 集成 (木/金可选) =====

class TestMergeFilesOptional:
    """merge_files 端到端: 木/金可选 + 配对错误处理。"""

    def test_gold_only_missing_painting_raises_cleanly(self, tmp_path):
        """金 only 且金文件缺某画 → _raise_pairing_error 干净触发 (无 None 崩溃)。"""
        from amazon_excel_processor.merger import merge_files
        main_wb, _ = _create_main_workbook(["Art A", "Art B"])
        gold_wb = _create_variant_workbook(["Art A"], role="gold")  # 缺 Art B
        main_p = tmp_path / "main.xlsx"
        gold_p = tmp_path / "gold.xlsx"
        main_wb.save(str(main_p))
        gold_wb.save(str(gold_p))
        with pytest.raises(ValueError) as excinfo:
            merge_files(main_path=main_p, wood_path=None, gold_path=gold_p,
                        sku_prefix="T", mode="new")
        msg = str(excinfo.value)
        assert "配对失败" in msg
        assert "未提供" in msg  # 木框未提供
        assert "金框" in msg

    def test_wood_only_merge_files_16_rows(self, tmp_path):
        """merge_files 木 only → 输出 16 行/组, 无金行。"""
        from amazon_excel_processor.merger import merge_files
        from openpyxl import load_workbook as _lw
        main_wb, _ = _create_main_workbook(["Art A"])
        wood_wb = _create_variant_workbook(["Art A"], role="wood")
        main_p = tmp_path / "main.xlsx"
        wood_p = tmp_path / "wood.xlsx"
        main_wb.save(str(main_p))
        wood_wb.save(str(wood_p))
        out = merge_files(main_path=main_p, wood_path=wood_p, gold_path=None,
                          sku_prefix="T", mode="new")
        ws = _lw(str(out))["Template"]
        # parent + 10 main + 5 wood = 16 行 (新格式 DATA_START_ROW=8, r8-r23)
        assert ws.cell(row=8, column=4).value == "Parent"
        assert ws.cell(row=9, column=4).value == "Child"
        assert ws.cell(row=23, column=55).value == "Vintage Wood Grain Frame-style"
        assert ws.cell(row=24, column=55).value in (None, "")  # 无金行


# ===== 向后兼容: 动态序列 == 21 行常量 =====

class TestDynamicSequencesBackwardCompat:
    """build_active_styles(True,True) 动态序列必须与现有 *_21 常量逐元素相等。"""

    def test_sequences_match_21_constants(self):
        from amazon_excel_processor.field_filler import (
            build_active_styles, _build_sequences,
            COLOR_SEQUENCE_21, SIZE_MAP_SEQUENCE_21, SIZE_32_21,
            LENGTH_32_21, WIDTH_32_21, WEIGHT_SEQUENCE_21, PRICE_SEQUENCE_21,
        )
        seqs = _build_sequences(build_active_styles(True, True))
        assert seqs["color"] == COLOR_SEQUENCE_21
        assert seqs["size_map"] == SIZE_MAP_SEQUENCE_21
        assert seqs["size_32"] == SIZE_32_21
        assert seqs["length"] == LENGTH_32_21
        assert seqs["width"] == WIDTH_32_21
        assert seqs["weight"] == WEIGHT_SEQUENCE_21
        assert seqs["price"] == PRICE_SEQUENCE_21
        assert seqs["edge"] == [1] + [12, 18, 24, 30, 36] * 4
        assert seqs["labels"][1:] == VARIANT_LABELS_21[1:]
