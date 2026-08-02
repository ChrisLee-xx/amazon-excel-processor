# 金框/木框 Excel 改为可选

## Context（背景）

当前三文件合并模式（`merger.py`）要求每次必须提供「普文件 + 木框 + 金框」三个文件，输出固定 21 行/组（parent + Frame×5 + Unframe×5 + Wood×5 + Gold×5）。

实际业务中，**不一定每次都同时有金框和木框 Excel，可能只有一种**。因此需要把木框、金框文件都改为**独立可选**：

- 普文件（主）— 始终必填（Frame+Unframe 来源）
- 木框文件 — 可选
- 金框文件 — 可选

输出行数随之动态变化：1 + 5×(2 + 有木 + 有金)，即 11 / 16 / 16 / 21 四种情况。当木/金都存在时，输出必须与现状**完全一致**（向后兼容）。

## 核心设计：动态 style 计划

把"哪些 style 参与输出"参数化。Frame+Unframe 永远来自普文件（固定 11 行），Wood/Gold 按需追加在后面。

`active_styles = ["frame","unframe"] + (["wood"] if 有木 else []) + (["gold"] if 有金 else [])`

每种 style 的 5 尺寸字段值（color/size_map/size_32/length/width/weight/price/edge）从 `STYLE_SPECS` 注册表拼接出逐行序列，再按 `enumerate(rows)` 填充。当 `active_styles=[frame,unframe,wood,gold]` 时，动态序列与现有 `_21` 常量逐元素相等（向后兼容，已验证）。

**关键不变量**：普文件恒为 11 行（parent+10），变体行恒从第 11 个位置（index 11）开始追加。因此 `rewrite_sku` / `write_parent_sku_formulas` 的 old_variant 分支（用 `group[11:]`、`group[10]`）**无需修改**，对所有组合都正确。

## 修改清单

### 1. `src/amazon_excel_processor/field_filler.py`
- 新增 `STYLE_SPECS` 注册表（frame/unframe/wood/gold → label/weight/price）和共享的 5 尺寸常量（size_map/size_32/length/width/edge）。
- 新增 `build_active_styles(has_wood, has_gold)` 和 `_build_sequences(active_styles)`（返回含 parent 占位的逐行序列 dict：color/size_map/size_32/length/width/weight/price/edge/labels）。
- 把 `fill_group_21` 重构为 `fill_group_merged(ws, rows, col_map, ratio_type, active_styles)`：构建序列后按字段填充。**保留现有行为怪癖**：`fill_size` 对 square 跳过（保留用户预填），`fill_length`/`fill_width` 始终用 3:2 值（不随 ratio 变），`item_length_longer_edge` = `[1] + [12,18,24,30,36] * len(active_styles)`。
- `fill_group_21` 保留为薄包装（调用 `fill_group_merged(..., [frame,unframe,wood,gold])`）以向后兼容。
- 保留现有 `_21` 常量作为 21 行布局的参考定义（不删除）。

### 2. `src/amazon_excel_processor/merger.py`
- `merge_files`：`wood_path`/`gold_path` 改为 `Optional[Path]=None`（保持参数顺序，gui 用 kwargs 调用，位置调用也兼容）。仅当提供时才 `load_workbook`/`group_rows`/`index_groups_by_name`/role 校验。`has_wood`/`has_gold` 标志。`merged_group_size = 1 + 5*(2+has_wood+has_gold)`，`out_row += merged_group_size`。
- **配对循环（关键修复）**：跳过条件改为 `(has_wood and idx >= len(wood_list)) or (has_gold and idx >= len(gold_list))`；`wood_g = wood_list[idx] if has_wood else None`。文件整体缺失时**不报错**，只有"文件存在但缺某画"才进 skipped。
- `max_col_for_snapshot` 和 `_snapshot_row` 调用处都要 `if has_wood`/`if has_gold` 守卫（两处崩溃点：merger.py:600 和 merge_one_painting 内 252-253）。
- `logger.info` 行对 None 路径守卫（`wood_path.name if wood_path else "无"`）。
- `_raise_pairing_error`：absent 文件按"未提供/N/A"报告，不列候选（守卫 `list(wood_by_name.keys())` 和 `_get_raw_name(wood_ws,...)` 的 None 崩溃）。
- `merge_one_painting`：`wood_group`/`gold_group`/`wood_ws`/`gold_ws` 加 `=None` 默认值，内部 `has_wood = wood_group is not None` 推断（零改动现有调用方/测试）。断言改为仅当存在时校验。**变体行写动态紧凑偏移**：wood 在 offset 11（若有），gold 在 `11 + 5*has_wood`（若有）——不要再用硬编码 16。`merged_rows` 动态构建。
  - new 模式：`normalize_group_merged` + `fill_group_merged(active_styles)` + `_fill_meta_columns`。
  - old_variant 模式：`variant_rows = merged_rows[11:]`（恒从 11 起），调用 `_normalize_variant_names(variant_styles=active_styles[2:])` + `_fill_variant_fields(variant_styles=...)` + `_fill_meta_columns_variant`。
- `normalize_group_21` → `normalize_group_merged(ws, rows, name_col, ratio_type, active_styles)`：labels 从 STYLE_SPECS+active_styles 构建，`sizes[(i-1)%5]`，**始终用 SIZES_32**（保留现状，不随 ratio 切换）。
- `_normalize_variant_names` / `_fill_variant_fields`：加 `variant_styles` 参数，构建**仅含变体 style** 的序列，从 index 0 应用到 variant_rows（不再用 offset 11 索引 21 元素序列）。`edge_values = [12,18,24,30,36] * len(variant_styles)`（修复硬编码 `*2`）。
- `rewrite_sku` / `write_parent_sku_formulas` / `_check_pairing` / `_fill_meta_columns*`：**不改**。

### 3. `src/amazon_excel_processor/gui_entry.py`
- 交互式合并：木框/金框路径提示改为"可留空跳过"。校验改为"普文件必填，木/金可选"。`_run_merge` 在选完上架类型后立即校验：old_variant 模式且无任何变体文件 → 报错"老品补充变体需要至少一个木框或金框文件"；new 模式允许木/金皆无（输出 11 行）。
- CLI：保留 3 位置参数（主 木 金）向后兼容；新增 `--wood PATH` / `--gold PATH` 可选 flag 支持"只有木"或"只有金"。1 位置参数=单文件。2 位置参数 → 报错并提示用 `--wood`/`--gold`。3 位置 + flag 同时给 → 3 位置优先并告警。

### 4. `src/amazon_excel_processor/name_normalizer.py`
- 保留 `VARIANT_LABELS_21`（`tests/test_merger.py:8` 有未使用但存在的 import，删除常量会致收集期报错）。

### 5. `tests/test_merger.py`
- 现有测试（木+金→21 行）**保持不变应通过**。
- 新增：木 only → 16 行（wood color 在 11-15，无 gold）；金 only → 16 行（gold color 在 11-15）；new 模式 main only → 11 行；old_variant 木 only → main 11 行不动 + 5 木行处理。
- 新增 `merge_files` 集成测试：金 only 且金文件缺某画 → `_raise_pairing_error` 干净触发（无 None 崩溃，金候选列出，木报"未提供"）。
- 可选：加一个守卫测试，断言 `build_active_styles(True,True)` 构建的序列与现有 `_21` 常量逐元素相等（锁定向后兼容）。

### 6. `README.md`
- 更新合并模式章节：木框/金框可选，输出行数 11/16/21 视组合而定；补充 `--wood`/`--gold` CLI 用法。

## 验证方式

1. `poetry run pytest tests/ -q` — 全绿（含新增测试）。
2. 交互式跑 `poetry run python -m amazon_excel_processor.gui_entry`：
   - 选合并 → 只输普+木 → 新品上架 → 输出 16 行/组，木 color 正确，无金行。
   - 只输普+金 → 输出 16 行/组，金 color 在变体位。
   - 普+木+金 → 输出 21 行/组（与旧版一致）。
   - 只输普 → 新品上架 → 11 行/组；选老品补充 → 报错。
3. CLI：`gui_entry 普.xlsm --wood 木.xlsm`（16 行）；`gui_entry 普.xlsm 木.xlsm 金.xlsm`（21 行，向后兼容）。
4. 用 MCP `excel_read_sheet` 抽查输出文件的 Color/Your Price/Parent SKU 列是否符合预期 style 顺序与价格序列。
