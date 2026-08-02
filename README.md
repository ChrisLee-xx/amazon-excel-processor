# Amazon Excel Processor

亚马逊上架商品 Excel 模板批量规范化处理工具。

支持两种模式：
- **单文件模式**：1 个 Excel 走 11 行/组规范化处理
- **合并模式**：以"普文件"（11 行/组，Frame+Unframe）为主，"木框文件"+"金框文件"（各 6 行/组）**可选**地合并进来。木框/金框可只提供一个、都提供、或都不提供，输出 11/16/21 行/组

## 直接使用（无需安装开发环境）

### Mac
1. 从 `dist/` 目录拿到 `amazon-excel-processor` 文件
2. 使用方式：
   - **拖拽**：把 `.xlsm` 文件拖到 `amazon-excel-processor` 图标上（单文件模式）
   - **双击**：双击运行后按提示选择模式 1 或 2

> 首次打开可能提示"无法验证开发者"，右键选"打开"即可。

### Windows
1. 从 `dist/` 目录拿到 `amazon-excel-processor.exe` 文件
2. 使用方式：
   - **拖拽**：把 `.xlsm` 文件拖到 `.exe` 图标上
   - **双击**：双击运行后粘贴文件路径 / 选择模式

## 使用流程

### 单文件模式
把单个 `.xlsm` 文件拖到程序上即可。处理完成后输出 `{原文件名}_processed.xlsm`，在同一目录下。

### 合并模式（木/金可选）

适用于"多种框型合并到一个 listing"的场景。**普文件必填，木框/金框可选**（可只提供一种），由用户在 GUI 中明确指定每个文件的角色，避免金文件内部 Wood/Gold 顺序混淆：
- **普文件（主）**：11 行/组，含 `Frame-style` + `Unframe-style` 两种框型（必填）
- **木框文件**：6 行/组，每画 1 个 group，对应 `Vintage Wood Grain Frame-style`（可选）
- **金框文件**：6 行/组，每画 1 个 group，对应 `Vintage Ornate Gold Frame-style`（可选）
- **输出**：每组行数 = 1 + 5×(2 + 有木 + 有金)，即 11/16/21 行
  - 普文件 only → 11 行（Frame×5 + Unframe×5）
  - 普 + 木（或 普 + 金）→ 16 行（追加 1 种变体 ×5）
  - 普 + 木 + 金 → 21 行（Frame×5 + Unframe×5 + Wood×5 + Gold×5）

> **为什么要拆多个文件？** 金/木文件内部的所有字段（Frame Type / Frame Material / Color / Theme）都相同，无法从字段区分哪组是 Wood 哪组是 Gold。拆成独立文件后由用户在 GUI 中明确指定，最可靠。
>
> **木/金可选**：实际业务中不一定每次都同时有金框和木框 Excel，可能只有一种。木框、金框都改为独立可选，文件整体缺失时跳过该变体（不报错）；只有"文件存在但缺某画"才报配对错误。

#### 步骤
1. 双击运行 `.exe` / 程序
2. 选 `2) 合并模式 (普 + 木 + 金，木/金可选)`
3. 按提示依次输入文件路径（木框/金框可留空跳过）：
   - 普文件（主文件，必须含 Frame-style + Unframe-style）
   - 木框文件（每画 1 个 group，留空跳过）
   - 金框文件（每画 1 个 group，留空跳过）
4. 选择上架类型：
   - `1) 新品上架`：所有 SKU（普 + 提供的木/金）按新命名规则重新编号
   - `2) 老品补充变体`：普文件原 SKU 保留，仅新增的木/金 SKU 按新规则编号（此模式需要至少一个木/金文件）
5. 输入 SKU 命名（推荐格式：店铺名+日期+主题，如 `HM725`，非空即可）
6. 程序自动生成合并文件，输出为 `{普文件名}_processed.xlsm`

#### 合并规则
- **识别**：
  - 11 行/组的文件 = 普文件（主）
  - 6 行/组的文件 = 木框或金框（由用户在 GUI 指定）
- **配对**：按归一化后的 Product Name base name 配对（去 `-数字`、去 Frame-/Unframe-、替换 `-` 为空格、转小写）。文件整体缺失则跳过该变体；文件存在但缺某画则报错（带模糊匹配候选）
- **输出顺序**：Frame-style → Unframe-style → Vintage Wood Grain Frame-style（若提供木）→ Vintage Ornate Gold Frame-style（若提供金）
- **上架类型**：
  - **新品上架**：全部行 × N 画 SKU 从 `{前缀}-1` 开始连续编号，所有 parent SKU 公式重写
  - **老品补充变体**：普文件原 11 行（parent + Frame×5 + Unframe×5）SKU 和 parent SKU 公式保留不变；新增的变体行（Wood×5 和/或 Gold×5）SKU 从 `{前缀}-1` 开始连续编号，parent SKU 公式从 `=AA{prev_unframe_last}` 开始链式引用
- **Parent SKU 公式规则**：
  - 新品上架：parent 行清空，第 1 个 child `=B{parent_row}`，后续 `=AA{prev_row}`
  - 老品补充变体：普文件原 11 行保留，新增变体行从 `=AA{unframe_last_row}` 开始链式
- **Parentage**：parent=Parent，所有 child=Child
- **Relationship Type**：parent 留空，child=Variation
- **List Price = Your Price**（每行同步填）
- **Price 序列**（按 style × 5 size）：
  - Frame：19.9 / 29.9 / 45 / 75 / 99
  - Unframe：11.9 / 14.9 / 19.9 / 24.9 / 34.9
  - Wood：26.9 / 39.9 / 59.9 / 99.9 / 129.9
  - Gold：26.9 / 39.9 / 59.9 / 99.9 / 129.9
- **Size / Size Map / Length / Width / Weight**：金框和木框的与 Frame×5 完全一致
- **Image URL**：Wood 行用木框文件的图片，Gold 行用金框文件的图片（来自对应输入文件）

## 打包（开发者）

需要先安装 Python 开发环境。

```bash
cd amazon-excel-processor
poetry install

# 打包当前平台的可执行文件
poetry run python build.py
# 输出在 dist/ 目录下
```

> **注意**：Mac 上打包只能生成 Mac 版，Windows 上打包只能生成 Windows 版。需分别在两个平台上执行打包。

## 开发者命令行用法

```bash
poetry install

# 单文件模式
poetry run excel-process 你的文件.xlsm
poetry run excel-process 你的文件.xlsm -v          # 详细日志
poetry run excel-process 你的文件.xlsm -o 输出.xlsm # 指定输出路径

# 合并模式 (交互式, 程序会询问上架类型/SKU 命名)
# 三文件: 顺序 普文件 木框文件 金框文件 (向后兼容)
poetry run python -m amazon_excel_processor.gui_entry 普文件.xlsm 木框文件.xlsm 金框文件.xlsm

# 木/金可选: 1 个普文件 + --wood / --gold 指定 (各自可选)
poetry run python -m amazon_excel_processor.gui_entry 普文件.xlsm --wood 木框文件.xlsm        # 只木 → 16 行
poetry run python -m amazon_excel_processor.gui_entry 普文件.xlsm --gold 金框文件.xlsm       # 只金 → 16 行
poetry run python -m amazon_excel_processor.gui_entry 普文件.xlsm --wood 木.xlsm --gold 金.xlsm  # 都有 → 21 行
poetry run python -m amazon_excel_processor.gui_entry 普文件.xlsm --mode merge              # 仅普 → 11 行 (新品上架)
```

## 处理内容

### 单文件模式：每 11 行一组（1 parent + 5 Frame + 5 Unframe）

#### Product Name 规范化
1. 多空格合并为单空格
2. 变体行按固定顺序重构为 `{标题} Frame-style {尺寸}` / `{标题} Unframe-style {尺寸}`
3. 删除 `-1`、`-2` 等数字后缀
4. 连字符 `-` 替换为空格（保留 `Frame-style` / `Unframe-style`）
5. 下划线 `_` 替换为空格
6. 单词去重（同一单词最多保留 2 次）

#### 字段填充
| 字段 | 填充值 |
|------|--------|
| Variation Theme | `color-size` |
| Paint Type | `Oil` |
| Color Map | `Multi` |
| Color | 空, Frame-style×5, Unframe-style×5 |
| Size | 3:2 填固定值；正方形保留用户预填值不覆盖 |
| Size Map | 空, X-Small, Small, Medium, Large, X-Large ×2 |
| Length | 按比例类型填充 |
| Width | 3:2 填宽度值；正方形与 Length 一致 |
| Weight | 空, 0.18, 0.28, 0.48, 0.68, 0.88, 0.02, 0.04, 0.07, 0.15, 0.25 |
| Your Price | 空, 19.9, 29.9, 45, 75, 99, 11.9, 14.9, 19.9, 24.9, 34.9 |

#### 比例类型自动检测
- 解析 Size 列预填值中的两个数字，L==W → 正方形（保留预填值不覆盖）
- L!=W → 3:2（脚本填充固定尺寸）
- Size 列为空 → 3:2（脚本填充固定尺寸）

### 合并模式：每组 11/16/21 行（1 parent + 提供的 style × 5 size）

输出文件与单文件模式的字段填充规则一致，但按提供的变体动态扩展：
- **Color** 按提供的变体追加：`Vintage Wood Grain Frame-style` ×5（若有木）、`Vintage Ornate Gold Frame-style` ×5（若有金）
- **Your Price** 按提供的变体追加价格序列：Wood / Gold 各 5 个
- **List Price** 列与 Your Price 同步填相同值
- **Seller SKU** 全部重写为 `{前缀}-N` 格式（新品上架）；老品补充变体只重写新增变体行
- **Parent SKU** 用公式 `=B{parent_row}` 和 `=AA{prev_row}` 引用

## 输出

输出文件保留原文件所有 sheet，仅替换 Template tab 中的数据。
