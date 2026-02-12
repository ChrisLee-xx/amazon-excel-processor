# Amazon Excel Processor

亚马逊上架商品 Excel 模板批量规范化处理工具。

## 直接使用（无需安装开发环境）

### Mac
1. 从 `dist/` 目录拿到 `amazon-excel-processor` 文件
2. 两种使用方式：
   - **拖拽**：把 `.xlsm` 文件拖到 `amazon-excel-processor` 图标上
   - **终端**：`./amazon-excel-processor '你的文件.xlsm'`
3. 处理完成后输出文件为 `{原文件名}_processed.xlsm`，在同一目录下

> 首次打开可能提示"无法验证开发者"，右键选"打开"即可。

### Windows
1. 从 `dist/` 目录拿到 `amazon-excel-processor.exe` 文件
2. 两种使用方式：
   - **拖拽**：把 `.xlsm` 文件拖到 `.exe` 图标上
   - **双击**：双击运行后粘贴文件路径
3. 处理完成后输出文件为 `{原文件名}_processed.xlsm`，在同一目录下

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

# 基本用法
poetry run excel-process 你的文件.xlsm

# 显示详细日志
poetry run excel-process 你的文件.xlsm -v

# 指定输出路径
poetry run excel-process 你的文件.xlsm -o 输出文件.xlsm
```

## 处理内容

程序读取 Excel 中的 **Template** tab，按每 11 行一组（1 parent + 5 Frame + 5 Unframe）处理：

### Product Name 规范化
1. 多空格合并为单空格
2. 变体行按固定顺序重构为 `{标题} Frame-style {尺寸}` / `{标题} Unframe-style {尺寸}`
3. 删除 `-1`、`-2` 等数字后缀
4. 连字符 `-` 替换为空格（保留 `Frame-style` / `Unframe-style`）
5. 下划线 `_` 替换为空格
6. 单词去重（同一单词最多保留 2 次）

### 字段填充
| 字段 | 填充值 |
|------|--------|
| Variation Theme | `color-size` |
| Paint Type | `Oil` |
| Color Map | `Multi` |
| Color | 空, Frame-style×5, Unframe-style×5 |
| Size | 按比例类型填充（3:2 或正方形） |
| Size Map | 空, X-Small, Small, Medium, Large, X-Large ×2 |
| Length | 按比例类型填充 |
| Weight | 空, 0.18, 0.28, 0.48, 0.68, 0.88, 0.02, 0.04, 0.07, 0.15, 0.25 |

### 比例类型自动检测
- 标题含 `12x12`/`16x16`/`20x20`/`24x24`/`28x28` → 正方形
- 其余 → 3:2

## 输出

输出文件保留原文件所有 sheet，仅替换 Template tab 中的数据。
