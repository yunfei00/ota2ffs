# ota2ffs

OTA2FFS Converter 是一个 Python 桌面工具，用于将 OTA 暗室测试 Excel 数据转换为 CST 可导入的 `.ffs` 远场源文件。

## 功能

- 选择一个 Excel 文件。
- 输入频率并选择单位，支持 `MHz`、`GHz`、`KHz`、`Hz`。V2 数据会优先使用 Excel 表格自带频率，V1 留空默认使用 `1e9 Hz`。
- 读取一个或多个 sheet。
- 自动识别每个 sheet 的 V1 或 V2 数据格式。
- 每个 sheet 独立转换，某个 sheet 失败不会影响其他 sheet。
- 为每个可转换 sheet 输出 RX / TX `.ffs` 文件。
- 输出目录默认使用工具路径下的 `output/`，也可以手动选择。
- 支持打开输出目录。
- 转换日志默认生成到工具路径下的 `log/`。
- 生成雷达图报表 Excel，直接读取原始 V1/V2 数据块，不依赖 `.ffs` 文件。

重要说明：本工具不会修改原始 Excel，所有转换、分析和图表结果都会生成到新的输出文件中。

## FFS 输出格式

工具生成的 `.ffs` 文件会先写入 CST 远场源文件头部。每个输出文件当前包含一个频点，所以 `#Frequencies` 为 `1`，频率统一换算成 Hz 后写入头部最后一行：

```text
// CST Farfield Source File

// Version:
3.0

// Data Type
Farfield

// #Frequencies
1

// Position
0 0 0          // 单位 m

// z-Axis
0 0 1

// x-Axis
1 0 0

// 对每个频点，各 4 行
-1
-1
-1
2450000000

// >> Total #phi samples, total #theta samples
7   4
```

频率下方会写入真实测试数据中的 Phi / Theta 采样数量。头部之后的数据主体为空格分隔文本，列顺序固定为：

```text
// >> Phi, Theta, Re(E_Theta), Im(E_Theta), Re(E_Phi), Im(E_Phi):
0 0 10 0 1.99526231497 0
```

V2 sheet 会使用表格中自带的频率，单位固定为 `MHz`。V1 sheet 或没有自带频率的数据会使用界面中输入的频率；如果界面频率留空，则默认使用 `1e9 Hz`。

其中：

- `Im(E_Theta)` 始终为 `0`。
- `Im(E_Phi)` 始终为 `0`。
- Excel 中缺失的角度或数据点会以线性值 `0` 输出。

## RX / TX 区别

RX 文件使用 Excel 中的原始 dB 值转换为线性值：

```text
linear = 10 ** (db / 20)
```

TX 文件会先将原始 dB 值乘以 `-1`，再转换为线性值：

```text
linear = 10 ** ((-db) / 20)
```

## V1 Excel 格式

V1 sheet 至少包含 `Theta`、`Phi` 两个数据块，也可以额外包含 `Total` 数据块：

- 第一块：`A1 = Theta`，`B1` 内容包含 `Phi Angle`，`B2` 内容包含 `Theta Angle`（允许额外符号或单位）。
- `C1` 开始为 Theta 角度。
- `B3` 开始为 Phi 原始角度。
- 第一块交叉区域写入 `Re(E_Theta)`，也就是 FFS 第 3 列。
- 第二块：结构相同，典型位置为 `A30 = Phi`，`B30` 内容包含 `Phi Angle`，`B31` 内容包含 `Theta Angle`。
- 第二块交叉区域写入 `Re(E_Phi)`，也就是 FFS 第 5 列。
- 可选 `Total` 块结构相同，会单独生成 `_total` 后缀 FFS，数据写入 `Re(E_Theta)`（第 3 列），`Re(E_Phi)` 全部为 `0`。
- Phi 输出角度为 Excel 中 Phi 原始角度加 `180`。
- Phi 输出范围补齐为 `0~360`。
- Theta 输出范围补齐为 `0~180`。
- 角度步进根据 Excel 实际数据自动识别，不写死为 `15`。
- V1 数据块可以放在 sheet 内不同起始行列，只要块内相对结构保持一致即可。

## V2 Excel 格式

V2 sheet 中包含三个表格，每个表格左上角单元格为 `Polarization`：

- `Polarization` 后一个单元格分别为 `Theta`、`Phi`、`Total`。
- 频率在当前表格 `Polarization` 行内读取：优先使用 `Freq` 单元格后一个单元格中的数字；如果没有该数字，则使用同一行最后一个有效数字。
- 频率数值不带单位。
- V2 频率单位固定按 `MHz` 处理。
- 下一行格式为 `Phi\Theta, 0, 30, 60, ..., 180`。
- 后续行为 `Phi角度, value1, value2, ...`。

转换规则：

- `Theta` 表格写入 `Re(E_Theta)`。
- `Phi` 表格写入 `Re(E_Phi)`。
- `Theta + Phi` 合成普通 FFS。
- `Total` 表格单独生成拓图 FFS，数据写入 `Re(E_Theta)`，`Re(E_Phi)` 全部为 `0`。
- V2 输出文件头部频率使用表格自带频率，不使用界面输入的频率。
- V2 三个表格可以放在 sheet 内不同起始行列，只要每个表格内部相对结构保持一致即可。

## 第二阶段：雷达图报表

Radar Report Generator 会直接读取原始 Excel 中的 V1/V2 数据块，解析为统一矩阵对象后，新建一个报表 Excel：

```text
原始Excel文件名_Radar_Report.xlsx
```

程序流程为：

```text
读取原始 Excel（只读）
     ↓
解析为标准矩阵对象 PatternMatrix
     ↓
创建新的 Workbook
     ↓
写入标准化数据、雷达图和处理日志
     ↓
保存为新的 report xlsx 文件
```

原始 Excel 使用 `read_only=True` 打开，报表使用 `Workbook()` 新建；程序不会在原始 workbook 上新增 sheet、修改单元格、插入图表或保存。

报表至少包含：

- `Radar_Report`：雷达图展示区。
- `Normalized_Data`：标准化后的数据表，尽量保留原始矩阵形状。
- `Process_Log`：处理日志。

`Normalized_Data` 布局：

- 每个 `Theta`、`Phi`、`Total` 会生成独立纵向区域，不会把三类数据横向塞在同一行。
- 左侧按原始矩阵形状写入标准化数据：第一行是 Theta 角度，第一列是 Phi 角度，中间是乘以 `-1` 后的正值。
- 多个 sheet 的同一 `block_name` 会在同一区域内上下排列，便于对照原始数据结构。
- 多 sheet 对比数据写在同一 block 区域右侧，供对比雷达图引用，避免把图表数据拆散到报表各处。
- 右侧对比数据通过公式引用左侧标准化矩阵。手动修改左侧矩阵数值后，右侧对比数据和雷达图会随 Excel 重新计算而同步更新。

雷达图规则：

- 雷达图采用接近 Excel 原生“插入雷达图”的标准样式：右侧图例、默认坐标轴标签和默认网格线，不额外叠加数据标签。
- 所有原始数值在绘图前乘以 `-1`，例如 `-20` 绘制为 `20`。
- 空值、缺失值、非数字值按 `0` 处理。
- 不进行平均聚合。
- 每个矩阵的每一行生成一个 Row 雷达图。
- 每个矩阵的每一列生成一个 Col 雷达图。
- V1 原始 Phi 行角度按 FFS 规则统一加 `180` 后参与绘图和多 sheet 对比，例如 `-180~180` 会转换为 `0~360`。

Row 雷达图含义：

- 雷达轴为 `col_angles`，通常是 Theta 角度。
- 数据为当前行所有列的正值。
- 标题格式为 `{sheet_name}_{block_name}_Row_{row_angle}`。

Col 雷达图含义：

- 雷达轴为 `row_angles`，通常是 Phi 角度。
- 数据为当前列所有行的正值。
- 标题格式为 `{sheet_name}_{block_name}_Col_{col_angle}`。

多 sheet 对比图：

- 当用户勾选多个 sheet 时，会生成 `Compare Charts` 区域，并将不同 sheet 中相同 `block_name`、相同角度的 Row/Col 图合并为一张雷达图。
- 只比较相同 `block_name` 的数据。
- Row 对比图按相同 `row_angle` 合并，不同 sheet 作为不同系列。
- Col 对比图按相同 `col_angle` 合并，不同 sheet 作为不同系列。
- 某个 sheet 缺少角度时，该系列对应值补 `0`。
- 多 sheet 模式不会再重复生成每个 sheet 的独立 Row/Col 图，避免相同格式数据占用过多横向空间。
- `Theta`、`Phi`、`Total` 对比图按 block 分成多个纵向区域；每个 block 内第一行放 Row 对比图，第二行放 Col 对比图。

场景差值图：

- 勾选“生成场景差值图”后，会以第一个选中的 sheet 作为基础场景。
- 后续每个 sheet 会分别与基础场景做差值，差值公式为 `目标 sheet 标准化值 - 基础 sheet 标准化值`。
- 差值数据写入 `Normalized_Data` 的 `Delta Data` 区域，并通过公式引用左侧标准化矩阵。
- 差值图写入 `Radar_Report` 的 `Delta Charts` 区域；相同 block 和角度的多个差值 series 会合并到同一张雷达图中，便于直接对比。
- 只对相同 `block_name` 且双方共同存在的角度生成差值数据和雷达图。

## 输出文件命名

假设 Excel 文件名为 `Input.xlsx`，sheet 名为 `Sheet1`，输出会生成到 `output/Input/` 子目录：

- V1 或 V2 普通文件：
  - `output/Input/Sheet1_Rx.ffs`
  - `output/Input/Sheet1_Tx.ffs`
- V1/V2 Total 文件：
  - `output/Input/Sheet1_total_Rx.ffs`
  - `output/Input/Sheet1_total_Tx.ffs`
- 转换日志：
  - `log/ota2ffs_conversion_YYYYMMDD_HHMMSS.log`

默认输出目录：

- `output/`

## 示例文件

`samples/` 目录下包含用于验收和回归测试的 Excel 文件：

- `ota_v1_sample.xlsx`：V1 基础样例。
- `ota_v2_sample.xlsx`：V2 基础样例。
- `ota_v1_default_frequency_sample.xlsx`：V1 偏移位置样例，界面不填频率时默认输出 `1e9 Hz`。
- `ota_mixed_multisheet_sample.xlsx`：V1/V2 混合多 sheet 样例，表格内容放在不同起始位置。

## 安装

建议使用 Python 3.11 或更新版本。

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 运行

```bash
python main.py
```

## 测试

```bash
pytest
```

## GitHub Actions 打包

仓库包含 Windows exe 自动构建流程：

- 每次 push 会运行测试并构建 `OTA2FFS.exe`。
- 构建结果会打包为 zip 并上传到 GitHub Actions artifact。
- zip 内包含：
  - `OTA2FFS.exe`
  - `README.md`
  - `使用说明.txt`
  - `samples/` 下的 Excel 示例文件
- push tag 时会自动创建 GitHub Release，并上传同一个 zip 包。

发布示例：

```bash
git tag v0.1.0
git push origin v0.1.0
```
