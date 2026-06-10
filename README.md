# ota2ffs

OTA2FFS Converter 是一个 Python 桌面工具，用于将 OTA 暗室测试 Excel 数据转换为 CST 可导入的 `.ffs` 远场源文件。

## 功能

- 选择一个 Excel 文件。
- 输入频率并选择单位，支持 `MHz`、`GHz`、`KHz`、`Hz`。V2 数据会优先使用 Excel 表格自带频率。
- 读取一个或多个 sheet。
- 自动识别每个 sheet 的 V1 或 V2 数据格式。
- 每个 sheet 独立转换，某个 sheet 失败不会影响其他 sheet。
- 为每个可转换 sheet 输出 RX / TX `.ffs` 文件。
- 输出目录默认使用工具路径下的 `output/`，也可以手动选择。
- 支持打开输出目录。
- 转换日志默认生成到工具路径下的 `log/`。

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
```

头部之后的数据主体为逗号分隔文本，列顺序固定为：

```text
Phi,Theta,Re(E_Theta),Im(E_Theta),Re(E_Phi),Im(E_Phi)
```

V2 sheet 会使用表格中自带的频率，单位固定为 `MHz`。V1 sheet 或没有自带频率的数据会使用界面中输入的频率。

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

V1 sheet 包含两个数据块：

- 第一块：`A1 = Theta`，`B1 = Phi Angle`，`B2 = Theta Angle`。
- `C1` 开始为 Theta 角度。
- `B3` 开始为 Phi 原始角度。
- 第一块交叉区域写入 `Re(E_Theta)`，也就是 FFS 第 3 列。
- 第二块：结构相同，典型位置为 `A30 = Phi`，`B30 = Phi Angle`，`B31 = Theta Angle`。
- 第二块交叉区域写入 `Re(E_Phi)`，也就是 FFS 第 5 列。
- Phi 输出角度为 Excel 中 Phi 原始角度加 `180`。
- Phi 输出范围补齐为 `0~360`。
- Theta 输出范围补齐为 `0~180`。
- 角度步进根据 Excel 实际数据自动识别，不写死为 `15`。

## V2 Excel 格式

V2 sheet 中包含三个表格，每个表格左上角单元格为 `Polarization`：

- `Polarization` 后一个单元格分别为 `Theta`、`Phi`、`Total`。
- 频率位于当前表格 `Polarization` 行中、最后一个 Theta 角度列的上方，例如 `180` 上方的 `2450`。
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

## 输出文件命名

假设 Excel 文件名为 `Input.xlsx`，sheet 名为 `Sheet1`，输出会生成到 `output/Input/` 子目录：

- V1 或 V2 普通文件：
  - `output/Input/Sheet1_Rx.ffs`
  - `output/Input/Sheet1_Tx.ffs`
- V2 Total 文件：
  - `output/Input/Sheet1_total_Rx.ffs`
  - `output/Input/Sheet1_total_Tx.ffs`
- 转换日志：
  - `log/ota2ffs_conversion_YYYYMMDD_HHMMSS.log`

默认输出目录：

- `output/`

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
