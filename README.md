# ota2ffs

OTA2FFS Converter 是一个 Python 桌面工具，用于将 OTA 暗室测试 Excel 数据转换为 CST 可导入的 `.ffs` 远场源文件。

## 功能

- 选择一个 Excel 文件。
- 读取一个或多个 sheet。
- 自动识别每个 sheet 的 V1 或 V2 数据格式。
- 每个 sheet 独立转换，某个 sheet 失败不会影响其他 sheet。
- 为每个可转换 sheet 输出 RX / TX `.ffs` 文件。
- 选择输出目录。
- 生成转换日志。

## FFS 输出格式

工具生成的 `.ffs` 主体为逗号分隔文本，列顺序固定为：

```text
Phi,Theta,Re(E_Theta),Im(E_Theta),Re(E_Phi),Im(E_Phi)
```

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
- 同一行最后一个有效单元格为频率，单位 MHz。
- 下一行格式为 `Phi\Theta, 0, 30, 60, ..., 180`。
- 后续行为 `Phi角度, value1, value2, ...`。

转换规则：

- `Theta` 表格写入 `Re(E_Theta)`。
- `Phi` 表格写入 `Re(E_Phi)`。
- `Theta + Phi` 合成普通 FFS。
- `Total` 表格单独生成拓图 FFS，数据写入 `Re(E_Theta)`，`Re(E_Phi)` 全部为 `0`。

## 输出文件命名

假设 sheet 名为 `Sheet1`：

- V1 或 V2 普通文件：
  - `Sheet1_RX.ffs`
  - `Sheet1_TX.ffs`
- V2 Total 拓图文件：
  - `Sheet1_拓图_RX.ffs`
  - `Sheet1_拓图_TX.ffs`
- 转换日志：
  - `ota2ffs_conversion_YYYYMMDD_HHMMSS.log`

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
