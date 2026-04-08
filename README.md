# City Center PDF to Excel

这个项目会读取根目录下的 `cityCenter.pdf`，提取 PDF 中的城市中心点经纬度，并生成一个 Excel 文件：`city_center_output.xlsx`。

## 环境要求

- Python `3.12`
- `uv`

## 安装依赖

```bash
uv sync
```

## 运行方式

```bash
uv run python main.py
```

## 输出文件

运行完成后会在项目根目录生成：

- `city_center_output.xlsx`

其中包含两个工作表：

- `Sheet1`
  - 仅包含指定城市
  - 列顺序为：`城市`、`经度`、`纬度`、`瓦片`、`UTM`
  - 如果指定城市不在 PDF 提取结果中，则该行保留城市名，其余列留空

- `Sheet2`
  - 包含从 `cityCenter.pdf` 中实际提取出的全部条目
  - 列顺序为：`城市`、`经度`、`纬度`、`瓦片`、`UTM`

## 字段规则

- `UTM`
  - 按 `WGS 1984 UTM Zone` 的分区号计算
  - 输出格式为 `分区号 + 半球`，例如：`49N`

- `瓦片`
  - 使用 `5° × 5°` 瓦片命名
  - 格式为：`{e/w}{lon_min}_{n/s}{lat_max}_{e/w}{lon_max}_{n/s}{lat_min}`
  - 例如西安：`e105_n35_e110_n30`

## 样式说明

生成的 Excel 工作簿统一使用字体：`等线`。
