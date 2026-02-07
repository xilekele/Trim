# Trim
xls to csv - Excel文件处理工具

## Cline
开发一个名为Trim的项目，源代码放在trim目录下。trim命令行，具备以下功能：
1. `trim info`提取表格文件的统计信息。
2. `trim parse`多表格文件解析，并生成csv文件。在文件解析过程中如果同时提供行范围与列范围，程序应自动计算出数据区域，例如：`-h "D3:U4" -v "A6:B25"`参数制定了行与列的标题范围，那么数据区域就应该是`"D6:U25"`，不是数字的单元格置空。
3. `-h --haxis "B1:H2"`、`-v --vaxis`可指定行/列标题范围，并自动合并成新的行/列标题用_连接。特别强调如果遇到合并单元格的情况，要进行逐项合并，标题前后不能有空行。
4. 提供`-m --merge`参数，新的列标题的格式为原行标题|原列标题，将sheet的名字做解析（第一列为企业全称、第二列为企业简称、第三列为企业ID、新增`-n --name "CPGC"` 参数用来标记第四列数据集的内容、新增`-t --timestamp "2512"` 参数用来标记第五列的内容、第六列为sheet名字括号里的内容缩写，具体为BB/HB/CE/XE），合并后一个sheet一行数据。`企业全称：企业简称：企业ID`三者关系见表。
5. 

## 功能

1. **`trim info`** - 提取表格文件的统计信息
2. **`trim parse`** - 多表格文件解析，并生成csv文件
3. **`-h --haxis`** - 指定列标题范围，自动合并成新的列标题
4. **`-v --vaxis`** - 指定行标题范围，自动合并成新的行标题
5. **`trim merge`** - 合并csv文件，sheet名字作为新的行标题

## 安装

```bash
pip install -e .
```

## 使用方法

### 提取文件信息

```bash
# 查看文件统计信息
trim info input.xlsx

# 以JSON格式输出
trim info input.xlsx --json
```

### 文件解析

```bash
# 基本解析
trim parse input.xlsx

# 指定列标题范围
trim parse input.xlsx -h "B1:H2"

# 指定行标题范围
trim parse input.xlsx -v "A2:B10"

# 同时指定行和列标题范围
trim parse input.xlsx -h "B1:H2" -v "A2:B10"

# 指定输出目录
trim parse input.xlsx -p /path/to/output
```

### 文件合并

```bash
# 合并多个CSV文件
trim merge file1.csv file2.csv file3.csv

# 合并Excel文件的所有sheet
trim merge input.xlsx

# 指定输出目录和文件名
trim merge file1.csv file2.csv -p /path/to/output -o merged_result
```

## 参数说明

### haxis (水平轴/列标题)

指定列标题的范围，格式如 `B1:H2`，表示从B1到H2的矩形区域。范围内的单元格内容会按列合并为新的列标题，多行标题之间用 ` | ` 分隔。

### vaxis (垂直轴/行标题)

指定行标题的范围，格式如 `A2:B10`，表示从A2到B10的矩形区域。范围内的单元格内容会按行合并为新的行标题，多列标题之间用 ` | ` 分隔。

## 示例

假设有一个Excel文件 `data.xlsx`，包含以下结构：

```
    | B       C       D       E
----+--------+--------+--------+--------
1   | 类别1   类别1   类别2   类别2
2   | 项目A  项目B  项目A  项目B
3   | 100    200    300    400
4   | 150    250    350    450
```

使用以下命令解析：

```bash
trim parse data.xlsx -h "B1:E2"
```

输出CSV文件，列标题会自动合并为：
- `类别1 | 项目A`
- `类别1 | 项目B`
- `类别2 | 项目A`
- `类别2 | 项目B`


实际例子


```bash
trim parse .trash\快报财务指标表.xlsx -h "A4:B291" -v "C3:G3" -p .\.trash\ -n "KBCW" -t 202601 -m
trim parse .trash\快报业务指标表.xlsx -h "A4:B180" -v "C3:G3" -p .\.trash\ -n "KBYW" -t 202601 -m
trim parse .trash\JTCBB_CBB01.xlsx -h ":" -v ":" -p .\.trash\ -n "" -t 202601 -m
trim parse .trash\JTCBB_CBB02矿山作业成本项目构成表.xlsx -h "C3:Z4" -v "A6:B61" -p .\.trash\ -n "KSZY" -t 202601 -m
trim parse .trash\JTCBB_CBB03矿山成本要素表.xlsx -h "C3:P4" -v "A6:B58" -p .\.trash\ -n "KSYS" -t 202601 -m
trim parse .trash\JTCBB_CBB04矿山企业产品综合成本构成表.xlsx -h "D3:U4" -v "A6:C25" -p .\.trash\ -n "KSCP" -t 202601 -m
trim parse .trash\JTCBB_CBB05矿山企业制造费用明细表.xlsx -h "C3:R5" -v "A7:B64" -p .\.trash\ -n "KSZZ" -t 202601 -m
trim parse .trash\JTCBB_CBB06外委作业成本表.xlsx -h "D3:U4" -v "A6:C32" -p .\.trash\ -n "WW" -t 202601 -m
trim parse .trash\JTCBB_CBB07定额材料、动力消耗统计表.xlsx -h "D3:AA4" -v "A6:C52" -p .\.trash\ -n "CLDL" -t 202601 -m
trim parse .trash\JTCBB_CBB08冶炼企业吨矿加工成本表.xlsx -h "D3:J3" -v "A5:C11" -p .\.trash\ -n "YLDK" -t 202601 -m
trim parse .trash\JTCBB_CBB09冶炼生产指标表.xlsx -h "D3:J3" -v "A5:C28" -p .\.trash\ -n "YLSC" -t 202601 -m
trim parse .trash\JTCBB_CBB10冶炼企业产品成本表.xlsx -h "D3:O4" -v "A6:C33" -p .\.trash\ -n "YLCP" -t 202601 -m
trim parse .trash\JTCBB_CBB11冶炼企业作业成本表.xlsx -h "D3:AI4" -v "A6:C16" -p .\.trash\ -n "YLZY" -t 202601 -m
trim parse .trash\JTCBB_CBB12乌山堆浸制造费用明细表.xlsx -h "C3:N5" -v "A7:B81" -p .\.trash\ -n "WSDJ" -t 202601 -m
trim parse .trash\JTCBB_CBB13冶炼成本要素表.xlsx -h "C3:H4" -v "A6:B58" -p .\.trash\ -n "YLYS" -t 202601 -m
```


## 项目结构

```
Trim/
├── pyproject.toml          # 项目配置
├── README.md               # 项目说明
├── LICENSE                 # MIT许可证
├── .gitignore              # Git忽略配置
└── trim/
    ├── __init__.py         # 包初始化
    ├── cli.py              # 命令行入口
    ├── excel_reader.py     # Excel读取工具
    ├── csv_exporter.py     # CSV导出工具
    ├── info.py             # info命令
    ├── parse.py            # parse命令
    └── merge.py            # merge命令
```
