"""trim parse 命令 - 多表格文件解析"""

import re
import warnings
from pathlib import Path
from typing import Optional, Tuple, List
import pandas as pd
from .excel_reader import ExcelReader
from .csv_exporter import CSVExporter


def parse_cell_range(range_str: str) -> Tuple[int, int, int, int]:
    """解析单元格范围字符串，如 "B1:H2"
    
    Args:
        range_str: 范围字符串，格式如 "B1:H2" 或 "A1"
    
    Returns:
        (start_col, start_row, end_col, end_row)
    """
    if ":" in range_str:
        start, end = range_str.split(":")
        start_col, start_row = parse_cell_address(start)
        end_col, end_row = parse_cell_address(end)
    else:
        start_col, start_row = parse_cell_address(range_str)
        end_col, end_row = start_col, start_row
    
    return start_col, start_row, end_col, end_row


def parse_cell_address(address: str) -> Tuple[int, int]:
    """解析单元格地址，如 "B1"
    
    Args:
        address: 单元格地址
    
    Returns:
        (column, row)
    """
    match = re.match(r"([A-Z]+)(\d+)", address.upper())
    if not match:
        raise ValueError(f"无效的单元格地址: {address}")
    
    col_str, row_str = match.groups()
    col = column_letter_to_number(col_str)
    row = int(row_str)
    
    return col, row


def column_letter_to_number(col: str) -> int:
    """将列字母转换为数字，如 A->1, B->2, AA->27"""
    result = 0
    for char in col.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def column_number_to_letter(num: int) -> str:
    """将数字转换为列字母，如 1->A, 2->B, 27->AA"""
    result = ""
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _get_merged_cell_value(ws, row: int, col: int, strip=True):
    """获取单元格值，处理合并单元格的情况
    
    Args:
        ws: worksheet对象
        row: 行号
        col: 列号
        strip: 是否去除前后空格
    
    Returns:
        单元格值，如果是合并单元格则返回合并区域的值
    """
    cell = ws.cell(row=row, column=col)
    
    # 检查是否在合并范围内
    for merged_range in ws.merged_cells.ranges:
        if row in range(merged_range.min_row, merged_range.max_row + 1):
            if col in range(merged_range.min_col, merged_range.max_col + 1):
                # 返回合并区域左上角的单元格值
                top_left_cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                val = top_left_cell.value
                if val and strip:
                    val = str(val).strip()
                return val
    
    val = cell.value
    if val and strip:
        val = str(val).strip()
    return val


def parse_excel_with_axis(
    file_path: str,
    output_dir: str = ".",
    haxis: Optional[str] = None,
    vaxis: Optional[str] = None,
    merge: bool = False,
) -> List[str]:
    """使用行列轴解析Excel文件
    
    Args:
        file_path: Excel文件路径
        output_dir: 输出目录
        haxis: 列标题范围，如 "B1:H2"
        vaxis: 行标题范围，如 "A2:B10"
        merge: 是否合并模式（所有sheet合并到一个文件，sheet名作为行标识）
    
    Returns:
        导出的CSV文件路径列表
    """
    reader = ExcelReader(file_path)
    exporter = CSVExporter(output_dir)
    
    try:
        output_files = []
        
        all_dfs = []
        
        for sheet_name in reader.get_sheet_names():
            ws = reader.workbook[sheet_name]
            
            # 解析轴范围
            h_start_col, h_start_row, h_end_col, h_end_row = (1, 1, 0, 0)
            v_start_col, v_start_row, v_end_col, v_end_row = (1, 1, 0, 0)
            
            if haxis:
                h_start_col, h_start_row, h_end_col, h_end_row = parse_cell_range(haxis)
            
            if vaxis:
                v_start_col, v_start_row, v_end_col, v_end_row = parse_cell_range(vaxis)
            
            # 确定数据区域
            if haxis and vaxis:
                # 两者都有：数据区域是行列标题范围的交叉区域
                data_start_row = v_start_row
                data_start_col = h_start_col
                data_end_row = v_end_row
                data_end_col = h_end_col
            elif haxis:
                # 只有列标题：数据在列标题下方，同一列范围内
                data_start_row = h_end_row + 1
                data_start_col = h_start_col
                data_end_row = ws.max_row
                data_end_col = h_end_col
            elif vaxis:
                # 只有行标题：数据在行标题右侧，同一行范围内
                data_start_row = v_start_row
                data_start_col = v_end_col + 1
                data_end_row = v_end_row
                data_end_col = ws.max_column
            else:
                # 都没有：读取整个sheet
                data_start_row = ws.min_row
                data_start_col = ws.min_column
                data_end_row = ws.max_row
                data_end_col = ws.max_column
            
            # 读取数据
            data = []
            for row in range(data_start_row, data_end_row + 1):
                row_data = []
                for col in range(data_start_col, data_end_col + 1):
                    val = _get_merged_cell_value(ws, row, col, strip=False)
                    # 非数字的单元格置空
                    if val is not None and not isinstance(val, (int, float)):
                        val = None
                    row_data.append(val)
                data.append(row_data)
            
            if not data:
                continue
            
            # 读取并合并列标题
            merged_col_headers = []
            if haxis:
                for col_idx in range(h_start_col, h_end_col + 1):
                    parts = []
                    for row_idx in range(h_start_row, h_end_row + 1):
                        val = _get_merged_cell_value(ws, row_idx, col_idx)
                        if val:
                            parts.append(val)
                    merged_col_headers.append("_".join(parts) if parts else f"Col_{col_idx}")
            else:
                from openpyxl.utils import get_column_letter
                for col in range(data_start_col, data_end_col + 1):
                    merged_col_headers.append(get_column_letter(col))
            
            # 读取并合并行标题
            if vaxis:
                row_headers = []
                for row_idx in range(v_start_row, v_end_row + 1):
                    parts = []
                    for col_idx in range(v_start_col, v_end_col + 1):
                        val = _get_merged_cell_value(ws, row_idx, col_idx)
                        if val:
                            parts.append(val)
                    row_headers.append("_".join(parts) if parts else f"Row_{row_idx}")
            else:
                row_headers = None
            
            # 创建DataFrame
            if merge:
                # 合并模式：每个sheet一行数据
                # 每行数据是一个成本项目，每列是一个指标
                # 列标题格式：行标题|列标题
                num_data_rows = len(data)
                num_data_cols = len(data[0]) if data else 0
                
                # 构建新的列标题：行标题|列标题
                new_col_headers = []
                for row_idx in range(num_data_rows):
                    row_header = row_headers[row_idx] if row_headers and row_idx < len(row_headers) else f"Row_{row_idx}"
                    for col_idx in range(num_data_cols):
                        col_header = merged_col_headers[col_idx] if col_idx < len(merged_col_headers) else f"Col_{col_idx}"
                        new_col_header = f"{row_header}|{col_header}"
                        new_col_headers.append(new_col_header)
                
                # 构建行数据：每列对应一个单元格值（第一个非空值）
                row_data = []
                for row_idx in range(num_data_rows):
                    for col_idx in range(num_data_cols):
                        val = None
                        for r_idx in range(row_idx, row_idx + 1):  # 只取对应行的值
                            v = data[r_idx][col_idx] if col_idx < len(data[r_idx]) else None
                            if v is not None and not pd.isna(v):
                                val = v
                                break
                        row_data.append(val)
                
                df = pd.DataFrame([row_data], columns=new_col_headers)
                df.insert(0, "sheet_name", sheet_name)
                all_dfs.append(df)
            else:
                # 普通模式：每个sheet一个文件
                df = pd.DataFrame(data, columns=merged_col_headers)
                
                # 添加行标题列
                if row_headers:
                    repeat_count = (len(df) + len(row_headers) - 1) // len(row_headers)
                    extended_headers = row_headers * repeat_count
                    df.insert(0, "row_header", extended_headers[:len(df)])
                
                # 导出CSV
                safe_sheet_name = re.sub(r'[\\/*?:"<>|]', "_", sheet_name)
                output_path = exporter.export(df, f"{safe_sheet_name}.csv")
                output_files.append(output_path)
        
        # 合并模式：导出合并后的文件
        if merge and all_dfs:
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=FutureWarning)
                merged_df = pd.concat(all_dfs, ignore_index=True)
            safe_file_name = re.sub(r'[\\/*?:"<>|]', "_", Path(file_path).stem)
            output_path = exporter.export(merged_df, f"{safe_file_name}_merged.csv")
            output_files.append(output_path)
        
        return output_files
    
    finally:
        reader.close()
