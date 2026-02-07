"""splice命令 - 多文件首尾拼接"""

import re
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional, Union
import pandas as pd
from .excel_reader import ExcelReader
from .csv_exporter import CSVExporter


def is_csv_file(file_path: str) -> bool:
    """判断是否为CSV文件"""
    return file_path.lower().endswith('.csv')


def get_file_extension(file_path: str) -> str:
    """获取文件扩展名（小写）"""
    return Path(file_path).suffix.lower()


def parse_frozen_range(range_str: str) -> Tuple[int, int]:
    """解析冻结列范围，如 "A:E"
    
    Args:
        range_str: 范围字符串，格式如 "A:E" 或 "A"
    
    Returns:
        (start_col, end_col) 列号（从1开始）
    """
    if ":" in range_str:
        start, end = range_str.split(":")
        start_col = column_letter_to_number(start)
        end_col = column_letter_to_number(end)
    else:
        start_col = end_col = column_letter_to_number(range_str)
    
    return start_col, end_col


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


def _get_cell_value(ws, row: int, col: int) -> Any:
    """获取单元格值"""
    cell = ws.cell(row=row, column=col)
    return cell.value


def splice_with_headers(
    file_paths: List[str],
    frozen_range: str,
    header_rows: int = 1,
    output_dir: str = ".",
    output_name: str = "spliced_output.csv",
    skip_range: Optional[str] = None,
) -> str:
    """拼接多个文件的数据（带表头）
    
    正确逻辑：输出列 = 冻结列 + 文件1的数据列 + 文件2的数据列 + ... + 文件n的数据列
    
    Args:
        file_paths: 文件路径列表（支持Excel和CSV）
        frozen_range: 冻结列范围，如 "A:E"
        header_rows: 表头行数
        output_dir: 输出目录
        output_name: 输出文件名
        skip_range: 跳过列范围，如 "E:E"（这些列不包含在输出中）
    
    Returns:
        输出文件的完整路径
    """
    # 解析冻结列范围
    frozen_start, frozen_end = parse_frozen_range(frozen_range)
    
    # 解析跳过列范围
    skip_cols: set = set()
    if skip_range:
        skip_start, skip_end = parse_frozen_range(skip_range)
        for col in range(skip_start, skip_end + 1):
            skip_cols.add(col)
    
    # 存储所有数据：{唯一标识: {file_idx: data_row}}
    data_map: Dict[str, Dict[int, Dict[str, Any]]] = {}
    
    # 存储每个文件的列标题列表
    file_headers: List[List[str]] = []
    
    for file_idx, file_path in enumerate(file_paths):
        ext = get_file_extension(file_path)
        
        if ext == '.csv':
            # CSV文件使用pandas读取
            df = pd.read_csv(file_path, header=None, dtype=str)
            
            # 获取表头（第一行数据行）- 同时记录列索引
            data_cols: List[Tuple[int, str]] = []  # (col_idx, header)
            for col_idx in range(frozen_end, len(df.columns)):
                if col_idx + 1 in skip_cols:
                    continue
                header = df.iloc[0, col_idx] if col_idx < len(df.columns) else f"Col_{col_idx + 1}"
                data_cols.append((col_idx, str(header) if header and str(header).strip() else f"Col_{col_idx + 1}"))
            
            headers = [h for _, h in data_cols]
            file_headers.append(headers)
            
            # 读取数据行（从第二行开始）
            for row_idx in range(1, len(df)):
                # 构建唯一标识 - 从冻结列起始列开始
                key_parts = []
                for col_idx in range(frozen_start - 1, frozen_end):
                    val = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
                    if val is not None and str(val).strip():
                        key_parts.append(str(val).strip())
                    else:
                        key_parts.append("")
                
                if not any(key_parts):
                    continue
                
                key = "_".join(key_parts)
                
                # 读取数据列 - 使用记录的实际列索引
                data_row: Dict[str, Any] = {}
                for col_idx, header in data_cols:
                    val = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
                    if val is not None and str(val).strip():
                        try:
                            float(val)
                            data_row[header] = val
                        except (ValueError, TypeError):
                            data_row[header] = ""
                    else:
                        data_row[header] = ""
                
                data_row["_source_file"] = Path(file_path).name
                data_row["_source_sheet"] = "CSV"
                
                if key not in data_map:
                    data_map[key] = {}
                data_map[key][file_idx] = data_row
        
        else:
            # Excel文件使用ExcelReader读取
            reader = ExcelReader(file_path)
            
            try:
                for sheet_name in reader.get_sheet_names():
                    ws = reader.workbook[sheet_name]
                    
                    min_row = ws.min_row
                    max_row = ws.max_row
                    min_col = ws.min_column
                    max_col = ws.max_column
                    
                    # 读取表头 - 记录列索引
                    data_cols = []
                    for col in range(frozen_end + 1, max_col + 1):
                        if col in skip_cols:
                            continue
                        val = _get_cell_value(ws, min_row, col)
                        data_cols.append((col, str(val) if val else f"Col_{col}"))
                    
                    headers = [h for _, h in data_cols]
                    file_headers.append(headers)
                    
                    # 读取数据行
                    for row in range(min_row + header_rows, max_row + 1):
                        # 构建唯一标识
                        key_parts = []
                        for col in range(frozen_start, frozen_end + 1):
                            val = _get_cell_value(ws, row, col)
                            if val is not None:
                                key_parts.append(str(val).strip())
                            else:
                                key_parts.append("")
                        
                        if not any(key_parts):
                            continue
                        
                        key = "_".join(key_parts)
                        
                        # 读取数据列 - 使用记录的实际列索引
                        data_row: Dict[str, Any] = {}
                        for col, header in data_cols:
                            val = _get_cell_value(ws, row, col)
                            if val is not None and str(val).strip():
                                if not isinstance(val, (int, float)):
                                    data_row[header] = ""
                                else:
                                    data_row[header] = val
                            else:
                                data_row[header] = ""
                        
                        data_row["_source_file"] = Path(file_path).name
                        data_row["_source_sheet"] = sheet_name
                        
                        if key not in data_map:
                            data_map[key] = {}
                        data_map[key][file_idx] = data_row
            
            finally:
                reader.close()
    
    # 构建输出数据
    sorted_keys = sorted(data_map.keys())
    
    # 构建输出列头
    output_columns = []
    for col_idx in range(frozen_start, frozen_end + 1):
        output_columns.append(column_number_to_letter(col_idx))
    for headers in file_headers:
        for header in headers:
            output_columns.append(header)
    
    # 构建行数据
    output_data = []
    for key in sorted_keys:
        file_data = data_map[key]
        
        key_parts = key.split("_")
        row_list = []
        
        # 冻结列
        for i, part in enumerate(key_parts):
            row_list.append(part)
        
        # 检查所有数据列是否全为空
        has_data = False
        data_values = []
        
        # 每个文件的数据列
        for file_idx in range(len(file_paths)):
            for header in file_headers[file_idx]:
                if file_idx in file_data:
                    value = file_data[file_idx].get(header, "")
                    data_values.append(value)
                    if value and str(value).strip():
                        has_data = True
                else:
                    data_values.append("")
        
        # 如果所有数据列都为空，跳过该行
        if not has_data:
            continue
        
        row_list.extend(data_values)
        output_data.append(row_list)
    
    df = pd.DataFrame(output_data, columns=output_columns)
    
    # 导出
    exporter = CSVExporter(output_dir)
    output_path = exporter.export(df, output_name)
    
    return output_path
