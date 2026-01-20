"""trim info 命令 - 提取表格文件的统计信息"""

import json
from pathlib import Path
from typing import Optional
from .excel_reader import ExcelReader


def get_file_info(file_path: str, output_json: bool = False) -> dict:
    """获取Excel文件的统计信息
    
    Args:
        file_path: Excel文件路径
        output_json: 是否以JSON格式输出
    
    Returns:
        统计信息字典
    """
    reader = ExcelReader(file_path)
    
    try:
        sheets_info = reader.get_all_sheets_info()
        
        result = {
            "file": str(Path(file_path).resolve()),
            "file_name": Path(file_path).name,
            "sheet_count": len(sheets_info),
            "sheets": sheets_info,
        }
        
        if output_json:
            return json.dumps(result, ensure_ascii=False, indent=2)
        
        return result
    finally:
        reader.close()


def print_file_info(file_path: str):
    """打印文件统计信息"""
    result = get_file_info(file_path)
    
    print(f"\n文件: {result['file_name']}")
    print(f"路径: {result['file']}")
    print(f"Sheet数量: {result['sheet_count']}")
    print("\nSheet详情:")
    print("-" * 50)
    
    for sheet in result['sheets']:
        print(f"  名称: {sheet['name']}")
        print(f"  行数: {sheet['rows']}")
        print(f"  列数: {sheet['columns']}")
        print(f"  使用范围: {sheet['used_range']}")
        print()
