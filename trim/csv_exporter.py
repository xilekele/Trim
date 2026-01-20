"""CSV文件导出工具"""

from pathlib import Path
from typing import List, Optional
import pandas as pd


class CSVExporter:
    """CSV文件导出器"""
    
    def __init__(self, output_dir: str = "."):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def export(
        self,
        df: pd.DataFrame,
        filename: str,
        index: bool = False,
        encoding: str = "utf-8-sig",
    ) -> str:
        """导出DataFrame到CSV文件
        
        Args:
            df: DataFrame对象
            filename: 文件名
            index: 是否导出索引
            encoding: 编码格式
        
        Returns:
            导出文件的完整路径
        """
        output_path = self.output_dir / filename
        df.to_csv(output_path, index=index, encoding=encoding)
        return str(output_path)
    
    def export_with_prefix(
        self,
        df: pd.DataFrame,
        prefix: str,
        index: bool = False,
        encoding: str = "utf-8-sig",
    ) -> str:
        """导出DataFrame，使用前缀命名
        
        Args:
            df: DataFrame对象
            prefix: 文件前缀
            index: 是否导出索引
            encoding: 编码格式
        
        Returns:
            导出文件的完整路径
        """
        filename = f"{prefix}.csv"
        return self.export(df, filename, index, encoding)
    
    def export_multiple(
        self,
        data: dict,
        encoding: str = "utf-8-sig",
    ) -> List[str]:
        """导出多个DataFrame到CSV文件
        
        Args:
            data: {sheet_name: DataFrame} 字典
            encoding: 编码格式
        
        Returns:
            导出文件路径列表
        """
        output_paths = []
        for name, df in data.items():
            safe_name = self._sanitize_filename(name)
            path = self.export(df, f"{safe_name}.csv", encoding=encoding)
            output_paths.append(path)
        return output_paths
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名中的非法字符"""
        import re
        # 移除或替换Windows文件名中的非法字符
        return re.sub(r'[\\/*?:"<>|]', "_", filename)
    
    def merge_csv_files(
        self,
        file_paths: List[str],
        output_filename: str,
        sheet_column: str = "sheet_name",
        merge_headers: bool = True,
    ) -> str:
        """合并多个CSV文件
        
        Args:
            file_paths: CSV文件路径列表
            output_filename: 输出文件名
            sheet_column: 用于标识sheet来源的列名
            merge_headers: 是否将原行列标题合并为新列标题
        
        Returns:
            合并后的文件路径
        """
        all_data = []
        
        for file_path in file_paths:
            df = pd.read_csv(file_path)
            
            # 从文件名提取sheet名称
            sheet_name = Path(file_path).stem
            
            # 添加sheet名称作为新列
            if sheet_column:
                df.insert(0, sheet_column, sheet_name)
            
            all_data.append(df)
        
        # 合并所有数据
        merged_df = pd.concat(all_data, ignore_index=True)
        
        # 导出合并后的文件
        output_path = self.output_dir / output_filename
        merged_df.to_csv(output_path, index=False, encoding="utf-8-sig")
        
        return str(output_path)
