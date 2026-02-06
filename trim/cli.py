"""trim命令行入口"""

import sys
from pathlib import Path
from typing import Tuple
import click
from .info import get_file_info, print_file_info
from .parse import parse_excel_with_axis
from .splice import splice_with_headers


@click.group()
def main():
    """Excel文件处理工具"""
    pass


@main.command()
@click.argument("file_path", type=click.Path(exists=True))
@click.option("--json", "-j", is_flag=True, help="以JSON格式输出")
def info(file_path: str, json: bool):
    """提取表格文件的统计信息
    
    FILE_PATH: Excel文件路径
    """
    if json:
        result = get_file_info(file_path, output_json=True)
        click.echo(result)
    else:
        print_file_info(file_path)


@main.command()
@click.argument("file_path", type=click.Path(exists=True))
@click.option("-h", "--haxis", help="列标题范围，如 B1:H2")
@click.option("-v", "--vaxis", help="行标题范围，如 A2:B10")
@click.option("-p", "--path", "output_dir", default=".", help="输出目录")
@click.option("-m", "--merge", is_flag=True, help="合并模式：列标题格式为 原行标题|原列标题")
@click.option("-t", "--timestamp", default=None, help="合并模式：会计期间值，如 -t 2512")
@click.option("-n", "--name", default=None, help="合并模式：数据集名称值，如 -n CPGC")
def parse(file_path: str, haxis: str, vaxis: str, output_dir: str, merge: bool, timestamp: str, name: str):
    """多表格文件解析，并生成csv文件
    
    FILE_PATH: Excel文件路径
    
    合并模式说明：
    - 第一列：企业全称（从sheet名字解析）
    - 第二列：企业简称（从映射表查找）
    - 第三列：企业ID（从映射表查找）
    - 第四列：数据集名称（配合 -n/--name 参数使用）
    - 第五列：会计期间（配合 -t/--timestamp 参数使用）
    - 第六列：报表类型（括号缩写：BB/HB/CE/GL）
    - 后续列：原行标题|原列标题 格式的数据列
    - 合并后每个sheet一行数据
    """
    output_dir = Path(output_dir)
    click.echo(f"解析文件: {file_path}")
    
    if haxis:
        click.echo(f"列标题范围: {haxis}")
    if vaxis:
        click.echo(f"行标题范围: {vaxis}")
    if merge:
        click.echo("合并模式: sheet名作为行标题，列标题格式为 原行标题|原列标题")
    if timestamp:
        click.echo(f"会计期间: {timestamp}")
    if name:
        click.echo(f"数据集名称: {name}")
    
    output_files = parse_excel_with_axis(
        file_path,
        output_dir=str(output_dir),
        haxis=haxis,
        vaxis=vaxis,
        merge=merge,
        timestamp=timestamp,
        name=name,
    )
    
    click.echo(f"\n生成 {len(output_files)} 个文件:")
    for f in output_files:
        click.echo(f"  - {f}")


@main.command()
@click.argument("file_paths", nargs=-1, type=click.Path(exists=True))
@click.option("-r", "--range", "frozen_range", default="A:A", help="冻结列范围，如 A:D（作为行唯一标识）")
@click.option("-s", "--skip", "skip_range", default=None, help="跳过列范围，如 E:E（这些列不包含在输出中）")
@click.option("-o", "--output", "output_file", default="merged.csv", help="输出文件名")
@click.option("-p", "--path", "output_dir", default=".", help="输出目录")
@click.option("-H", "--headers", type=int, default=1, help="表头行数")
def splice(file_paths: Tuple[str], frozen_range: str, skip_range: str, output_file: str, output_dir: str, headers: int):
    """拼接多个文件的数据
    
    用法:
      trim splice file1.csv file2.csv -r A:D -s E:E -H 1 -p ./output -o merged.csv
    """
    # 合并所有文件路径
    all_files = list(file_paths)
    
    if not all_files:
        click.echo("错误：请指定至少一个文件")
        sys.exit(1)
    
    output_dir = Path(output_dir)
    click.echo(f"拼接文件: {len(all_files)} 个")
    click.echo(f"冻结列范围: {frozen_range}")
    if skip_range:
        click.echo(f"跳过列范围: {skip_range}")
    click.echo(f"表头行数: {headers}")
    click.echo(f"输出目录: {output_dir}")
    
    output_path = splice_with_headers(
        all_files,
        frozen_range=frozen_range,
        header_rows=headers,
        output_dir=str(output_dir),
        output_name=output_file,
        skip_range=skip_range,
    )
    
    click.echo(f"\n✅ 已生成文件: {output_path}")


if __name__ == "__main__":
    main()
