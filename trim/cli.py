"""trim命令行入口"""

import sys
from pathlib import Path
import click
from .info import get_file_info, print_file_info
from .parse import parse_excel_with_axis


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
@click.option("-m", "--merge", is_flag=True, help="合并模式：sheet名作为行标题，列标题格式为 原行标题|原列标题")
def parse(file_path: str, haxis: str, vaxis: str, output_dir: str, merge: bool):
    """多表格文件解析，并生成csv文件
    
    FILE_PATH: Excel文件路径
    """
    output_dir = Path(output_dir)
    click.echo(f"解析文件: {file_path}")
    
    if haxis:
        click.echo(f"列标题范围: {haxis}")
    if vaxis:
        click.echo(f"行标题范围: {vaxis}")
    if merge:
        click.echo("合并模式: sheet名作为行标题，列标题格式为 原行标题|原列标题")
    
    output_files = parse_excel_with_axis(
        file_path,
        output_dir=str(output_dir),
        haxis=haxis,
        vaxis=vaxis,
        merge=merge,
    )
    
    click.echo(f"\n生成 {len(output_files)} 个文件:")
    for f in output_files:
        click.echo(f"  - {f}")


if __name__ == "__main__":
    main()
