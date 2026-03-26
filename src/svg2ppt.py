#!/usr/bin/env python
"""
SVG to PPTX 转换命令行工具
支持通配符批量转换

Usage:
    python svg2ppt.py <path>              # 支持单个文件或通配符
    python svg2ppt.py diagram.svg         # 转换单个文件
    python svg2ppt.py *.svg               # 转换所有SVG
    python svg2ppt.py F:\path\to\*.svg    # 指定目录通配符

Output:
    转换后的PPT文件保存在同目录，文件名相同，扩展名改为.pptx
"""

import sys
import os
import glob
import argparse

# 添加 src 到路径
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

try:
    from svg_converter import SVGToPPTXConverter
except ImportError as e:
    print(f"错误: 无法导入 svg_converter 模块: {e}")
    print("请确保在正确的目录运行此脚本")
    sys.exit(1)


def convert_single_file(svg_path: str) -> str:
    """转换单个SVG文件"""
    if not os.path.exists(svg_path):
        print(f"  [SKIP] File not found: {svg_path}")
        return None

    # 生成输出路径（同目录，同名，.pptx扩展名）
    base_name = os.path.splitext(svg_path)[0]
    output_path = base_name + '.pptx'

    print(f"  输入: {svg_path}")
    print(f"  输出: {output_path}")

    try:
        converter = SVGToPPTXConverter(svg_path)
        result = converter.convert(output_path)
        print(f"  [OK] 转换成功: {os.path.basename(result)}")
        return result
    except Exception as e:
        print(f"  [FAIL] 转换失败: {e}")
        return None


def expand_wildcard(pattern: str) -> list:
    """展开通配符路径"""
    # 如果包含通配符，使用glob
    if '*' in pattern or '?' in pattern:
        # 转换Windows路径分隔符
        pattern = pattern.replace('/', os.sep)
        matches = glob.glob(pattern)
        # 过滤出.svg文件
        return [f for f in matches if f.lower().endswith('.svg')]
    else:
        # 单个文件
        return [pattern]


def main():
    parser = argparse.ArgumentParser(
        description='将SVG文件转换为PowerPoint可编辑形状',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python svg2ppt.py diagram.svg              # 转换单个文件
  python svg2ppt.py *.svg                    # 转换当前目录所有SVG
  python svg2ppt.py "F:\\path\\to\\*.svg"      # 指定目录通配符
  python svg2ppt.py diagram1.svg diagram2.svg # 转换多个文件
        """
    )

    parser.add_argument(
        'paths',
        nargs='+',
        help='SVG文件路径，支持通配符（如 *.svg）'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='显示详细输出'
    )

    args = parser.parse_args()

    print("=" * 70)
    print("SVG to PPTX Batch Converter")
    print("=" * 70)
    print()

    # 收集所有要处理的文件
    all_files = []
    for pattern in args.paths:
        files = expand_wildcard(pattern)
        if files:
            all_files.extend(files)
        else:
            print(f"警告: 未找到匹配的文件: {pattern}")

    if not all_files:
        print("\n[ERROR] No SVG files found")
        sys.exit(1)

    # 去重并保持顺序
    seen = set()
    unique_files = []
    for f in all_files:
        if f not in seen:
            unique_files.append(f)
            seen.add(f)

    print(f"找到 {len(unique_files)} 个SVG文件:\n")

    # 转换统计
    success_count = 0
    failed_count = 0

    for i, svg_file in enumerate(unique_files, 1):
        print(f"[{i}/{len(unique_files)}] 处理中...")
        result = convert_single_file(svg_file)
        if result:
            success_count += 1
        else:
            failed_count += 1
        print()

    # 总结
    print("=" * 70)
    print("转换完成!")
    print(f"  成功: {success_count}/{len(unique_files)}")
    print(f"  失败: {failed_count}/{len(unique_files)}")
    print("=" * 70)

    if failed_count > 0:
        sys.exit(1)


if __name__ == '__main__':
    main()
