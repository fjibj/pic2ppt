#!/usr/bin/env python
"""
Pic2Shape 批量转换工具
支持 SVG/PNG/JPG → PPTX
"""

import sys
import os
import glob
import argparse
from pathlib import Path

# 添加 src 到路径
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

try:
    from svg_converter import SVGToPPTXConverter
    from pipeline import PNGToPPTXPipeline
except ImportError as e:
    print(f"错误: 无法导入模块: {e}")
    print("请确保在正确的目录运行此脚本")
    sys.exit(1)


def convert_single_file(file_path: str, api_key: str = None, provider: str = None,
                        keep_svg: bool = True) -> str:
    """转换单个文件"""
    if not os.path.exists(file_path):
        print(f"  [SKIP] 文件不存在: {file_path}")
        return None

    path = Path(file_path)
    suffix = path.suffix.lower()

    # 生成输出路径（同目录，同名，.pptx扩展名）
    base_name = os.path.splitext(file_path)[0]
    output_path = base_name + '.pptx'

    print(f"  输入: {file_path}")
    print(f"  输出: {output_path}")

    try:
        if suffix == '.svg':
            # SVG 直接转换
            converter = SVGToPPTXConverter(file_path)
            result = converter.convert(output_path)

        elif suffix in ['.png', '.jpg', '.jpeg', '.bmp', '.gif']:
            # PNG/JPG 需要 AI 转换
            print("  [INFO] 需要 AI 生成 SVG...")
            pipeline = PNGToPPTXPipeline(provider=provider, api_key=api_key)
            result = pipeline.convert(file_path, output_pptx=output_path, keep_svg=keep_svg)

        else:
            print(f"  [SKIP] 不支持的格式: {suffix}")
            return None

        print(f"  [OK] 转换成功: {os.path.basename(result)}")
        return result

    except Exception as e:
        print(f"  [FAIL] 转换失败: {e}")
        return None


def expand_wildcard(pattern: str, extensions: list = None) -> list:
    """展开通配符路径"""
    if extensions is None:
        extensions = ['.svg', '.png', '.jpg', '.jpeg']

    # 如果包含通配符，使用glob
    if '*' in pattern or '?' in pattern:
        # 转换Windows路径分隔符
        pattern = pattern.replace('/', os.sep)
        matches = glob.glob(pattern)
        # 过滤出支持的文件
        return [f for f in matches if Path(f).suffix.lower() in extensions]
    else:
        # 单个文件
        return [pattern]


def main():
    parser = argparse.ArgumentParser(
        description='将 SVG/PNG/JPG 文件批量转换为 PowerPoint',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # SVG 转 PPTX
  python pic2shape.py diagram.svg
  python pic2shape.py *.svg

  # PNG/JPG 转 PPTX（需要 API Key）
  python pic2shape.py diagram.png
  python pic2shape.py *.png --api-key sk-xxx
  python pic2shape.py *.png --provider kimi
  python pic2shape.py *.png --provider glm

  # 批量转换
  python pic2shape.py *.svg *.png
        """
    )

    parser.add_argument(
        'paths',
        nargs='+',
        help='文件路径，支持通配符（如 *.svg, *.png）'
    )
    parser.add_argument(
        '--api-key',
        help='API Key（用于 PNG/JPG 转换，默认从环境变量读取）'
    )
    parser.add_argument(
        '--provider',
        choices=['claude', 'kimi', 'glm', 'openai'],
        help='AI 提供商 (claude/kimi/glm/openai，默认自动检测)'
    )
    parser.add_argument(
        '--no-keep-svg',
        action='store_true',
        help='PNG/JPG 转换后不保留中间 SVG 文件'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='显示详细输出'
    )

    args = parser.parse_args()

    print("=" * 70)
    print("Pic2Shape Batch Converter (SVG/PNG/JPG → PPTX)")
    print("=" * 70)
    print()

    # 支持的扩展名
    supported_ext = ['.svg', '.png', '.jpg', '.jpeg', '.bmp', '.gif']

    # 收集所有要处理的文件
    all_files = []
    for pattern in args.paths:
        files = expand_wildcard(pattern, supported_ext)
        if files:
            all_files.extend(files)
        else:
            print(f"警告: 未找到匹配的文件: {pattern}")

    if not all_files:
        print("\n[ERROR] 未找到可转换的文件")
        print(f"支持的格式: {', '.join(supported_ext)}")
        sys.exit(1)

    # 去重并保持顺序
    seen = set()
    unique_files = []
    for f in all_files:
        if f not in seen:
            unique_files.append(f)
            seen.add(f)

    print(f"找到 {len(unique_files)} 个文件:\n")

    # 转换统计
    success_count = 0
    failed_count = 0
    skipped_count = 0

    for i, file_path in enumerate(unique_files, 1):
        print(f"[{i}/{len(unique_files)}] 处理中...")
        result = convert_single_file(
            file_path,
            api_key=args.api_key,
            provider=args.provider,
            keep_svg=not args.no_keep_svg
        )
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
