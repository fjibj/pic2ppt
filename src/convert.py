#!/usr/bin/env python
"""
Pic2Shape 转换器 - 命令行工具
支持 SVG → PPTX 和 PNG/JPG → SVG → PPTX
"""

import sys
import os
import argparse
from pathlib import Path

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from svg_converter import SVGToPPTXConverter


def convert_svg(svg_file: str, output_path: str = None, **kwargs) -> str:
    """转换 SVG 到 PPTX"""
    converter = SVGToPPTXConverter(svg_file)
    return converter.convert(output_path=output_path, **kwargs)


def convert_png(image_file: str, output_path: str = None, api_key: str = None,
                provider: str = None, keep_svg: bool = True) -> str:
    """转换 PNG/JPG 到 PPTX（通过 AI 生成 SVG）"""
    from pipeline import PNGToPPTXPipeline
    pipeline = PNGToPPTXPipeline(provider=provider, api_key=api_key)
    return pipeline.convert(image_file, output_pptx=output_path, keep_svg=keep_svg)


def main():
    parser = argparse.ArgumentParser(
        description='将 SVG/PNG/JPG 转换为 PowerPoint 可编辑形状',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # SVG 转 PPTX
  python convert.py diagram.svg
  python convert.py diagram.svg -o output.pptx

  # PNG/JPG 转 PPTX（需要 API Key）
  python convert.py diagram.png
  python convert.py diagram.jpg --api-key sk-xxx
  python convert.py diagram.png --provider kimi    # 使用 Kimi
  python convert.py diagram.png --provider glm     # 使用 智谱 GLM

  # 指定幻灯片尺寸（仅 SVG）
  python convert.py diagram.svg -W 16 -H 9
        """
    )

    parser.add_argument('input_file', help='输入文件路径 (SVG/PNG/JPG)')
    parser.add_argument('-o', '--output', help='输出PPTX文件路径（默认与输入同名）')
    parser.add_argument('-W', '--width', type=float, default=13.333,
                       help='幻灯片宽度（英寸，默认13.333，仅 SVG）')
    parser.add_argument('-H', '--height', type=float, default=10.0,
                       help='幻灯片高度（英寸，默认10.0，仅 SVG）')
    parser.add_argument('-m', '--margin', type=float, default=0.5,
                       help='边距（英寸，默认0.5，仅 SVG）')
    parser.add_argument('--api-key', help='API Key（用于 PNG/JPG 转换，默认从环境变量读取）')
    parser.add_argument('--provider', choices=['claude', 'kimi', 'glm', 'openai'],
                       help='AI 提供商 (claude/kimi/glm/openai，默认自动检测)')
    parser.add_argument('--no-keep-svg', action='store_true',
                       help='PNG/JPG 转换后不保留中间 SVG 文件')

    args = parser.parse_args()

    input_path = Path(args.input_file)

    # 检查输入文件
    if not input_path.exists():
        print(f"[FAIL] 错误: 文件不存在: {args.input_file}")
        sys.exit(1)

    # 根据文件类型选择转换方式
    suffix = input_path.suffix.lower()

    try:
        if suffix == '.svg':
            # SVG 直接转换
            print(f"[INFO] 转换 SVG: {input_path.name}")
            result = convert_svg(
                str(input_path),
                output_path=args.output,
                slide_width=args.width,
                slide_height=args.height,
                margin=args.margin
            )
            print(f"[OK] 转换成功: {result}")

        elif suffix in ['.png', '.jpg', '.jpeg', '.bmp', '.gif']:
            # PNG/JPG 需要 AI 转换
            print(f"[INFO] 转换图片: {input_path.name}")
            print("[INFO] 需要 AI 生成 SVG，请确保已设置 API Key 环境变量")
            result = convert_png(
                str(input_path),
                output_path=args.output,
                api_key=args.api_key,
                provider=args.provider,
                keep_svg=not args.no_keep_svg
            )
            print(f"[OK] 转换成功: {result}")

        else:
            print(f"[FAIL] 错误: 不支持的文件格式: {suffix}")
            print("支持的格式: .svg, .png, .jpg, .jpeg, .bmp, .gif")
            sys.exit(1)

    except Exception as e:
        print(f"[FAIL] 转换失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
