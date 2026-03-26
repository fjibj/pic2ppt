"""
示例：使用新的通用SVG转换器
"""

import sys
import os

# 添加路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from svg_converter import SVGToPPTXConverter, HandlerRegistry


def main():
    # 示例SVG文件列表
    svg_files = [
        "ai_component_management_platform.svg",
        "reinforcement-fine-tuning-flowchart.svg",
    ]

    # 查找存在的文件
    found_files = []
    for f in svg_files:
        if os.path.exists(f):
            found_files.append(f)

    if not found_files:
        print("未找到SVG文件，请确保以下文件存在于当前目录:")
        for f in svg_files:
            print(f"  - {f}")
        return

    print("=" * 70)
    print("通用SVG到PPTX转换器 v2.0")
    print("=" * 70)
    print()

    # 显示已注册的处理器
    registry = HandlerRegistry()
    print("已注册的处理器:")
    for handler_info in registry.list_handlers():
        print(f"  - {handler_info}")
    print()

    # 转换每个文件
    for svg_file in found_files:
        output_file = svg_file.replace('.svg', '-v2.pptx')

        print(f"\n转换: {svg_file}")
        print("-" * 70)

        try:
            converter = SVGToPPTXConverter(svg_file)
            converter.convert(output_file)
            print(f"输出: {output_file}")
        except Exception as e:
            print(f"错误: {e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 70)
    print("所有转换完成！")
    print("=" * 70)


if __name__ == "__main__":
    main()
