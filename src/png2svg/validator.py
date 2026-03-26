#!/usr/bin/env python3
"""
SVG 验证器
验证和修复 AI 生成的 SVG
"""

import re
from typing import List, Tuple


class SVGValidator:
    """验证 SVG 格式是否正确"""

    @staticmethod
    def validate(svg_content: str) -> Tuple[bool, List[str]]:
        """
        验证 SVG 内容

        Returns:
            (是否有效, 错误列表)
        """
        errors = []

        # 检查 XML 声明
        if not svg_content.strip().startswith('<?xml'):
            errors.append("缺少 XML 声明")

        # 检查 SVG 标签
        if '<svg' not in svg_content:
            errors.append("缺少 SVG 根元素")

        # 检查命名空间
        if 'xmlns="http://www.w3.org/2000/svg"' not in svg_content:
            errors.append("缺少 SVG 命名空间")

        # 检查标签闭合
        tags = re.findall(r'<([a-zA-Z][a-zA-Z0-9]*)[^>]*?[^/]>', svg_content)
        for tag in set(tags):
            if tag not in ['br', 'hr', 'img', 'input', 'meta', 'link']:
                open_count = svg_content.count(f'<{tag}')
                close_count = svg_content.count(f'</{tag}>')
                if open_count != close_count:
                    errors.append(f"<{tag}> 标签未正确闭合 (开:{open_count}, 闭:{close_count})")

        # 检查常见错误
        if '&nbsp;' in svg_content:
            errors.append("包含未转义的 &nbsp;，应使用 &#160;")

        return len(errors) == 0, errors

    @staticmethod
    def fix_common_issues(svg_content: str) -> str:
        """修复常见的 SVG 问题"""
        # 修复未转义的特殊字符
        svg_content = svg_content.replace('&nbsp;', '&#160;')
        svg_content = svg_content.replace('&', '&amp;')
        svg_content = svg_content.replace('&amp;amp;', '&amp;')
        svg_content = svg_content.replace('&amp;lt;', '&lt;')
        svg_content = svg_content.replace('&amp;gt;', '&gt;')
        svg_content = svg_content.replace('&amp;quot;', '&quot;')

        # 移除空文本元素
        svg_content = re.sub(r'<text[^>]*>\s*</text>', '', svg_content)

        # 确保 viewBox 存在
        if 'viewBox' not in svg_content and 'width' in svg_content and 'height' in svg_content:
            # 尝试从 width/height 推断 viewBox
            width_match = re.search(r'width=["\']([^"\']+)["\']', svg_content)
            height_match = re.search(r'height=["\']([^"\']+)["\']', svg_content)
            if width_match and height_match:
                width = width_match.group(1).replace('px', '')
                height = height_match.group(1).replace('px', '')
                try:
                    svg_content = svg_content.replace(
                        '<svg',
                        f'<svg viewBox="0 0 {width} {height}"',
                        1
                    )
                except ValueError:
                    pass

        return svg_content

    @staticmethod
    def get_statistics(svg_content: str) -> dict:
        """获取 SVG 统计信息"""
        stats = {
            'total_length': len(svg_content),
            'element_counts': {}
        }

        # 统计各种元素数量
        elements = ['rect', 'circle', 'ellipse', 'line', 'path', 'polygon',
                    'text', 'g', 'defs', 'style']
        for elem in elements:
            count = len(re.findall(rf'<{elem}\b', svg_content, re.IGNORECASE))
            if count > 0:
                stats['element_counts'][elem] = count

        return stats


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("用法: python validator.py <svg文件路径>")
        sys.exit(1)

    from pathlib import Path
    svg_file = Path(sys.argv[1])

    if not svg_file.exists():
        print(f"错误: 文件不存在 {svg_file}")
        sys.exit(1)

    content = svg_file.read_text(encoding='utf-8')

    validator = SVGValidator()

    # 验证
    is_valid, errors = validator.validate(content)
    print(f"验证结果: {'通过' if is_valid else '失败'}")

    if errors:
        print("错误列表:")
        for err in errors:
            print(f"  - {err}")

    # 统计
    stats = validator.get_statistics(content)
    print(f"\n文件大小: {stats['total_length']} 字节")
    print("元素统计:")
    for elem, count in sorted(stats['element_counts'].items()):
        print(f"  <{elem}>: {count}")
