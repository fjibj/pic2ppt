"""
颜色解析工具
支持多种颜色格式的解析
"""

import re
from typing import Optional
from pptx.dml.color import RGBColor


class ColorParser:
    """颜色解析器"""

    # 命名颜色映射表
    NAMED_COLORS = {
        'black': (0, 0, 0),
        'white': (255, 255, 255),
        'red': (255, 0, 0),
        'green': (0, 128, 0),
        'lime': (0, 255, 0),
        'blue': (0, 0, 255),
        'yellow': (255, 255, 0),
        'cyan': (0, 255, 255),
        'magenta': (255, 0, 255),
        'silver': (192, 192, 192),
        'gray': (128, 128, 128),
        'grey': (128, 128, 128),
        'maroon': (128, 0, 0),
        'olive': (128, 128, 0),
        'purple': (128, 0, 128),
        'teal': (0, 128, 128),
        'navy': (0, 0, 128),
        'orange': (255, 165, 0),
        'pink': (255, 192, 203),
        'brown': (165, 42, 42),
    }

    @classmethod
    def parse(cls, color_str: Optional[str]) -> Optional[RGBColor]:
        """
        解析颜色字符串为RGBColor

        支持格式:
        - rgb(r, g, b)
        - rgba(r, g, b, a) - 忽略透明度
        - #RRGGBB
        - #RGB
        - 命名颜色 (black, white, red等)
        """
        if not color_str or color_str.lower() in ('none', 'transparent'):
            return None

        color_str = color_str.strip()

        # 1. 解析 rgba(r, g, b, a)
        match = re.match(r'rgba\((\d+),\s*(\d+),\s*(\d+),\s*[\d.]+\)', color_str, re.IGNORECASE)
        if match:
            return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))

        # 2. 解析 rgb(r, g, b)
        match = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color_str, re.IGNORECASE)
        if match:
            return RGBColor(int(match.group(1)), int(match.group(2)), int(match.group(3)))

        # 3. 解析十六进制 #RRGGBB 或 #RGB
        if color_str.startswith('#'):
            hex_color = color_str[1:]
            if len(hex_color) == 6:
                try:
                    r = int(hex_color[0:2], 16)
                    g = int(hex_color[2:4], 16)
                    b = int(hex_color[4:6], 16)
                    return RGBColor(r, g, b)
                except ValueError:
                    pass
            elif len(hex_color) == 3:
                try:
                    r = int(hex_color[0] * 2, 16)
                    g = int(hex_color[1] * 2, 16)
                    b = int(hex_color[2] * 2, 16)
                    return RGBColor(r, g, b)
                except ValueError:
                    pass

        # 4. 解析命名颜色
        color_lower = color_str.lower()
        if color_lower in cls.NAMED_COLORS:
            r, g, b = cls.NAMED_COLORS[color_lower]
            return RGBColor(r, g, b)

        return None

    @classmethod
    def parse_length(cls, value: Optional[str]) -> float:
        """解析长度值（支持px, pt, em等单位）"""
        if not value:
            return 0.0
        value = str(value).strip()
        match = re.match(r'([\d.]+)(?:px|pt|em|%)?', value)
        return float(match.group(1)) if match else 0.0
