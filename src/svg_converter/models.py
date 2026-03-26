"""
数据模型定义
包含SVGElement、Style、BoundingBox、Point等核心类
"""

from dataclasses import dataclass, field
from typing import Dict, List, Tuple, Optional, Any
import re


@dataclass
class Point:
    """2D点"""
    x: float
    y: float

    def __add__(self, other: 'Point') -> 'Point':
        return Point(self.x + other.x, self.y + other.y)

    def __sub__(self, other: 'Point') -> 'Point':
        return Point(self.x - other.x, self.y - other.y)

    def distance_to(self, other: 'Point') -> float:
        """计算到另一点的距离"""
        return ((self.x - other.x) ** 2 + (self.y - other.y) ** 2) ** 0.5


@dataclass
class BoundingBox:
    """包围盒"""
    x: float
    y: float
    width: float
    height: float

    @property
    def center(self) -> Point:
        return Point(self.x + self.width / 2, self.y + self.height / 2)

    @property
    def aspect_ratio(self) -> float:
        return self.width / self.height if self.height > 0 else 1.0

    @property
    def left(self) -> float:
        return self.x

    @property
    def top(self) -> float:
        return self.y

    @property
    def right(self) -> float:
        return self.x + self.width

    @property
    def bottom(self) -> float:
        return self.y + self.height


@dataclass
class Style:
    """样式属性（支持继承）"""
    fill: Optional[str] = None
    stroke: Optional[str] = None
    stroke_width: float = 1.0
    stroke_dasharray: Optional[str] = None
    opacity: float = 1.0
    font_size: Optional[float] = None
    font_family: Optional[str] = None
    font_weight: Optional[str] = None
    text_anchor: str = "start"
    dominant_baseline: str = "auto"

    def merge(self, other: 'Style') -> 'Style':
        """合并样式（other优先级更高）"""
        return Style(
            fill=other.fill if other.fill is not None else self.fill,
            stroke=other.stroke if other.stroke is not None else self.stroke,
            stroke_width=other.stroke_width if other.stroke_width != 1.0 else self.stroke_width,
            stroke_dasharray=other.stroke_dasharray if other.stroke_dasharray is not None else self.stroke_dasharray,
            opacity=other.opacity * self.opacity,
            font_size=other.font_size if other.font_size is not None else self.font_size,
            font_family=other.font_family if other.font_family is not None else self.font_family,
            font_weight=other.font_weight if other.font_weight is not None else self.font_weight,
            text_anchor=other.text_anchor if other.text_anchor != "start" else self.text_anchor,
            dominant_baseline=other.dominant_baseline if other.dominant_baseline != "auto" else self.dominant_baseline,
        )


@dataclass
class SVGElement:
    """SVG元素基类"""
    tag: str
    attrib: Dict[str, str]
    style: Style
    children: List['SVGElement'] = field(default_factory=list)
    parent: Optional['SVGElement'] = None
    text_content: str = ""

    def get_bounding_box(self) -> Optional[BoundingBox]:
        """计算元素包围盒（子类可重写）"""
        return None

    def get_float_attr(self, name: str, default: float = 0.0) -> float:
        """获取浮点数属性"""
        try:
            return float(self.attrib.get(name, default))
        except (ValueError, TypeError):
            return default

    def get_str_attr(self, name: str, default: str = "") -> str:
        """获取字符串属性"""
        return self.attrib.get(name, default)


class ShapeType:
    """形状类型常量"""
    # 基础形状
    RECTANGLE = "RECTANGLE"
    ROUNDED_RECTANGLE = "ROUNDED_RECTANGLE"
    OVAL = "OVAL"

    # 多边形
    DIAMOND = "DIAMOND"
    ISOSCELES_TRIANGLE = "ISOSCELES_TRIANGLE"
    RIGHT_TRIANGLE = "RIGHT_TRIANGLE"
    HEXAGON = "HEXAGON"
    OCTAGON = "OCTAGON"
    PENTAGON = "PENTAGON"
    TRAPEZOID = "TRAPEZOID"
    PARALLELOGRAM = "PARALLELOGRAM"

    # 线条
    STRAIGHT_LINE = "STRAIGHT_LINE"
    ELBOW_LINE = "ELBOW_LINE"
    CURVED_LINE = "CURVED_LINE"

    # 其他
    TEXTBOX = "TEXTBOX"
    FREEFORM = "FREEFORM"
    GROUP = "GROUP"
    UNKNOWN = "UNKNOWN"
