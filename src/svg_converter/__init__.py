"""
通用SVG到PPTX转换器
支持完整的SVG元素映射到PowerPoint可编辑形状
"""

from .converter import SVGToPPTXConverter
from .models import SVGElement, Style, BoundingBox, Point
from .handlers import HandlerRegistry

__version__ = "2.0.0"
__all__ = [
    "SVGToPPTXConverter",
    "SVGElement",
    "Style",
    "BoundingBox",
    "Point",
    "HandlerRegistry",
]
