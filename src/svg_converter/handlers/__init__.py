"""
处理器注册表
管理所有SVG元素处理器
"""

from abc import ABC, abstractmethod
from typing import List, Tuple, Optional, Any
from ..models import SVGElement, Style


class ElementHandler(ABC):
    """元素处理器抽象基类"""

    @abstractmethod
    def can_handle(self, element: SVGElement) -> bool:
        """判断是否能处理该元素"""
        pass

    @abstractmethod
    def handle(self, element: SVGElement, context: 'RenderContext') -> Any:
        """处理元素，返回PPT形状对象或对象列表"""
        pass


class HandlerRegistry:
    """处理器注册中心 - 单例模式"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._handlers: List[Tuple[int, ElementHandler]] = []
        return cls._instance

    def register(self, handler: ElementHandler, priority: int = 0):
        """注册处理器，支持优先级（高优先级优先）"""
        self._handlers.append((priority, handler))
        self._handlers.sort(key=lambda x: -x[0])  # 降序排列

    def unregister(self, handler_type: type):
        """注销特定类型的处理器"""
        self._handlers = [
            (p, h) for p, h in self._handlers
            if not isinstance(h, handler_type)
        ]

    def get_handler(self, element: SVGElement) -> Optional[ElementHandler]:
        """获取能处理该元素的处理器"""
        for _, handler in self._handlers:
            if handler.can_handle(element):
                return handler
        return None

    def clear(self):
        """清空所有处理器"""
        self._handlers.clear()

    def list_handlers(self) -> List[str]:
        """列出所有已注册的处理器"""
        return [f"{h.__class__.__name__} (priority={p})" for p, h in self._handlers]


class RenderContext:
    """渲染上下文"""

    def __init__(self, slide: Any, scale: float = 1.0,
                 offset_x: float = 0.0, offset_y: float = 0.0,
                 parent_style: Style = None, depth: int = 0):
        self.slide = slide
        self.scale = scale
        self.offset_x = offset_x
        self.offset_y = offset_y
        self.parent_style = parent_style or Style()
        self.depth = depth

    def with_style(self, style: Style) -> 'RenderContext':
        """创建带新样式的上下文（继承当前样式）"""
        new_style = self.parent_style.merge(style)
        return RenderContext(
            slide=self.slide,
            scale=self.scale,
            offset_x=self.offset_x,
            offset_y=self.offset_y,
            parent_style=new_style,
            depth=self.depth + 1
        )

    def svg_to_inches(self, value: float) -> float:
        """SVG像素转英寸（96 DPI）"""
        return value / 96.0 * self.scale
