"""
分组处理器
处理 <g> 元素和样式继承
"""

from typing import List, Any

from . import ElementHandler, RenderContext, HandlerRegistry
from ..models import SVGElement


class GroupHandler(ElementHandler):
    """分组处理器 - 处理<g>元素"""

    def can_handle(self, element: SVGElement) -> bool:
        return element.tag == 'g'

    def handle(self, element: SVGElement, context: RenderContext) -> List[Any]:
        """处理分组，递归处理子元素"""
        # 合并样式（分组样式继承）
        new_context = context.with_style(element.style)

        shapes = []
        registry = HandlerRegistry()

        for child in element.children:
            handler = registry.get_handler(child)
            if handler:
                try:
                    result = handler.handle(child, new_context)
                    if result:
                        if isinstance(result, list):
                            shapes.extend(result)
                        else:
                            shapes.append(result)
                except Exception as e:
                    print(f"处理子元素 {child.tag} 时出错: {e}")

        return shapes
