"""
SVG解析器
解析SVG文件为SVGElement树
"""

import re
import xml.etree.ElementTree as ET
from typing import List, Optional, Dict

from .models import SVGElement, Style, BoundingBox
from .color_utils import ColorParser


class SVGParser:
    """SVG解析器"""

    def __init__(self, svg_path: str):
        self.svg_path = svg_path
        self.viewbox: tuple = (0, 0, 100, 100)
        self.width: float = 100
        self.height: float = 100
        self._css_classes: Dict[str, Dict[str, str]] = {}  # CSS类样式缓存

    def parse(self) -> List[SVGElement]:
        """解析SVG文件，返回顶层元素列表"""
        tree = ET.parse(self.svg_path)
        root = tree.getroot()

        # 解析viewBox
        self._parse_viewbox(root)

        # 先解析CSS类样式（从defs/style）
        self._parse_css_styles(root)

        # 解析根元素的子元素
        elements = []
        for child in root:
            elem = self._parse_element(child)
            if elem:
                elements.append(elem)

        return elements

    def _parse_css_styles(self, root: ET.Element):
        """解析CSS类样式（从<style>标签）"""
        for defs in root.findall('.//defs'):
            for style_elem in defs.findall('.//style'):
                style_text = style_elem.text or ''
                self._css_classes = self._parse_css_text(style_text)

    def _parse_css_text(self, css_text: str) -> Dict[str, Dict[str, str]]:
        """解析CSS文本，提取类规则"""
        css_classes = {}

        # 正则匹配 .classname { prop: value; ... }
        pattern = r'\.([a-zA-Z_-][a-zA-Z0-9_-]*)\s*\{([^}]+)\}'
        matches = re.findall(pattern, css_text, re.DOTALL)

        for class_name, content in matches:
            props = {}
            for line in content.split(';'):
                if ':' in line:
                    key, value = line.split(':', 1)
                    props[key.strip()] = value.strip()
            if props:
                css_classes[class_name] = props

        return css_classes

    def _get_class_style(self, class_attr: str) -> Dict[str, str]:
        """获取CSS类的样式属性"""
        result = {}
        if not class_attr:
            return result

        # 支持多个类（用空格分隔）
        for class_name in class_attr.split():
            class_name = class_name.strip()
            if class_name and class_name in self._css_classes:
                result.update(self._css_classes[class_name])

        return result

    def _parse_viewbox(self, root: ET.Element):
        """解析viewBox和尺寸信息"""
        # 尝试获取viewBox
        viewbox_attr = root.get('viewBox')
        if viewbox_attr:
            parts = viewbox_attr.replace(',', ' ').split()
            if len(parts) >= 4:
                try:
                    self.viewbox = (float(parts[0]), float(parts[1]),
                                   float(parts[2]), float(parts[3]))
                    self.width = float(parts[2])
                    self.height = float(parts[3])
                except ValueError:
                    pass

        # 如果没有viewBox，尝试width和height
        if self.width == 100:
            width_attr = root.get('width', '')
            self.width = self._parse_dimension(width_attr)

        if self.height == 100:
            height_attr = root.get('height', '')
            self.height = self._parse_dimension(height_attr)

    def _parse_dimension(self, value: str) -> float:
        """解析尺寸值（支持px单位）"""
        if not value:
            return 100
        try:
            # 移除单位（px, pt等）
            import re
            match = re.match(r'([\d.]+)', value)
            return float(match.group(1)) if match else 100
        except ValueError:
            return 100

    def _parse_element(self, xml_element: ET.Element,
                      parent: Optional[SVGElement] = None) -> Optional[SVGElement]:
        """递归解析元素"""
        tag = xml_element.tag.split('}')[-1] if '}' in xml_element.tag else xml_element.tag

        # 跳过defs等非渲染元素
        if tag in ('defs', 'marker', 'linearGradient', 'radialGradient', 'pattern'):
            return None

        # 跳过style标签（已单独解析）
        if tag == 'style':
            return None

        # 合并CSS类样式到属性
        class_attr = xml_element.get('class', '')
        if class_attr and class_attr in self._css_classes:
            css_props = self._get_class_style(class_attr)
            # 将CSS样式合并到内联style属性
            existing_style = xml_element.get('style', '')
            css_style_str = '; '.join(f"{k}: {v}" for k, v in css_props.items())
            if existing_style:
                xml_element.set('style', existing_style + '; ' + css_style_str)
            else:
                xml_element.set('style', css_style_str)

        # 解析样式
        style = self._parse_style(xml_element.attrib)

        # 创建SVGElement
        element = SVGElement(
            tag=tag,
            attrib=dict(xml_element.attrib),
            style=style,
            parent=parent
        )

        # 递归解析子元素
        for child in xml_element:
            child_elem = self._parse_element(child, element)
            if child_elem:
                element.children.append(child_elem)

        # 提取文本内容（对于text元素）
        if tag == 'text':
            element.text_content = ''.join(xml_element.itertext()).strip()

        return element

    def _parse_style(self, attrib: Dict[str, str]) -> Style:
        """解析样式属性"""
        style = Style()

        # 直接属性
        if 'fill' in attrib:
            style.fill = attrib['fill']
        if 'stroke' in attrib:
            style.stroke = attrib['stroke']
        if 'stroke-width' in attrib:
            style.stroke_width = ColorParser.parse_length(attrib['stroke-width'])
        if 'stroke-dasharray' in attrib:
            style.stroke_dasharray = attrib['stroke-dasharray']
        if 'opacity' in attrib:
            try:
                style.opacity = float(attrib['opacity'])
            except ValueError:
                pass
        if 'font-size' in attrib:
            style.font_size = ColorParser.parse_length(attrib['font-size'])
        if 'font-family' in attrib:
            style.font_family = attrib['font-family']
        if 'font-weight' in attrib:
            style.font_weight = attrib['font-weight']
        if 'text-anchor' in attrib:
            style.text_anchor = attrib['text-anchor']
        if 'dominant-baseline' in attrib:
            style.dominant_baseline = attrib['dominant-baseline']

        # style属性（内联样式）
        style_attr = attrib.get('style', '')
        if style_attr:
            style = self._merge_inline_style(style, style_attr)

        return style

    def _merge_inline_style(self, style: Style, style_str: str) -> Style:
        """合并内联样式"""
        new_style = Style()

        for item in style_str.split(';'):
            if ':' not in item:
                continue
            key, value = item.split(':', 1)
            key = key.strip()
            value = value.strip()

            if key == 'fill':
                new_style.fill = value
            elif key == 'stroke':
                new_style.stroke = value
            elif key == 'stroke-width':
                new_style.stroke_width = ColorParser.parse_length(value)
            elif key == 'stroke-dasharray':
                new_style.stroke_dasharray = value
            elif key == 'opacity':
                try:
                    new_style.opacity = float(value)
                except ValueError:
                    pass
            elif key == 'font-size':
                new_style.font_size = ColorParser.parse_length(value)
            elif key == 'font-family':
                new_style.font_family = value
            elif key == 'font-weight':
                new_style.font_weight = value
            elif key == 'text-anchor':
                new_style.text_anchor = value
            elif key == 'dominant-baseline':
                new_style.dominant_baseline = value

        return style.merge(new_style)
