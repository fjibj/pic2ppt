"""
文本处理器
"""

from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from . import ElementHandler, RenderContext
from ..models import SVGElement
from ..color_utils import ColorParser


class TextHandler(ElementHandler):
    """文本处理器"""

    def can_handle(self, element: SVGElement) -> bool:
        return element.tag == 'text'

    def handle(self, element: SVGElement, context: RenderContext):
        x = element.get_float_attr('x')
        y = element.get_float_attr('y')
        text_anchor = element.get_str_attr('text-anchor', 'start')
        writing_mode = element.get_str_attr('writing-mode', 'lr')  # lr = left-to-right, tb = top-to-bottom

        # 获取文本内容
        text = element.text_content
        if not text:
            return None

        # 获取样式（父样式合并子样式，子样式优先级更高）
        style = context.parent_style.merge(element.style)

        # 解析字体大小（SVG默认16px）
        font_size = style.font_size or 16
        # 转换为点（pt），考虑缩放
        # SVG px 到 pt: 1px = 0.75pt (72pt / 96dpi)
        font_size_pt = max(font_size * context.scale * 0.75, 8)

        # 判断是否为竖向文字
        is_vertical = writing_mode in ('tb', 'tb-rl', 'vertical-rl', 'vertical-lr')

        # 估算文本尺寸
        # 中文字符宽度约为字体大小的1.1倍
        char_width_inch = (font_size_pt * 1.1) / 72
        char_height_inch = (font_size_pt * 1.5) / 72

        if is_vertical:
            # 竖向文字：宽度是单个字符宽度，高度是所有字符累计
            text_width = char_width_inch
            text_height = len(text) * char_height_inch
        else:
            # 横向文字
            text_width = len(text) * char_width_inch
            text_height = char_height_inch

        # 计算位置（SVG坐标转英寸）
        base_x = context.offset_x + context.svg_to_inches(x)
        base_y = context.offset_y + context.svg_to_inches(y)

        # 根据text-anchor调整水平位置
        # SVG: start/middle/end
        # PPT: LEFT/CENTER/RIGHT
        if text_anchor == 'middle':
            if is_vertical:
                left = Inches(base_x - text_width / 2)
            else:
                left = Inches(base_x - text_width / 2)
            align = PP_ALIGN.CENTER
        elif text_anchor == 'end':
            if is_vertical:
                left = Inches(base_x - text_width)
            else:
                left = Inches(base_x - text_width)
            align = PP_ALIGN.RIGHT
        else:  # start
            left = Inches(base_x)
            align = PP_ALIGN.LEFT

        # 垂直位置调整
        # SVG的y坐标是文本基线位置（baseline）
        # PPT文本框的top是框的顶部，需要向上偏移基线到顶部的距离
        # 对于中文字体，基线通常在字体大小的 0.7-0.8 倍处
        if is_vertical:
            # 竖向文字：y是文本框中心或起始位置
            baseline_offset = text_height / 2
        else:
            # 横向文字：基线到顶部的距离
            # SVG: y = baseline position
            # PPT: top = textbox top
            # 需要偏移: baseline_to_top = font_size * baseline_ratio
            # 中文字体的基线通常在em-box的下部，约0.75-0.85倍字体大小
            baseline_offset = (font_size_pt * 0.85) / 72  # 转换英寸

        # 确保偏移量不超过位置本身（避免负值）
        top = Inches(max(0, base_y - baseline_offset))

        # 确保文本框有足够的大小
        min_width = max(text_width, 0.1)  # 最小0.1英寸
        min_height = max(text_height, 0.05)  # 最小0.05英寸

        # 创建文本框
        txBox = context.slide.shapes.add_textbox(
            Inches(max(left.inches, 0)),
            Inches(max(top.inches, 0)),
            Inches(min_width),
            Inches(min_height)
        )

        tf = txBox.text_frame
        tf.word_wrap = False  # 不自动换行，保持原有布局

        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(font_size_pt)
        p.font.name = style.font_family or 'Microsoft YaHei'
        p.alignment = align

        # 设置竖向文字（通过XML设置vert属性）
        if is_vertical:
            # 在PPT中，竖向文字应该使用bodyPr的vert属性
            # vert="eaVert" 表示东亚竖排文字
            from lxml import etree
            bodyPr = txBox._element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
            if bodyPr is not None:
                bodyPr.set('vert', 'eaVert')  # 东亚竖排
                # 移除旋转，竖排不需要旋转
                if 'rot' in bodyPr.attrib:
                    del bodyPr.attrib['rot']

        # 设置颜色（使用fill属性）
        fill_color = style.fill
        if fill_color and fill_color.lower() != 'none':
            color = ColorParser.parse(fill_color)
            if color:
                p.font.color.rgb = color
        else:
            # 默认黑色
            p.font.color.rgb = RGBColor(0, 0, 0)

        # 设置字体粗细
        if style.font_weight:
            p.font.bold = style.font_weight in ('bold', '700', 'bolder')

        return txBox
