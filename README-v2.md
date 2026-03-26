# 通用SVG到PPTX转换器 v2.0

基于插件架构的通用SVG转换工具，自动识别SVG元素并映射到PowerPoint可编辑形状。

## 架构特点

### 1. 插件式处理器架构

```
SVG文件 → SVGParser → ElementHandler → PPT形状
                          ↓
                   HandlerRegistry
                   (处理器注册表)
                          ↓
              RectHandler, CircleHandler,
              PolygonHandler, TextHandler...
```

- **每个SVG元素类型对应独立处理器**
- **自动匹配最合适的处理器**
- **易于扩展：添加新元素只需注册新处理器**

### 2. 几何分析引擎

自动从 `<polygon points="...">` 识别形状类型：

| 点数 | 自动识别 | PPT形状 |
|------|---------|---------|
| 3 | 等腰/直角三角形 | ISOSCELES_TRIANGLE / RIGHT_TRIANGLE |
| 4 | 菱形/矩形/梯形 | DIAMOND / RECTANGLE / TRAPEZOID |
| 6 | 正六边形 | HEXAGON |
| 8 | 正八边形 | OCTAGON |

### 3. 完整的SVG元素支持

| SVG元素 | 检测方式 | PPT映射 | 状态 |
|---------|----------|---------|------|
| `<rect>` | 属性读取 | RECTANGLE / ROUNDED_RECTANGLE | ✅ |
| `<circle>` | 属性读取 | OVAL (正圆) | ✅ |
| `<ellipse>` | 属性读取 | OVAL (椭圆) | ✅ |
| `<line>` | 属性读取 | STRAIGHT_CONNECTOR | ✅ |
| `<polygon>` | 几何分析 | DIAMOND/HEXAGON等 | ✅ |
| `<polyline>` | 几何分析 | 折线/连接器 | ✅ |
| `<path>` | 路径分析 | 对应形状/FREEFORM | ✅ |
| `<text>` | 属性读取 | 文本框 | ✅ |
| `<g>` | 递归处理 | 样式继承 | ✅ |

## 使用方法

### 命令行使用

```bash
# 基本用法
python src/convert.py diagram.svg

# 指定输出文件
python src/convert.py diagram.svg -o output.pptx

# 16:9 比例
python src/convert.py diagram.svg -W 16 -H 9

# 小边距
python src/convert.py diagram.svg -m 0.2
```

### Python API使用

```python
from svg_converter import SVGToPPTXConverter

# 基本转换
converter = SVGToPPTXConverter("diagram.svg")
converter.convert("output.pptx")

# 自定义幻灯片尺寸
converter.convert(
    "output.pptx",
    slide_width=16,      # 16英寸
    slide_height=9,      # 9英寸
    margin=0.5           # 0.5英寸边距
)
```

### 便捷的转换函数

```python
from svg_converter import convert_svg_to_pptx

# 一行代码转换
convert_svg_to_pptx("diagram.svg", "output.pptx")
```

## 项目结构

```
src/
├── svg_converter/           # 核心包
│   ├── __init__.py
│   ├── models.py            # 数据模型（SVGElement, Style, BoundingBox）
│   ├── parser.py            # SVG解析器
│   ├── converter.py         # 主转换器
│   ├── geometry.py          # 几何分析引擎
│   ├── color_utils.py       # 颜色解析工具
│   └── handlers/            # 处理器包
│       ├── __init__.py      # 处理器基类和注册表
│       ├── basic_shapes.py  # 基础形状（rect, circle, line）
│       ├── polygons.py      # 多边形（polygon, path）
│       ├── text.py          # 文本处理
│       └── group.py         # 分组处理
├── convert.py               # 命令行工具
└── demo.py                  # 示例脚本
```

## 扩展开发

### 添加自定义处理器

```python
from svg_converter.handlers import ElementHandler, RenderContext
from svg_converter import HandlerRegistry

class StarHandler(ElementHandler):
    """五角星处理器示例"""

    def can_handle(self, element: SVGElement) -> bool:
        # 判断是否为五角星
        return element.tag == 'polygon' and self._is_star(element)

    def handle(self, element: SVGElement, context: RenderContext):
        # 处理逻辑
        # ... 创建五角星形状 ...
        return shape

# 注册到系统
registry = HandlerRegistry()
registry.register(StarHandler(), priority=100)
```

### 扩展几何分析器

```python
from svg_converter.geometry import GeometryAnalyzer

class ExtendedAnalyzer(GeometryAnalyzer):
    def _analyze_pentagon(self, points, bbox):
        # 自定义五边形检测
        if self._is_regular_polygon(points, 5):
            return ShapeType.PENTAGON, bbox
        return ShapeType.FREEFORM, bbox
```

## 支持的样式

### 颜色
- `rgb(r, g, b)` ✅
- `rgba(r, g, b, a)` ✅ (忽略透明度)
- `#RRGGBB` / `#RGB` ✅
- 命名颜色 (black, white, red等) ✅

### 边框
- `stroke` - 边框颜色 ✅
- `stroke-width` - 边框宽度 ✅
- `stroke-dasharray` - 虚线样式 ✅

### 文本
- `font-size` - 字体大小 ✅
- `font-family` - 字体 ✅
- `font-weight` - 字重 ✅
- `text-anchor` - 对齐方式 (start/middle/end) ✅

### 其他
- `fill` - 填充颜色 ✅
- `opacity` - 透明度 (相乘计算) ✅
- `style` 属性 (内联CSS) ✅

## 对比 v1.0 vs v2.0

| 特性 | v1.0 | v2.0 |
|------|------|------|
| 架构 | 硬编码 | 插件式处理器 |
| 新增元素支持 | 需修改多处 | 只需添加处理器 |
| 多边形识别 | 手动判断 | 自动几何分析 |
| 样式继承 | 不支持 | 支持 `<g>` 分组 |
| circle/ellipse | ❌ | ✅ |
| path | ❌ | ✅ (基础) |
| 扩展性 | 低 | 高 |

## 测试结果

### ai_component_management_platform.svg
- 顶层元素：29个
- 成功转换：29个 (100%)
- 包含：矩形、文本、线条、箭头、虚线框

### reinforcement-fine-tuning-flowchart.svg
- 顶层元素：60个
- 成功转换：60个 (100%)
- 包含：矩形、**菱形**、文本、线条、箭头

## 注意事项

1. **WPS支持**：WPS不支持"转换为形状"功能，请使用Microsoft PowerPoint
2. **复杂路径**：`<path>` 元素目前使用简化处理，复杂路径会映射为矩形
3. **渐变填充**：暂不支持，会使用近似纯色填充
4. **变换矩阵**：`transform` 属性暂不支持

## 依赖

```
python-pptx>=0.6.21
lxml>=4.9.0
```

## License

MIT License
