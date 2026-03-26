"""
几何分析引擎
用于从points/path自动识别形状类型
"""

import math
import re
from typing import List, Tuple, Optional
from .models import Point, BoundingBox, ShapeType


class GeometryAnalyzer:
    """几何分析引擎 - 从points/path自动识别形状类型"""

    # 角度容差（度）
    ANGLE_TOLERANCE = 15
    # 边长比例容差
    RATIO_TOLERANCE = 0.15
    # 正多边形检测容差
    REGULAR_TOLERANCE = 0.15

    def analyze_points(self, points: List[Tuple[float, float]]) -> Tuple[str, BoundingBox]:
        """
        分析点集，返回形状类型和包围盒

        Args:
            points: 点坐标列表 [(x1,y1), (x2,y2), ...]

        Returns:
            (shape_type, bounding_box)
        """
        if not points:
            return ShapeType.UNKNOWN, BoundingBox(0, 0, 0, 0)

        # 计算包围盒
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        bbox = BoundingBox(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))

        # 根据点数判断形状类型
        point_count = len(points)

        if point_count == 2:
            return ShapeType.STRAIGHT_LINE, bbox
        elif point_count == 3:
            return self._analyze_triangle(points, bbox)
        elif point_count == 4:
            return self._analyze_quadrilateral(points, bbox)
        elif point_count == 5:
            return self._analyze_polygon(points, bbox, ShapeType.PENTAGON)
        elif point_count == 6:
            return self._analyze_polygon(points, bbox, ShapeType.HEXAGON)
        elif point_count == 8:
            return self._analyze_polygon(points, bbox, ShapeType.OCTAGON)

        # 复杂多边形
        return ShapeType.FREEFORM, bbox

    def _calculate_sides(self, points: List[Tuple[float, float]]) -> List[float]:
        """计算多边形各边长度"""
        n = len(points)
        return [self._distance(points[i], points[(i + 1) % n]) for i in range(n)]

    def _distance(self, p1: Tuple[float, float], p2: Tuple[float, float]) -> float:
        """计算两点距离"""
        return math.sqrt((p1[0] - p2[0]) ** 2 + (p1[1] - p2[1]) ** 2)

    def _calculate_angles(self, points: List[Tuple[float, float]]) -> List[float]:
        """计算多边形各内角（度）"""
        n = len(points)
        angles = []
        for i in range(n):
            p1 = points[(i - 1) % n]
            p2 = points[i]
            p3 = points[(i + 1) % n]
            angle = self._calculate_angle(p1, p2, p3)
            angles.append(angle)
        return angles

    def _calculate_angle(self, p1: Tuple[float, float], p2: Tuple[float, float],
                        p3: Tuple[float, float]) -> float:
        """计算三点形成的角度（p2为顶点）"""
        v1 = (p1[0] - p2[0], p1[1] - p2[1])
        v2 = (p3[0] - p2[0], p3[1] - p2[1])

        dot = v1[0] * v2[0] + v1[1] * v2[1]
        mag1 = math.sqrt(v1[0] ** 2 + v1[1] ** 2)
        mag2 = math.sqrt(v2[0] ** 2 + v2[1] ** 2)

        if mag1 == 0 or mag2 == 0:
            return 0

        cos_angle = dot / (mag1 * mag2)
        cos_angle = max(-1, min(1, cos_angle))  # 处理数值误差
        return math.degrees(math.acos(cos_angle))

    def _is_regular_polygon(self, points: List[Tuple[float, float]], expected_sides: int) -> bool:
        """检查是否为正多边形"""
        if len(points) != expected_sides:
            return False

        sides = self._calculate_sides(points)
        avg_side = sum(sides) / len(sides)

        # 检查边长是否接近相等
        return all(abs(s - avg_side) < avg_side * self.REGULAR_TOLERANCE for s in sides)

    def _analyze_triangle(self, points: List[Tuple[float, float]], bbox: BoundingBox) -> Tuple[str, BoundingBox]:
        """分析三角形类型"""
        sides = self._calculate_sides(points)
        sides.sort()
        a, b, c = sides[0], sides[1], sides[2]

        # 检查是否为直角三角形（勾股定理）
        is_right = abs(a ** 2 + b ** 2 - c ** 2) < (c ** 2 * 0.1)

        # 检查是否为等腰
        is_isosceles = abs(sides[0] - sides[1]) < c * 0.1 or abs(sides[1] - sides[2]) < c * 0.1

        if is_right:
            return ShapeType.RIGHT_TRIANGLE, bbox
        elif is_isosceles:
            return ShapeType.ISOSCELES_TRIANGLE, bbox

        return ShapeType.ISOSCELES_TRIANGLE, bbox  # 默认为等腰三角形

    def _analyze_quadrilateral(self, points: List[Tuple[float, float]], bbox: BoundingBox) -> Tuple[str, BoundingBox]:
        """分析四边形类型"""
        sides = self._calculate_sides(points)
        angles = self._calculate_angles(points)

        # 计算对角线
        diag1 = self._distance(points[0], points[2])
        diag2 = self._distance(points[1], points[3])

        # 检查是否为菱形（四边相等）
        avg_side = sum(sides) / 4
        is_rhombus = all(abs(s - avg_side) < avg_side * 0.1 for s in sides)

        # 检查是否为矩形（对角线相等，四个角接近90度）
        is_rectangle = abs(diag1 - diag2) < avg_side * 0.1 and \
                       all(abs(a - 90) < self.ANGLE_TOLERANCE for a in angles)

        # 检查是否为正方形（菱形+矩形）
        if is_rhombus and is_rectangle:
            return ShapeType.DIAMOND, bbox  # 正方形可以映射为菱形

        if is_rhombus:
            return ShapeType.DIAMOND, bbox

        if is_rectangle:
            aspect_ratio = bbox.width / bbox.height if bbox.height > 0 else 1
            # 如果接近正方形，返回菱形
            if 0.8 < aspect_ratio < 1.2:
                return ShapeType.DIAMOND, bbox
            return ShapeType.RECTANGLE, bbox

        # 检查是否为梯形（至少一对边平行）
        # 简单检测：检查上下边是否平行
        def is_parallel(p1, p2, p3, p4):
            """检查两条线段是否平行"""
            dx1, dy1 = p2[0] - p1[0], p2[1] - p1[1]
            dx2, dy2 = p4[0] - p3[0], p4[1] - p3[1]
            if dx1 == 0 and dx2 == 0:
                return True
            if dx1 == 0 or dx2 == 0:
                return False
            return abs(dy1 / dx1 - dy2 / dx2) < 0.1

        if is_parallel(points[0], points[1], points[2], points[3]) or \
           is_parallel(points[1], points[2], points[3], points[0]):
            return ShapeType.TRAPEZOID, bbox

        return ShapeType.FREEFORM, bbox

    def _analyze_polygon(self, points: List[Tuple[float, float]], bbox: BoundingBox,
                        default_type: str) -> Tuple[str, BoundingBox]:
        """分析正多边形"""
        if self._is_regular_polygon(points, len(points)):
            return default_type, bbox
        return ShapeType.FREEFORM, bbox

    def parse_points(self, points_str: str) -> List[Tuple[float, float]]:
        """解析SVG points属性"""
        if not points_str:
            return []

        points = []
        # 支持格式: "x1,y1 x2,y2" 或 "x1 y1 x2 y2" 或 "x1,y1,x2,y2"
        # 首先尝试匹配 "x,y x,y" 格式（最常见）
        coords = re.findall(r'(-?\d+\.?\d*)[\s,]+(-?\d+\.?\d*)', points_str)
        for x, y in coords:
            try:
                points.append((float(x), float(y)))
            except ValueError:
                continue

        # 如果没解析到，尝试 "x y x y" 格式（空格分隔）
        if not points:
            # 提取所有数字
            numbers = re.findall(r'-?\d+\.?\d*', points_str)
            for i in range(0, len(numbers) - 1, 2):
                try:
                    x = float(numbers[i])
                    y = float(numbers[i + 1])
                    points.append((x, y))
                except (ValueError, IndexError):
                    continue

        return points

    def analyze_path(self, path_data: str) -> Tuple[str, BoundingBox]:
        """
        分析SVG路径数据
        简化实现：检测常见路径模式
        """
        if not path_data:
            return ShapeType.UNKNOWN, BoundingBox(0, 0, 100, 100)

        d = path_data.strip()

        # 提取所有坐标
        coords = re.findall(r'[MmLlHhVv]\s*(-?[\d.]+)[,\s]*(-?[\d.]+)?', d)

        if not coords:
            return ShapeType.FREEFORM, BoundingBox(0, 0, 100, 100)

        # 计算包围盒
        all_x, all_y = [], []
        for x, y in coords:
            try:
                all_x.append(float(x))
                if y:
                    all_y.append(float(y))
            except ValueError:
                continue

        if not all_x or not all_y:
            return ShapeType.FREEFORM, BoundingBox(0, 0, 100, 100)

        bbox = BoundingBox(min(all_x), min(all_y),
                          max(all_x) - min(all_x),
                          max(all_y) - min(all_y))

        # 检测矩形路径模式：M x y H x V y H x Z
        if re.match(r'^[Mm][\d\s.,-]+[Hh][\d\s.,-]+[Vv][\d\s.,-]+[Hh][\d\s.,-]+[Zz]$', d):
            return ShapeType.RECTANGLE, bbox

        # 检测圆/椭圆路径（包含A/a命令）
        if 'A' in d or 'a' in d:
            return ShapeType.OVAL, bbox

        return ShapeType.FREEFORM, bbox
