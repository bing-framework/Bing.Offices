"""品牌主题。

所有渲染器只从 Theme 取值，禁止在渲染代码里硬编码字体、字号与颜色。
换主题 = 换这一个文件，不动任何渲染逻辑。
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Tuple


def hex_to_rgb(value: str) -> Tuple[int, int, int]:
    """'1F4E79' -> (31, 78, 121)"""
    v = value.lstrip("#")
    return int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)


def hex_to_rgb_float(value: str) -> Tuple[float, float, float]:
    r, g, b = hex_to_rgb(value)
    return r / 255.0, g / 255.0, b / 255.0


@dataclass(frozen=True)
class Theme:
    """文档品牌主题。"""

    brand_name: str = "Bing.Offices"
    organization: str = "Bing Framework"
    footer_text: str = "Bing.Offices 文档自动化流水线生成"

    # 颜色（不含 # 的 RRGGBB）
    primary: str = "1F4E79"
    accent: str = "2E75B6"
    text: str = "1F2328"
    muted: str = "59636E"
    surface: str = "F2F5F9"
    border: str = "D1D9E0"
    success: str = "1A7F37"
    danger: str = "B3254A"
    warning: str = "9A6700"

    # 字体
    font_heading: str = "微软雅黑"
    font_body: str = "微软雅黑"
    font_mono: str = "Consolas"

    # 字号（pt）
    h1: float = 22.0
    h2: float = 15.0
    h3: float = 12.5
    body: float = 10.5
    caption: float = 9.0

    # 图表配色（按顺序取用）
    chart_palette: List[str] = field(
        default_factory=lambda: ["2E75B6", "1F4E79", "5B9BD5", "9DC3E6", "8FAADC", "B4C7E7"]
    )

    def status_color(self, status: str) -> str:
        """状态 -> 颜色。"""
        return {
            "success": self.success,
            "danger": self.danger,
            "warning": self.warning,
            "neutral": self.muted,
        }.get(status, self.muted)

    def series_color(self, index: int) -> str:
        """第 index 个数据系列的颜色。"""
        palette = self.chart_palette
        return palette[index % len(palette)]

    def to_dict(self) -> Dict[str, object]:
        return asdict(self)


DEFAULT_THEME = Theme()
