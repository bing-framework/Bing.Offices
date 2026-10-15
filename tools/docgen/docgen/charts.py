"""位图图表渲染（供 DOCX / Markdown 使用）。

设计取舍：
  * PPTX 与 XLSX 使用 **Office 原生图表**（可编辑、可换数据），因此不走这里；
  * DOCX 与 Markdown 没有原生图表能力，才用 matplotlib 出位图。

禁止在渲染器里硬编码颜色，全部取自 Theme。
"""

from __future__ import annotations

import io
from typing import List, Sequence

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt  # noqa: E402

from .model import Chart  # noqa: E402
from .theme import Theme, hex_to_rgb_float  # noqa: E402

_CJK_CANDIDATES = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "PingFang SC", "DejaVu Sans"]


def _configure_fonts() -> None:
    """让 matplotlib 能正确渲染中文（含负号）。"""
    plt.rcParams["font.sans-serif"] = _CJK_CANDIDATES
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["axes.unicode_minus"] = False


def render_chart_png(
    chart: Chart,
    theme: Theme,
    width_inches: float = 9.0,
    height_inches: float = 4.4,
    dpi: int = 200,
) -> bytes:
    """把 Chart 渲染为 PNG 字节流。"""
    _configure_fonts()

    if not chart.categories or not chart.series:
        raise ValueError(f"图表「{chart.title}」缺少类别或数据系列")

    figure, axes = plt.subplots(figsize=(width_inches, height_inches), dpi=dpi)
    figure.patch.set_facecolor("white")

    categories = chart.categories
    positions = range(len(categories))

    if chart.kind == "pie":
        name, values = chart.series[0]
        axes.pie(
            values,
            labels=categories,
            autopct="%1.1f%%",
            colors=[hex_to_rgb_float(theme.series_color(i)) for i in range(len(categories))],
            textprops={"fontsize": 9},
        )
        axes.axis("equal")
    elif chart.kind == "line":
        for index, (name, values) in enumerate(chart.series):
            axes.plot(
                list(positions),
                values,
                marker="o",
                linewidth=2.0,
                markersize=4,
                label=name,
                color=hex_to_rgb_float(theme.series_color(index)),
            )
        axes.set_xticks(list(positions))
        axes.set_xticklabels(categories, rotation=20, ha="right", fontsize=9)
        if len(chart.series) > 1:
            axes.legend(fontsize=9, frameon=False)
    else:
        count = len(chart.series)
        slot = 0.8 / max(count, 1)
        if chart.horizontal:
            for index, (name, values) in enumerate(chart.series):
                offsets = [p + index * slot - 0.4 + slot / 2 for p in positions]
                axes.barh(
                    offsets,
                    values,
                    height=slot,
                    label=name,
                    color=hex_to_rgb_float(theme.series_color(index)),
                )
            axes.set_yticks(list(positions))
            axes.set_yticklabels(categories, fontsize=9)
            if count > 1:
                axes.legend(fontsize=9, frameon=False)
        else:
            for index, (name, values) in enumerate(chart.series):
                offsets = [p + index * slot - 0.4 + slot / 2 for p in positions]
                axes.bar(
                    offsets,
                    values,
                    width=slot,
                    label=name,
                    color=hex_to_rgb_float(theme.series_color(index)),
                )
            axes.set_xticks(list(positions))
            axes.set_xticklabels(categories, rotation=20, ha="right", fontsize=9)
            if count > 1:
                axes.legend(fontsize=9, frameon=False)

    if chart.title and chart.kind != "pie":
        axes.set_title(chart.title, fontsize=11, color=hex_to_rgb_float(theme.text), pad=10)
    if chart.unit and chart.kind != "pie":
        if chart.horizontal:
            axes.set_xlabel(chart.unit, fontsize=9, color=hex_to_rgb_float(theme.muted))
        else:
            axes.set_ylabel(chart.unit, fontsize=9, color=hex_to_rgb_float(theme.muted))

    if chart.kind != "pie":
        axes.grid(axis="y" if not chart.horizontal else "x", linestyle=":", alpha=0.35)
        axes.set_axisbelow(True)
        for side in ("top", "right"):
            axes.spines[side].set_visible(False)
        for side in ("left", "bottom"):
            axes.spines[side].set_color(hex_to_rgb_float(theme.border))
        axes.tick_params(colors=hex_to_rgb_float(theme.muted), labelsize=9)

    figure.tight_layout()
    buffer = io.BytesIO()
    figure.savefig(buffer, format="png", facecolor="white")
    plt.close(figure)
    return buffer.getvalue()
