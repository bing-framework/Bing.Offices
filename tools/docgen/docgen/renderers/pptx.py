"""PPTX 渲染器（python-pptx）。

设计要点：
  * 16:9，全篇使用空白版式 + 手工文本框，保证排版完全可控、不依赖母版；
  * 图表使用 **Office 原生图表**，汇报人可以现场改数据；
  * 每个章节最多承载有限内容块，超出的自动分页，避免文字溢出幻灯片。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Sequence

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from ..model import (
    Block,
    Bullets,
    Chart,
    Code,
    Document,
    Metrics,
    Note,
    PageBreak,
    Para,
    Section,
    Table,
)
from ..theme import Theme

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)
MARGIN = Inches(0.62)
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN * 2
TITLE_TOP = Inches(0.42)
CONTENT_TOP = Inches(1.55)
CONTENT_HEIGHT = Inches(5.35)
FOOTER_TOP = Inches(7.02)

_CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}


# --------------------------------------------------------------------------
def _textbox(slide, text: str, left, top, width, height, theme: Theme,
             size: Optional[float] = None, bold: bool = False, color: Optional[str] = None,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing: float = 1.15):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.line_spacing = line_spacing
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size if size is not None else theme.body)
    run.font.bold = bold
    run.font.name = theme.font_heading if bold else theme.font_body
    run.font.color.rgb = RGBColor.from_string(color or theme.text)
    return box


def _rect(slide, left, top, width, height, fill: str, shape=MSO_SHAPE.RECTANGLE, line: Optional[str] = None):
    box = slide.shapes.add_shape(shape, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        box.line.color.rgb = RGBColor.from_string(line)
        box.line.width = Pt(0.75)
    else:
        box.line.fill.background()
    box.shadow.inherit = False
    return box


def _blank_slide(prs, document: Document):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme = document.theme
    # 顶部品牌色条
    _rect(slide, Emu(0), Emu(0), SLIDE_WIDTH, Inches(0.09), theme.primary)
    # 页脚
    _textbox(slide, f"{theme.brand_name}  ·  {document.title}", MARGIN, FOOTER_TOP,
             CONTENT_WIDTH - Inches(1.2), Inches(0.32), theme, size=theme.caption,
             color=theme.muted)
    return slide


def _slide_title(slide, text: str, theme: Theme, subtitle: str = "") -> None:
    _textbox(slide, text, MARGIN, TITLE_TOP, CONTENT_WIDTH, Inches(0.62), theme,
             size=theme.h1, bold=True, color=theme.primary)
    _rect(slide, MARGIN, TITLE_TOP + Inches(0.66), Inches(1.5), Inches(0.035), theme.accent)
    if subtitle:
        _textbox(slide, subtitle, MARGIN, TITLE_TOP + Inches(0.76), CONTENT_WIDTH,
                 Inches(0.32), theme, size=theme.caption, color=theme.muted)


# --------------------------------------------------------------------------
# 幻灯片类型
# --------------------------------------------------------------------------
def _add_cover(prs, document: Document) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme = document.theme
    _rect(slide, Emu(0), Emu(0), SLIDE_WIDTH, SLIDE_HEIGHT, theme.primary)
    _rect(slide, Emu(0), Inches(4.55), SLIDE_WIDTH, Inches(0.05), theme.accent)

    _textbox(slide, theme.organization, MARGIN, Inches(1.55), CONTENT_WIDTH, Inches(0.4),
             theme, size=theme.h3, bold=True, color="9DC3E6")
    _textbox(slide, document.title, MARGIN, Inches(2.15), CONTENT_WIDTH, Inches(1.5),
             theme, size=38.0, bold=True, color="FFFFFF")
    if document.subtitle:
        _textbox(slide, document.subtitle, MARGIN, Inches(3.62), CONTENT_WIDTH, Inches(0.5),
                 theme, size=theme.h2, color="D6E4F0")

    meta_bits = [b for b in (document.version and f"版本 {document.version}",
                             document.generated_at,
                             document.classification) if b]
    _textbox(slide, "   ·   ".join(meta_bits), MARGIN, Inches(4.85), CONTENT_WIDTH,
             Inches(0.4), theme, size=theme.caption, color="D6E4F0")


def _add_agenda(prs, document: Document) -> None:
    sections = [s for s in document.sections if s.title]
    if len(sections) < 3:
        return
    theme = document.theme
    slide = _blank_slide(prs, document)
    _slide_title(slide, "目录", theme)

    columns = 2
    per_column = (len(sections) + columns - 1) // columns
    column_width = (CONTENT_WIDTH - Inches(0.4)) / columns
    for index, section in enumerate(sections):
        column = index // per_column
        row = index % per_column
        left = MARGIN + column * (column_width + Inches(0.4))
        top = CONTENT_TOP + Inches(row * 0.62)
        number = _rect(slide, left, top + Inches(0.04), Inches(0.34), Inches(0.34),
                       document.theme.accent, MSO_SHAPE.OVAL)
        frame = number.text_frame
        frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = frame.paragraphs[0].add_run()
        run.text = str(index + 1)
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string("FFFFFF")
        _textbox(slide, section.title, left + Inches(0.5), top, column_width - Inches(0.5),
                 Inches(0.4), document.theme, size=theme.body)


def _add_metrics_slide(prs, document: Document, section: Section, block: Metrics) -> None:
    slide = _blank_slide(prs, document)
    theme = document.theme
    _slide_title(slide, section.title, theme)

    count = max(1, min(len(block.items), 4))
    gap = Inches(0.28)
    card_width = (CONTENT_WIDTH - gap * (count - 1)) / count
    for index, metric in enumerate(block.items[:count]):
        left = MARGIN + index * (card_width + gap)
        _rect(slide, left, CONTENT_TOP + Inches(0.35), card_width, Inches(2.5),
              theme.surface, MSO_SHAPE.ROUNDED_RECTANGLE, theme.border)
        _rect(slide, left, CONTENT_TOP + Inches(0.35), card_width, Inches(0.075),
              theme.status_color(metric.status))
        _textbox(slide, metric.value, left + Inches(0.15), CONTENT_TOP + Inches(0.75),
                 card_width - Inches(0.3), Inches(0.9), theme, size=30.0, bold=True,
                 color=theme.status_color(metric.status), align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        _textbox(slide, metric.label, left + Inches(0.15), CONTENT_TOP + Inches(1.75),
                 card_width - Inches(0.3), Inches(0.45), theme, size=theme.body, bold=True,
                 align=PP_ALIGN.CENTER)
        if metric.note:
            _textbox(slide, metric.note, left + Inches(0.15), CONTENT_TOP + Inches(2.18),
                     card_width - Inches(0.3), Inches(0.45), theme, size=theme.caption,
                     color=theme.muted, align=PP_ALIGN.CENTER)


def _add_table_slide(prs, document: Document, section: Section, block: Table) -> None:
    slide = _blank_slide(prs, document)
    theme = document.theme
    _slide_title(slide, section.title, theme, block.caption)

    headers = block.headers or []
    rows = block.rows[:14]
    if not headers:
        return

    shape = slide.shapes.add_table(len(rows) + 1, len(headers), MARGIN, CONTENT_TOP + Inches(0.15),
                                   CONTENT_WIDTH, Inches(0.4) * (len(rows) + 1))
    table = shape.table
    for column, header in enumerate(headers):
        cell = table.cell(0, column)
        cell.text = str(header)
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(theme.caption)
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string("FFFFFF")
            run.font.name = theme.font_heading
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(theme.primary)

    for row_index, row in enumerate(rows):
        for column in range(len(headers)):
            cell = table.cell(row_index + 1, column)
            value = str(row[column]) if column < len(row) else ""
            cell.text = value
            paragraph = cell.text_frame.paragraphs[0]
            for run in paragraph.runs:
                run.font.size = Pt(theme.caption)
                run.font.name = theme.font_body
                if block.status_column == column:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor.from_string(
                        theme.status_color(value.strip().lower()))
                else:
                    run.font.color.rgb = RGBColor.from_string(theme.text)
            if row_index % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(theme.surface)

    if len(block.rows) > len(rows):
        _textbox(slide, f"（共 {len(block.rows)} 行，此处展示前 {len(rows)} 行，完整数据见附带表格文件）",
                 MARGIN, CONTENT_TOP + Inches(0.15) + Inches(0.4) * (len(rows) + 1) + Inches(0.1),
                 CONTENT_WIDTH, Inches(0.3), theme, size=theme.caption, color=theme.muted)


def _add_chart_slide(prs, document: Document, title: str, chart: Chart) -> None:
    slide = _blank_slide(prs, document)
    theme = document.theme
    _slide_title(slide, title, theme, chart.title if chart.title != title else "")

    if not chart.categories or not chart.series:
        _textbox(slide, "无图表数据", MARGIN, CONTENT_TOP, CONTENT_WIDTH, Inches(0.5),
                 theme, color=theme.muted)
        return

    data = CategoryChartData()
    data.categories = chart.categories
    for name, values in chart.series:
        data.add_series(name, values)

    chart_type = _CHART_TYPES.get(chart.kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
    if chart.kind == "bar" and chart.horizontal:
        chart_type = XL_CHART_TYPE.BAR_CLUSTERED

    frame = slide.shapes.add_chart(chart_type, MARGIN, CONTENT_TOP + Inches(0.1),
                                   CONTENT_WIDTH, CONTENT_HEIGHT - Inches(0.3), data)
    native = frame.chart
    native.has_title = False
    native.font.size = Pt(theme.caption)
    native.font.name = theme.font_body

    if chart.kind == "pie":
        native.has_legend = True
        native.legend.position = XL_LEGEND_POSITION.RIGHT
        native.legend.include_in_layout = False
        plot = native.plots[0]
        plot.has_data_labels = True
        plot.data_labels.show_percentage = True
        plot.data_labels.font.size = Pt(theme.caption)
    else:
        has_legend = len(chart.series) > 1
        native.has_legend = has_legend
        if has_legend:
            native.legend.position = XL_LEGEND_POSITION.BOTTOM
            native.legend.include_in_layout = False
        value_axis = native.value_axis
        value_axis.has_major_gridlines = True
        value_axis.major_gridlines.format.line.color.rgb = RGBColor.from_string(theme.border)
        if chart.unit:
            try:
                value_axis.axis_title.text_frame.text = chart.unit
                value_axis.has_title = True
            except (AttributeError, ValueError):
                # 部分图表类型/字体环境下轴标题不可用，静默跳过而不是让整份文档失败
                pass

    for index in range(len(chart.series)):
        series = native.series[index]
        color = RGBColor.from_string(theme.series_color(index))
        if chart.kind == "line":
            series.format.line.color.rgb = color
            series.format.line.width = Pt(2.25)
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color


def _add_content_slide(prs, document: Document, section: Section, blocks: Sequence[Block]) -> None:
    slide = _blank_slide(prs, document)
    theme = document.theme
    _slide_title(slide, section.title, theme)

    cursor = CONTENT_TOP + Inches(0.15)
    for block in blocks:
        if isinstance(block, Para):
            box = _textbox(slide, block.text, MARGIN, cursor, CONTENT_WIDTH, Inches(0.9),
                           theme, size=theme.body)
            cursor += Inches(0.32) + Inches(0.2) * max(0, len(block.text) // 70)
        elif isinstance(block, Bullets):
            for item in block.items:
                bullet = _rect(slide, MARGIN + Inches(0.06), cursor + Inches(0.11),
                               Inches(0.075), Inches(0.075), theme.accent, MSO_SHAPE.OVAL)
                bullet.text_frame.text = ""
                _textbox(slide, item, MARGIN + Inches(0.32), cursor,
                         CONTENT_WIDTH - Inches(0.32), Inches(0.4), theme, size=theme.body)
                cursor += Inches(0.46)
            cursor += Inches(0.1)
        elif isinstance(block, Note):
            label = {"info": "说明", "warning": "注意", "danger": "警告", "success": "结论"}.get(
                block.level, "说明")
            fill = {"info": "EAF3FD", "warning": "FFF8E5", "danger": "FDEFF2",
                    "success": "EAF6EC"}.get(block.level, "EAF3FD")
            _rect(slide, MARGIN, cursor, CONTENT_WIDTH, Inches(0.62), fill,
                  MSO_SHAPE.ROUNDED_RECTANGLE,
                  theme.status_color(block.level if block.level != "info" else "neutral"))
            _textbox(slide, f"{label}：{block.text}", MARGIN + Inches(0.18),
                     cursor + Inches(0.09), CONTENT_WIDTH - Inches(0.36), Inches(0.5),
                     theme, size=theme.caption, anchor=MSO_ANCHOR.MIDDLE)
            cursor += Inches(0.8)
        elif isinstance(block, Code):
            _rect(slide, MARGIN, cursor, CONTENT_WIDTH, Inches(0.9), theme.surface,
                  MSO_SHAPE.RECTANGLE, theme.border)
            _textbox(slide, block.text[:400], MARGIN + Inches(0.15), cursor + Inches(0.1),
                     CONTENT_WIDTH - Inches(0.3), Inches(0.7), theme, size=theme.caption,
                     color=theme.text)
            cursor += Inches(1.05)
        if cursor > SLIDE_HEIGHT - Inches(1.2):
            break


# --------------------------------------------------------------------------
def render_pptx(document: Document, out_path: Path) -> Path:
    """渲染为 .pptx 文件。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    _add_cover(prs, document)

    # 摘要页
    if document.summary or document.highlights:
        slide = _blank_slide(prs, document)
        _slide_title(slide, "摘要与关键结论", document.theme)
        cursor = CONTENT_TOP + Inches(0.1)
        if document.summary:
            _textbox(slide, document.summary, MARGIN, cursor, CONTENT_WIDTH, Inches(1.0),
                     document.theme, size=document.theme.body)
            cursor += Inches(1.15)
        for item in document.highlights[:6]:
            bullet = _rect(slide, MARGIN + Inches(0.06), cursor + Inches(0.11),
                           Inches(0.075), Inches(0.075), document.theme.accent, MSO_SHAPE.OVAL)
            bullet.text_frame.text = ""
            _textbox(slide, item, MARGIN + Inches(0.32), cursor,
                     CONTENT_WIDTH - Inches(0.32), Inches(0.4), document.theme,
                     size=document.theme.body)
            cursor += Inches(0.46)

    _add_agenda(prs, document)

    for section in document.sections:
        if not section.title:
            continue
        text_blocks: List[Block] = []
        for block in section.blocks:
            if isinstance(block, Chart):
                if text_blocks:
                    _add_content_slide(prs, document, section, text_blocks)
                    text_blocks = []
                _add_chart_slide(prs, document, section.title, block)
            elif isinstance(block, Metrics):
                if text_blocks:
                    _add_content_slide(prs, document, section, text_blocks)
                    text_blocks = []
                _add_metrics_slide(prs, document, section, block)
            elif isinstance(block, Table):
                if text_blocks:
                    _add_content_slide(prs, document, section, text_blocks)
                    text_blocks = []
                _add_table_slide(prs, document, section, block)
            elif isinstance(block, PageBreak):
                continue
            else:
                text_blocks.append(block)
        if text_blocks:
            _add_content_slide(prs, document, section, text_blocks)

    # 收尾页
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    theme = document.theme
    _rect(slide, Emu(0), Emu(0), SLIDE_WIDTH, SLIDE_HEIGHT, theme.primary)
    _textbox(slide, "谢谢", MARGIN, Inches(2.9), CONTENT_WIDTH, Inches(1.0), theme,
             size=36.0, bold=True, color="FFFFFF", align=PP_ALIGN.CENTER)
    _textbox(slide, f"{theme.brand_name} · {theme.organization}", MARGIN, Inches(4.05),
             CONTENT_WIDTH, Inches(0.45), theme, size=theme.body, color="D6E4F0",
             align=PP_ALIGN.CENTER)

    prs.save(str(out_path))
    return out_path
