"""PDF 渲染器（reportlab）。

设计要点：
  * 与 DOCX / XLSX / PPTX 渲染器一样，**只消费 Document 中间模型**，
    新增格式不需要改动任何采集器与组装逻辑；
  * 字号、颜色全部来自 Theme，正文里不出现硬编码格式；
  * 中文使用 reportlab 自带的 Adobe CID 字体 ``STSong-Light``（无需外部字体文件，
    CI 上零额外依赖）；西文同样落在该字体上，保证中英混排观感一致；
  * 目录采用「多趟构建 + 页码回填」：第一趟记录各章节实际页码，
    后续趟把页码填回目录表，直到页码稳定（最多 3 趟），
    因此目录页码是真实页码，不是占位符。
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Dict, List, Optional
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    XPreformatted,
)

from ..charts import render_chart_png
from ..model import (
    Bullets,
    Chart,
    Code,
    Document,
    Metrics,
    Note,
    PageBreak as PageBreakBlock,
    Para,
    Table as TableBlock,
)
from ..theme import Theme

_CJK_FONT = "STSong-Light"
_FALLBACK_FONT = "Helvetica"
_FONT_READY = False

_NOTE_FILL = {
    "info": "EAF3FD",
    "warning": "FFF8E5",
    "danger": "FDEFF2",
    "success": "EAF6EC",
}
_NOTE_LABEL = {"info": "说明", "warning": "注意", "danger": "警告", "success": "结论"}

_MARGIN_X = 2.4 * cm
_MARGIN_TOP = 2.2 * cm
_MARGIN_BOTTOM = 2.0 * cm


# --------------------------------------------------------------------------
# 字体
# --------------------------------------------------------------------------
def _font() -> str:
    """注册并返回中文字体名；失败时退回 Helvetica（西文可渲染，中文会缺字形）。"""
    global _FONT_READY
    if _FONT_READY:
        return _CJK_FONT
    try:
        pdfmetrics.registerFont(UnicodeCIDFont(_CJK_FONT))
        _FONT_READY = True
        return _CJK_FONT
    except Exception:
        return _FALLBACK_FONT


def _color(hex_value: str):
    """'1F4E79' -> reportlab Color。"""
    return colors.HexColor("#" + hex_value.lstrip("#"))


def _esc(text: object) -> str:
    return escape(str(text))


# --------------------------------------------------------------------------
# 样式
# --------------------------------------------------------------------------
def _styles(theme: Theme, font: str) -> Dict[str, ParagraphStyle]:
    base = dict(fontName=font, wordWrap="CJK")
    return {
        "title": ParagraphStyle("dg-title", fontName=font, fontSize=theme.h1 + 8,
                                leading=theme.h1 + 14, textColor=_color(theme.primary),
                                alignment=TA_CENTER, wordWrap="CJK"),
        "subtitle": ParagraphStyle("dg-subtitle", fontName=font, fontSize=theme.h3 + 1,
                                   leading=theme.h3 + 7, textColor=_color(theme.muted),
                                   alignment=TA_CENTER, wordWrap="CJK"),
        "h1": ParagraphStyle("dg-h1", fontName=font, fontSize=theme.h1, leading=theme.h1 + 8,
                             textColor=_color(theme.primary), spaceBefore=14, spaceAfter=8,
                             wordWrap="CJK"),
        "h2": ParagraphStyle("dg-h2", fontName=font, fontSize=theme.h2, leading=theme.h2 + 6,
                             textColor=_color(theme.primary), spaceBefore=10, spaceAfter=6,
                             wordWrap="CJK"),
        "h3": ParagraphStyle("dg-h3", fontName=font, fontSize=theme.h3, leading=theme.h3 + 5,
                             textColor=_color(theme.text), spaceBefore=8, spaceAfter=4,
                             wordWrap="CJK"),
        "body": ParagraphStyle("dg-body", fontName=font, fontSize=theme.body,
                               leading=theme.body * 1.65, textColor=_color(theme.text),
                               spaceAfter=6, wordWrap="CJK"),
        "caption": ParagraphStyle("dg-caption", fontName=font, fontSize=theme.caption,
                                  leading=theme.caption * 1.6, textColor=_color(theme.muted),
                                  spaceAfter=4, wordWrap="CJK"),
        "cell": ParagraphStyle("dg-cell", fontName=font, fontSize=theme.caption,
                               leading=theme.caption * 1.5, textColor=_color(theme.text),
                               wordWrap="CJK"),
        "cell_head": ParagraphStyle("dg-cell-head", fontName=font, fontSize=theme.caption,
                                    leading=theme.caption * 1.5, textColor=colors.white,
                                    wordWrap="CJK"),
        "metric_value": ParagraphStyle("dg-metric", fontName=font, fontSize=theme.h2,
                                       leading=theme.h2 + 6, alignment=TA_CENTER,
                                       wordWrap="CJK"),
        "metric_label": ParagraphStyle("dg-metric-label", fontName=font, fontSize=theme.caption,
                                       leading=theme.caption * 1.5, alignment=TA_CENTER,
                                       textColor=_color(theme.muted), wordWrap="CJK"),
        "code": ParagraphStyle("dg-code", fontName=font, fontSize=theme.caption,
                               leading=theme.caption * 1.55, textColor=_color(theme.text),
                               wordWrap="CJK"),
        "note": ParagraphStyle("dg-note", fontName=font, fontSize=theme.caption,
                               leading=theme.caption * 1.6, wordWrap="CJK"),
        "bullet": ParagraphStyle("dg-bullet", fontName=font, fontSize=theme.body,
                                 leading=theme.body * 1.6, leftIndent=12, bulletIndent=2,
                                 spaceAfter=3, wordWrap="CJK"),
        "toc1": ParagraphStyle("dg-toc1", fontName=font, fontSize=theme.body,
                               leading=theme.body * 1.6, wordWrap="CJK"),
        "toc2": ParagraphStyle("dg-toc2", fontName=font, fontSize=theme.body - 0.5,
                               leading=theme.body * 1.5, leftIndent=16,
                               textColor=_color(theme.muted), wordWrap="CJK"),
        "toc3": ParagraphStyle("dg-toc3", fontName=font, fontSize=theme.caption,
                               leading=theme.caption * 1.6, leftIndent=32,
                               textColor=_color(theme.muted), wordWrap="CJK"),
        "base": ParagraphStyle("dg-base", **base),
    }


class _Heading(Paragraph):
    """会记录自己真实页码的标题，用于目录页码回填。"""

    def __init__(self, text: str, style: ParagraphStyle, key: str, recorder: Dict[str, int]):
        super().__init__(text, style)
        self._key = key
        self._recorder = recorder

    def draw(self) -> None:
        super().draw()
        self._recorder[self._key] = self.canv.getPageNumber()


# --------------------------------------------------------------------------
# 页面装饰
# --------------------------------------------------------------------------
def _make_decorator(document: Document, theme: Theme, font: str):
    def decorate(canvas, doc) -> None:
        canvas.saveState()
        canvas.setFont(font, theme.caption - 0.5)
        canvas.setFillColor(_color(theme.muted))
        width, height = A4
        # 页眉：品牌 + 文档标题；页脚：页码 / 总页数
        canvas.drawString(_MARGIN_X, height - _MARGIN_TOP + 0.5 * cm,
                          f"{theme.brand_name}  ·  {document.title}")
        canvas.setStrokeColor(_color(theme.border))
        canvas.setLineWidth(0.4)
        canvas.line(_MARGIN_X, height - _MARGIN_TOP + 0.35 * cm,
                    width - _MARGIN_X, height - _MARGIN_TOP + 0.35 * cm)
        canvas.line(_MARGIN_X, _MARGIN_BOTTOM - 0.35 * cm,
                    width - _MARGIN_X, _MARGIN_BOTTOM - 0.35 * cm)
        canvas.drawCentredString(
            width / 2.0, _MARGIN_BOTTOM - 0.8 * cm,
            f"第 {canvas.getPageNumber()} 页   |   {theme.footer_text}",
        )
        canvas.restoreState()

    def first_page(canvas, doc) -> None:
        # 封面不压页眉页脚，只保留底部品牌行
        canvas.saveState()
        canvas.setFont(font, theme.caption)
        canvas.setFillColor(_color(theme.muted))
        canvas.drawCentredString(A4[0] / 2.0, 1.2 * cm, theme.brand_name)
        canvas.restoreState()

    return first_page, decorate


# --------------------------------------------------------------------------
# 块渲染
# --------------------------------------------------------------------------
def _story_para(block: Para, styles, theme: Theme) -> List[object]:
    return [Paragraph(_esc(block.text), styles["body"])]


def _story_bullets(block: Bullets, styles, theme: Theme) -> List[object]:
    items = []
    for index, item in enumerate(block.items, start=1):
        bullet = f"{index}." if block.ordered else "•"
        items.append(Paragraph(_esc(item), styles["bullet"], bulletText=bullet))
    return items


def _story_table(block: TableBlock, styles, theme: Theme, width: float) -> List[object]:
    if not block.headers:
        return []
    story: List[object] = []
    if block.caption:
        story.append(Paragraph(_esc(block.caption), styles["caption"]))

    column_count = len(block.headers)
    col_width = width / max(1, column_count)
    data = [[Paragraph(_esc(h), styles["cell_head"]) for h in block.headers]]
    for row in block.rows:
        cells = []
        for column in range(column_count):
            value = row[column] if column < len(row) else ""
            align = TA_RIGHT if column in block.numeric_columns else TA_LEFT
            style = ParagraphStyle(
                f"cell-{column}", parent=styles["cell"], alignment=align,
                textColor=_color(theme.status_color(str(value).strip().lower()))
                if block.status_column == column else _color(theme.text),
            )
            cells.append(Paragraph(_esc(value), style))
        data.append(cells)

    table = Table(data, colWidths=[col_width] * column_count, repeatRows=1, hAlign="CENTER")
    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), _color(theme.primary)),
        ("GRID", (0, 0), (-1, -1), 0.4, _color(theme.border)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for row_index in range(1, len(data)):
        if row_index % 2 == 0:
            table_style.append(("BACKGROUND", (0, row_index), (-1, row_index),
                                _color(theme.surface)))
    table.setStyle(TableStyle(table_style))
    story.append(table)
    story.append(Spacer(1, 8))
    return story


def _story_metrics(block: Metrics, styles, theme: Theme, width: float) -> List[object]:
    if not block.items:
        return []
    col_width = width / len(block.items)
    values, labels = [], []
    for metric in block.items:
        value_style = ParagraphStyle(
            "mv", parent=styles["metric_value"],
            textColor=_color(theme.status_color(metric.status)),
        )
        values.append(Paragraph(f"<b>{_esc(metric.value)}</b>", value_style))
        label = metric.label if not metric.note else f"{metric.label}<br/>{_esc(metric.note)}"
        labels.append(Paragraph(label, styles["metric_label"]))

    table = Table([values, labels], colWidths=[col_width] * len(block.items), hAlign="CENTER")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _color(theme.surface)),
        ("GRID", (0, 0), (-1, -1), 0.4, _color(theme.border)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return [table, Spacer(1, 10)]


def _story_note(block: Note, styles, theme: Theme, width: float) -> List[object]:
    label = _NOTE_LABEL.get(block.level, "说明")
    text_color = theme.status_color(
        "danger" if block.level == "danger" else
        "warning" if block.level == "warning" else
        "success" if block.level == "success" else "muted"
    )
    style = ParagraphStyle("note-text", parent=styles["note"], textColor=_color(text_color))
    table = Table([[Paragraph(f"<b>{label}</b>：{_esc(block.text)}", style)]],
                  colWidths=[width], hAlign="CENTER")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _color(_NOTE_FILL.get(block.level, "EAF3FD"))),
        ("BOX", (0, 0), (-1, -1), 0.4, _color(theme.border)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return [table, Spacer(1, 8)]


def _story_code(block: Code, styles, theme: Theme, width: float) -> List[object]:
    text = block.text.rstrip("\n")
    block_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _color(theme.surface)),
        ("BOX", (0, 0), (-1, -1), 0.4, _color(theme.border)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ])
    table = Table([[XPreformatted(text, styles["code"])]], colWidths=[width], hAlign="CENTER")
    table.setStyle(block_style)
    return [table, Spacer(1, 8)]


def _story_chart(block: Chart, styles, theme: Theme, width: float) -> List[object]:
    try:
        png = render_chart_png(block, theme)
    except Exception as error:  # 缺 matplotlib 或字体时降级为文字，不让整份文档失败
        return _story_note(Note(text=f"图表「{block.title}」渲染失败：{error}", level="warning"),
                           styles, theme, width)
    image = Image(io.BytesIO(png))
    ratio = image.imageHeight / float(image.imageWidth or 1)
    image.drawWidth = min(width, 16 * cm)
    image.drawHeight = image.drawWidth * ratio
    image.hAlign = "CENTER"
    story: List[object] = [image]
    if block.title:
        caption_style = ParagraphStyle("chart-caption", parent=styles["caption"],
                                       alignment=TA_CENTER)
        story.append(Paragraph(_esc(f"图：{block.title}"), caption_style))
    story.append(Spacer(1, 8))
    return story


# --------------------------------------------------------------------------
# 文档骨架
# --------------------------------------------------------------------------
def _key_value_table(rows: List[tuple], styles, theme: Theme, label_width: float) -> Table:
    data = [[Paragraph(_esc(k), ParagraphStyle("kv-k", parent=styles["cell"])),
             Paragraph(_esc(v), styles["cell"])] for k, v in rows]
    table = Table(data, colWidths=[label_width, 13.6 * cm - label_width], hAlign="CENTER")
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, _color(theme.border)),
        ("BACKGROUND", (0, 0), (0, -1), _color(theme.surface)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return table


def _cover(story: List[object], document: Document, theme: Theme, styles) -> None:
    story.append(Spacer(1, 2.6 * cm))
    story.append(Paragraph(_esc(theme.organization), styles["subtitle"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph(_esc(theme.brand_name), styles["caption"]))
    story.append(Spacer(1, 3.4 * cm))
    story.append(Paragraph(_esc(document.title), styles["title"]))
    if document.subtitle:
        story.append(Spacer(1, 0.6 * cm))
        story.append(Paragraph(_esc(document.subtitle), styles["subtitle"]))
    story.append(Spacer(1, 2.6 * cm))

    # 对外交付件：封面顶部客户抬头
    if document.delivery is not None:
        delivery = document.delivery
        head_rows = [(k, v) for k, v in [
            ("致", delivery.client),
            ("项目名称", delivery.project),
            ("客户方对接人", delivery.contact),
            ("交付方负责人", delivery.owner),
            ("生效日期", delivery.effective_date),
        ] if v]
        if head_rows:
            story.append(_key_value_table(head_rows, styles, theme, 3.6 * cm))
            story.append(Spacer(1, 0.8 * cm))

    meta = [
        ("文档编号", document.doc_id),
        ("版本", document.version),
        ("密级", document.classification),
        ("编制", document.author),
        ("所属部门", document.department),
        ("生成时间", document.generated_at),
        ("数据来源", document.source_commit),
    ]
    meta = [(key, value) for key, value in meta if value]
    if meta:
        data = [[Paragraph(_esc(k), styles["cell"]), Paragraph(_esc(v), styles["cell"])]
                for k, v in meta]
        table = Table(data, colWidths=[4.0 * cm, 9.6 * cm], hAlign="CENTER")
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, _color(theme.border)),
            ("BACKGROUND", (0, 0), (0, -1), _color(theme.surface)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)


def _toc_block(entries: List[tuple], pages: Dict[str, int], styles, theme: Theme,
               width: float) -> List[object]:
    """渲染目录；pages 为空时页码显示为「—」。"""
    if not entries:
        return []
    data = []
    for key, title, level in entries:
        style = styles["toc1"] if level <= 1 else styles["toc2"] if level == 2 else styles["toc3"]
        page = pages.get(key)
        data.append([Paragraph(_esc(title), style),
                     Paragraph(str(page) if page else "—",
                               ParagraphStyle("toc-page", parent=style, alignment=TA_RIGHT))])
    table = Table(data, colWidths=[width - 2.0 * cm, 2.0 * cm], hAlign="CENTER")
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, -2), 0.2, _color(theme.border)),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return [table]


def _compose(document: Document, theme: Theme, styles, font: str, width: float,
             recorder: Dict[str, int], pages: Dict[str, int]) -> List[object]:
    story: List[object] = []
    _cover(story, document, theme, styles)
    story.append(PageBreak())

    # ---- 交付说明（抬头页） ----
    if document.delivery is not None:
        delivery = document.delivery
        story.append(_Heading("交付说明", styles["h1"], "__letterhead__", recorder))
        head_rows = [(k, v) for k, v in [
            ("接收方", delivery.client),
            ("项目名称", delivery.project),
            ("客户方对接人", delivery.contact),
            ("交付方负责人", delivery.owner),
            ("生效日期", delivery.effective_date),
            ("保密约定", delivery.confidential),
        ] if v]
        if head_rows:
            story.append(_key_value_table(head_rows, styles, theme, 3.6 * cm))
        story.append(PageBreak())

    if document.summary:
        story.append(_Heading("摘要", styles["h1"], "__summary__", recorder))
        story.append(Paragraph(_esc(document.summary), styles["body"]))
        story.append(Spacer(1, 6))

    if document.highlights:
        story.append(_Heading("关键结论", styles["h1"], "__highlights__", recorder))
        story.extend(_story_bullets(Bullets(items=document.highlights), styles, theme))
        story.append(Spacer(1, 6))

    # ---- 目录 ----
    entries: List[tuple] = []
    if document.delivery is not None:
        entries.append(("__letterhead__", "交付说明", 1))
    if document.summary:
        entries.append(("__summary__", "摘要", 1))
    if document.highlights:
        entries.append(("__highlights__", "关键结论", 1))
    for index, section in enumerate(document.sections):
        if section.title:
            entries.append((f"sec-{index}", section.title, section.level))
    if document.delivery is not None:
        entries.append(("__terms__", "条款与约定", 1))
        entries.append(("__signature__", "签署页", 1))

    story.append(_Heading("目录", styles["h1"], "__toc__", recorder))
    story.extend(_toc_block(entries, pages, styles, theme, width))
    story.append(PageBreak())

    # ---- 正文 ----
    for index, section in enumerate(document.sections):
        if section.title:
            style_key = {1: "h1", 2: "h2", 3: "h3"}.get(section.level, "h2")
            story.append(_Heading(_esc(section.title), styles[style_key],
                                  f"sec-{index}", recorder))
        for block in section.blocks:
            if isinstance(block, Para):
                story.extend(_story_para(block, styles, theme))
            elif isinstance(block, Bullets):
                story.extend(_story_bullets(block, styles, theme))
            elif isinstance(block, TableBlock):
                story.extend(_story_table(block, styles, theme, width))
            elif isinstance(block, Metrics):
                story.extend(_story_metrics(block, styles, theme, width))
            elif isinstance(block, Note):
                story.extend(_story_note(block, styles, theme, width))
            elif isinstance(block, Code):
                story.extend(_story_code(block, styles, theme, width))
            elif isinstance(block, Chart):
                story.extend(_story_chart(block, styles, theme, width))
            elif isinstance(block, PageBreakBlock):
                story.append(PageBreak())

    # ---- 条款页 + 签署页（仅对外交付件） ----
    if document.delivery is not None:
        delivery = document.delivery
        story.append(PageBreak())
        story.append(_Heading("条款与约定", styles["h1"], "__terms__", recorder))
        if delivery.terms:
            story.extend(_story_bullets(Bullets(items=delivery.terms, ordered=True),
                                        styles, theme))
        if delivery.confidential:
            story.extend(_story_note(Note(text=delivery.confidential, level="warning"),
                                     styles, theme, width))
        if delivery.warranty:
            story.extend(_story_note(Note(text=delivery.warranty, level="info"),
                                     styles, theme, width))

        story.append(PageBreak())
        story.append(_Heading("签署页", styles["h1"], "__signature__", recorder))
        story.append(Paragraph(
            _esc("本文件经双方确认后生效，签署即代表接受本文件全部条款与约定。"),
            styles["body"]))
        story.append(Spacer(1, 8))
        roles = delivery.signatures or ["客户方负责人（签章）", "交付方负责人（签章）"]
        data = [[Paragraph(_esc(t), ParagraphStyle("sig-h", parent=styles["cell_head"]))
                 for t in ("签署角色", "签字 / 盖章", "日期")]]
        for role in roles:
            data.append([Paragraph(_esc(role), styles["cell"]),
                         Paragraph("", styles["cell"]),
                         Paragraph("", styles["cell"])])
        table = Table(data, colWidths=[width * 0.4, width * 0.35, width * 0.25], hAlign="CENTER")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), _color(theme.primary)),
            ("GRID", (0, 0), (-1, -1), 0.4, _color(theme.border)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(table)
    return story


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def render_pdf(document: Document, out_path: Path) -> Path:
    """渲染为 .pdf 文件。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    theme = document.theme
    font = _font()
    styles = _styles(theme, font)
    first_page, later_pages = _make_decorator(document, theme, font)

    pages: Dict[str, int] = {}
    previous: Dict[str, int] = {}
    recorder: Dict[str, int] = {}

    # 多趟构建：第一趟拿到页码，第二趟回填到目录，直到页码稳定（最多 3 趟）
    for _ in range(3):
        recorder = {}
        doc = SimpleDocTemplate(
            str(out_path),
            pagesize=A4,
            leftMargin=_MARGIN_X,
            rightMargin=_MARGIN_X,
            topMargin=_MARGIN_TOP,
            bottomMargin=_MARGIN_BOTTOM,
            title=document.title,
            author=document.author or theme.organization,
            subject=document.subtitle or "",
        )
        story = _compose(document, theme, styles, font, doc.width, recorder, pages)
        doc.build(story, onFirstPage=first_page, onLaterPages=later_pages)
        pages = dict(recorder)
        if pages == previous:
            break
        previous = dict(pages)

    return out_path
