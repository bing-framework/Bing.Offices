"""DOCX 渲染器（python-docx）。

设计要点：
  * **只使用文档样式** —— 字号、字体、颜色全部来自 Theme，正文里不出现硬编码格式；
  * 封面 + 目录域 + 页眉页脚（页码域）构成一份可直接交付的正式文档；
  * DOCX 无原生图表能力，图表以 matplotlib PNG 嵌入；
  * 表格使用 ``Table Grid`` 样式并显式设置表头底纹，保证跨平台观感一致。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from docx import Document as DocxDocument
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from ..charts import render_chart_png
from ..model import (
    Block,
    Bullets,
    Chart,
    Code,
    Delivery,
    Document,
    Metrics,
    Note,
    PageBreak,
    Para,
    Table,
)
from ..theme import Theme, hex_to_rgb

_NOTE_FILL = {
    "info": "EAF3FD",
    "warning": "FFF8E5",
    "danger": "FDEFF2",
    "success": "EAF6EC",
}
_NOTE_LABEL = {"info": "说明", "warning": "注意", "danger": "警告", "success": "结论"}


# --------------------------------------------------------------------------
# 低层工具
# --------------------------------------------------------------------------
def _apply_cjk_font(run, theme: Theme, name: Optional[str] = None, mono: bool = False) -> None:
    """为 run 设置西文 + 中日韩字体，缺一不可否则中文会回退。"""
    font_name = name or (theme.font_mono if mono else theme.font_body)
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attribute), font_name)


def _style_font(style, theme: Theme, name: str, size: float, bold: Optional[bool] = None,
                color: Optional[str] = None) -> None:
    style.font.name = name
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color:
        style.font.color.rgb = RGBColor.from_string(color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attribute), name)


def _shade(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _cell_text(cell, text: str, theme: Theme, size: float, bold: bool = False,
               color: Optional[str] = None, mono: bool = False,
               align: Optional[int] = None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(2)
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run(str(text))
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    _apply_cjk_font(run, theme, mono=mono)


def _field(paragraph, code: str, placeholder: str, theme: Theme, size: float) -> None:
    """插入一个 Word 域（如 PAGE / NUMPAGES / TOC）。"""
    run = paragraph.add_run()
    run.font.size = Pt(size)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = code
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instr, separate, text, end):
        run._r.append(element)
    _apply_cjk_font(run, theme)


# --------------------------------------------------------------------------
# 样式与页面
# --------------------------------------------------------------------------
def _prepare_styles(doc, theme: Theme) -> None:
    _style_font(doc.styles["Normal"], theme, theme.font_body, theme.body, bold=False, color=theme.text)

    for name, size, color in (
        ("Title", theme.h1 + 8, theme.primary),
        ("Heading 1", theme.h1, theme.primary),
        ("Heading 2", theme.h2, theme.primary),
        ("Heading 3", theme.h3, theme.text),
    ):
        style = doc.styles[name]
        _style_font(style, theme, theme.font_heading, size, bold=True, color=color)
        style.paragraph_format.space_before = Pt(14 if name != "Title" else 0)
        style.paragraph_format.space_after = Pt(8 if name != "Title" else 18)
        style.paragraph_format.keep_with_next = True

    caption = doc.styles["Caption"]
    _style_font(caption, theme, theme.font_body, theme.caption, bold=False, color=theme.muted)

    if "Table Grid" not in [s.name for s in doc.styles]:
        doc.styles.add_style("Table Grid", 3)


def _prepare_page(doc, theme: Theme, document: Document) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.6)
    section.right_margin = Cm(2.6)
    section.top_margin = Cm(2.4)
    section.bottom_margin = Cm(2.2)
    section.different_first_page_header_footer = True

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header_text = f"{theme.brand_name}  ·  {document.title}"
    if document.delivery is not None and document.delivery.client:
        header_text = f"{document.delivery.client}  ·  {document.title}  ·  交付件"
    run = header.add_run(header_text)
    run.font.size = Pt(theme.caption)
    run.font.color.rgb = RGBColor.from_string(theme.muted)
    _apply_cjk_font(run, theme)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("第 ")
    run.font.size = Pt(theme.caption)
    run.font.color.rgb = RGBColor.from_string(theme.muted)
    _apply_cjk_font(run, theme)
    _field(footer, "PAGE", "1", theme, theme.caption)
    run = footer.add_run(" 页 / 共 ")
    run.font.size = Pt(theme.caption)
    run.font.color.rgb = RGBColor.from_string(theme.muted)
    _apply_cjk_font(run, theme)
    _field(footer, "NUMPAGES", "1", theme, theme.caption)
    run = footer.add_run(f" 页   |   {theme.footer_text}")
    run.font.size = Pt(theme.caption)
    run.font.color.rgb = RGBColor.from_string(theme.muted)
    _apply_cjk_font(run, theme)


def _paragraph(doc, text: str, theme: Theme, size: Optional[float] = None,
               color: Optional[str] = None, bold: bool = False,
               align=None, space_after: float = 6.0, italic: bool = False):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(space_after)
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run(text)
    run.font.size = Pt(size if size is not None else theme.body)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    _apply_cjk_font(run, theme)
    return paragraph


# --------------------------------------------------------------------------
# 内容块
# --------------------------------------------------------------------------
def _add_bullets(doc, block: Bullets, theme: Theme) -> None:
    style = "List Number" if block.ordered else "List Bullet"
    for item in block.items:
        try:
            paragraph = doc.add_paragraph(style=style)
        except KeyError:
            paragraph = doc.add_paragraph()
            item = ("• " if not block.ordered else "") + item
        paragraph.paragraph_format.space_after = Pt(3)
        run = paragraph.add_run(item)
        run.font.size = Pt(theme.body)
        _apply_cjk_font(run, theme)


def _add_table(doc, block: Table) -> None:
    theme = _theme_of(doc)
    headers = block.headers or []
    rows = block.rows
    if not headers:
        return

    if block.caption:
        caption = doc.add_paragraph(style="Caption")
        run = caption.add_run(block.caption)
        run.font.size = Pt(theme.caption)
        run.font.color.rgb = RGBColor.from_string(theme.muted)
        _apply_cjk_font(run, theme)

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        _cell_text(cell, header, theme, theme.caption, bold=True, color="FFFFFF")
        _shade(cell, theme.primary)

    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for column, value in enumerate(list(row)[: len(headers)]):
            status_color = None
            if block.status_column == column:
                status_color = theme.status_color(str(value).strip().lower())
            align = WD_ALIGN_PARAGRAPH.RIGHT if column in block.numeric_columns else None
            _cell_text(
                cells[column],
                value,
                theme,
                theme.caption,
                bold=block.status_column == column,
                color=status_color,
                align=align,
            )
            if row_index % 2 == 1:
                _shade(cells[column], theme.surface)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_metrics(doc, block: Metrics, theme: Theme) -> None:
    if not block.items:
        return
    table = doc.add_table(rows=2, cols=len(block.items))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for index, metric in enumerate(block.items):
        value_cell = table.rows[0].cells[index]
        _cell_text(value_cell, metric.value, theme, theme.h2, bold=True,
                   color=theme.status_color(metric.status), align=WD_ALIGN_PARAGRAPH.CENTER)
        _shade(value_cell, theme.surface)

        label_cell = table.rows[1].cells[index]
        text = metric.label if not metric.note else f"{metric.label}\n{metric.note}"
        _cell_text(label_cell, text, theme, theme.caption, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_note(doc, block: Note, theme: Theme) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.rows[0].cells[0]
    label = _NOTE_LABEL.get(block.level, "说明")
    _cell_text(cell, f"{label}：{block.text}", theme, theme.caption,
               color=theme.status_color("danger" if block.level == "danger" else
                                        "warning" if block.level == "warning" else
                                        "success" if block.level == "success" else "muted"))
    _shade(cell, _NOTE_FILL.get(block.level, "EAF3FD"))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_code(doc, block: Code, theme: Theme) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(8)
    run = paragraph.add_run(block.text)
    run.font.size = Pt(theme.caption)
    run.font.color.rgb = RGBColor.from_string(theme.text)
    _apply_cjk_font(run, theme, mono=True)


def _add_chart(doc, block: Chart, theme: Theme) -> None:
    import io

    from docx.shared import Inches

    try:
        png = render_chart_png(block, theme)
    except Exception as error:
        _paragraph(doc, f"图表「{block.title}」渲染失败：{error}", theme, color=theme.danger)
        return
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(4)
    paragraph.add_run().add_picture(io.BytesIO(png), width=Inches(6.1))
    if block.title:
        caption = doc.add_paragraph(style="Caption")
        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = caption.add_run(f"图：{block.title}")
        run.font.size = Pt(theme.caption)
        run.font.color.rgb = RGBColor.from_string(theme.muted)
        _apply_cjk_font(run, theme)


# --------------------------------------------------------------------------
# 对外交付版式：客户抬头 / 条款页 / 签章页
# --------------------------------------------------------------------------
def _key_value_table(doc, rows: List[tuple], theme: Theme, label_width: float = 3.2) -> None:
    """两列键值表（左侧标签带底纹）。"""
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    table.columns[0].width = Cm(label_width)
    for index, (key, value) in enumerate(rows):
        _cell_text(table.rows[index].cells[0], key, theme, theme.caption, bold=True)
        _shade(table.rows[index].cells[0], theme.surface)
        _cell_text(table.rows[index].cells[1], value, theme, theme.caption)


def _delivery_cover_block(doc, delivery: Delivery, theme: Theme) -> None:
    """封面上的客户抬头：标明接收方与项目，避免交付件被误发。"""
    rows = [
        ("致", delivery.client),
        ("项目名称", delivery.project),
        ("客户方对接人", delivery.contact),
        ("交付方负责人", delivery.owner),
        ("生效日期", delivery.effective_date),
    ]
    rows = [(k, v) for k, v in rows if v]
    if rows:
        _key_value_table(doc, rows, theme)
        doc.add_paragraph().paragraph_format.space_after = Pt(10)


def _delivery_letterhead(doc, delivery: Delivery, theme: Theme) -> None:
    """交付说明页（抬头页）。"""
    heading = doc.add_heading("交付说明", level=1)
    for run in heading.runs:
        _apply_cjk_font(run, theme, theme.font_heading)

    rows = [
        ("接收方", delivery.client),
        ("项目名称", delivery.project),
        ("客户方对接人", delivery.contact),
        ("交付方负责人", delivery.owner),
        ("生效日期", delivery.effective_date),
        ("保密约定", delivery.confidential),
    ]
    rows = [(k, v) for k, v in rows if v]
    if rows:
        _key_value_table(doc, rows, theme)
        doc.add_paragraph().paragraph_format.space_after = Pt(8)


def _delivery_terms(doc, delivery: Delivery, theme: Theme) -> None:
    """条款与约定页。"""
    heading = doc.add_heading("条款与约定", level=1)
    for run in heading.runs:
        _apply_cjk_font(run, theme, theme.font_heading)

    if delivery.terms:
        _add_bullets(doc, Bullets(items=delivery.terms, ordered=True), theme)
    if delivery.confidential:
        _add_note(doc, Note(text=delivery.confidential, level="warning"), theme)
    if delivery.warranty:
        _add_note(doc, Note(text=delivery.warranty, level="info"), theme)


def _delivery_signature(doc, delivery: Delivery, theme: Theme) -> None:
    """签章页：角色 / 签字 / 日期。"""
    heading = doc.add_heading("签署页", level=1)
    for run in heading.runs:
        _apply_cjk_font(run, theme, theme.font_heading)

    _paragraph(doc, "本文件经双方确认后生效，签署即代表接受本文件全部条款与约定。",
               theme, space_after=10)

    roles = delivery.signatures or ["客户方负责人（签章）", "交付方负责人（签章）"]
    table = doc.add_table(rows=len(roles) + 1, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for column, title in enumerate(("签署角色", "签字 / 盖章", "日期")):
        _cell_text(table.rows[0].cells[column], title, theme, theme.caption,
                   bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
        _shade(table.rows[0].cells[column], theme.primary)

    for index, role in enumerate(roles, start=1):
        _cell_text(table.rows[index].cells[0], role, theme, theme.caption)
        # 签字与日期留白，供打印后手写
        for column in (1, 2):
            _cell_text(table.rows[index].cells[column], "", theme, theme.caption)
        table.rows[index].height = Cm(1.8)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)


_THEME_KEY = "_docgen_theme"


def _theme_of(doc) -> Theme:
    return getattr(doc, _THEME_KEY)


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def render_docx(document: Document, out_path: Path) -> Path:
    """渲染为 .docx 文件。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    theme = document.theme
    doc = DocxDocument()
    setattr(doc, _THEME_KEY, theme)

    _prepare_styles(doc, theme)
    _prepare_page(doc, theme, document)

    # ---- 封面 ----
    _paragraph(doc, theme.organization, theme, size=theme.h3, color=theme.accent,
               bold=True, space_after=2)
    _paragraph(doc, theme.brand_name, theme, size=theme.body, color=theme.muted, space_after=90)

    # 对外交付件：封面顶部带客户抬头
    if document.delivery is not None:
        _delivery_cover_block(doc, document.delivery, theme)

    title = doc.add_paragraph(style="Title")
    run = title.add_run(document.title)
    _apply_cjk_font(run, theme, theme.font_heading)
    run.font.size = Pt(theme.h1 + 8)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(theme.primary)

    if document.subtitle:
        _paragraph(doc, document.subtitle, theme, size=theme.h3, color=theme.muted, space_after=40)

    meta = [
        ("文档编号", document.doc_id),
        ("版本", document.version),
        ("密级", document.classification),
        ("编制", document.author),
        ("所属部门", document.department),
        ("生成时间", document.generated_at),
        ("数据来源", document.source_commit),
    ]
    meta = [(k, v) for k, v in meta if v]
    if meta:
        table = doc.add_table(rows=len(meta), cols=2)
        table.style = "Table Grid"
        for index, (key, value) in enumerate(meta):
            _cell_text(table.rows[index].cells[0], key, theme, theme.caption, bold=True)
            _shade(table.rows[index].cells[0], theme.surface)
            _cell_text(table.rows[index].cells[1], value, theme, theme.caption)

    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 交付说明（抬头页） ----
    if document.delivery is not None:
        _delivery_letterhead(doc, document.delivery, theme)
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 摘要 ----
    if document.summary:
        heading = doc.add_heading("摘要", level=1)
        for run in heading.runs:
            _apply_cjk_font(run, theme, theme.font_heading)
        _paragraph(doc, document.summary, theme, space_after=10)

    if document.highlights:
        heading = doc.add_heading("关键结论", level=1)
        for run in heading.runs:
            _apply_cjk_font(run, theme, theme.font_heading)
        _add_bullets(doc, Bullets(items=document.highlights), theme)

    # ---- 目录 ----
    heading = doc.add_heading("目录", level=1)
    for run in heading.runs:
        _apply_cjk_font(run, theme, theme.font_heading)
    toc_paragraph = doc.add_paragraph()
    _field(toc_paragraph, 'TOC \\o "1-3" \\h \\z \\u', "右键此处选择「更新域」以生成目录", theme, theme.body)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 正文 ----
    for section in document.sections:
        if section.title:
            heading = doc.add_heading(section.title, level=max(1, min(section.level, 3)))
            for run in heading.runs:
                _apply_cjk_font(run, theme, theme.font_heading)

        for block in section.blocks:
            if isinstance(block, Para):
                _paragraph(doc, block.text, theme)
            elif isinstance(block, Bullets):
                _add_bullets(doc, block, theme)
            elif isinstance(block, Table):
                _add_table(doc, block)
            elif isinstance(block, Metrics):
                _add_metrics(doc, block, theme)
            elif isinstance(block, Note):
                _add_note(doc, block, theme)
            elif isinstance(block, Code):
                _add_code(doc, block, theme)
            elif isinstance(block, Chart):
                _add_chart(doc, block, theme)
            elif isinstance(block, PageBreak):
                doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- 条款页 + 签章页（仅对外交付件） ----
    if document.delivery is not None:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        _delivery_terms(doc, document.delivery, theme)
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        _delivery_signature(doc, document.delivery, theme)

    doc.save(str(out_path))
    return out_path
