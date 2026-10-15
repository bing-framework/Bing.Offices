"""XLSX 渲染器（openpyxl）。

设计要点：
  * 用 **Office 原生图表**，使用者可以在 Excel 里直接改数据、换图型；
  * 表头冻结 + 自动筛选 + 状态列条件着色，让导出即「可用于分析的台账」；
  * 表格与图表各自独立 Sheet，避免相互干扰。
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
    Table,
)
from ..theme import Theme

_INVALID_SHEET = re.compile(r"[\\/*?:\[\]]")
_STATUS_FILL = {"success": "EAF6EC", "danger": "FDEFF2", "warning": "FFF8E5", "neutral": "F2F5F9"}


def _sheet_name(raw: str, used: set) -> str:
    """Excel 工作表名约束：≤31 字符、不含 \\/*?:[]、不重名。"""
    name = _INVALID_SHEET.sub("-", (raw or "Sheet").strip()) or "Sheet"
    name = name[:31]
    candidate, index = name, 2
    while candidate.lower() in used:
        suffix = f"-{index}"
        candidate = name[: 31 - len(suffix)] + suffix
        index += 1
    used.add(candidate.lower())
    return candidate


def _display_width(text: str) -> int:
    """中日韩字符按 2 个宽度计。"""
    width = 0
    for char in str(text):
        width += 2 if ord(char) > 0x2E80 else 1
    return width


def _fit_columns(worksheet, headers: List[str], rows: List[List[str]], start_column: int = 1,
                 minimum: int = 8, maximum: int = 46) -> None:
    for offset, header in enumerate(headers):
        column = start_column + offset
        widest = _display_width(header)
        for row in rows:
            if offset < len(row):
                widest = max(widest, _display_width(row[offset]))
        letter = get_column_letter(column)
        worksheet.column_dimensions[letter].width = max(minimum, min(maximum, widest + 3))


def _write_table(worksheet, table: Table, theme: Theme, start_row: int = 1,
                 start_column: int = 1) -> int:
    """写入一张表，返回下一个可用行号。"""
    headers = table.headers or []
    if not headers:
        return start_row

    thin = Side(style="thin", color=theme.border)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_font = Font(name=theme.font_heading, size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor=theme.primary)
    body_font = Font(name=theme.font_body, size=10, color=theme.text)

    row_cursor = start_row
    if table.caption:
        cell = worksheet.cell(row=row_cursor, column=start_column, value=table.caption)
        cell.font = Font(name=theme.font_heading, size=10, bold=True, color=theme.primary)
        row_cursor += 1

    for offset, header in enumerate(headers):
        cell = worksheet.cell(row=row_cursor, column=start_column + offset, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_row = row_cursor
    row_cursor += 1

    for row_index, row in enumerate(table.rows):
        for offset in range(len(headers)):
            value = row[offset] if offset < len(row) else ""
            cell = worksheet.cell(row=row_cursor, column=start_column + offset, value=value)
            cell.font = body_font
            cell.border = border
            if offset in table.numeric_columns:
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(vertical="top", wrap_text=False)
            if table.status_column == offset:
                fill = _STATUS_FILL.get(str(value).strip().lower(), _STATUS_FILL["neutral"])
                cell.fill = PatternFill("solid", fgColor=fill)
                cell.font = Font(name=theme.font_body, size=10, bold=True, color=theme.text)
            elif row_index % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=theme.surface)
        row_cursor += 1

    _fit_columns(worksheet, headers, table.rows, start_column)
    worksheet.freeze_panes = worksheet.cell(row=header_row + 1, column=start_column)
    worksheet.auto_filter.ref = (
        f"{get_column_letter(start_column)}{header_row}:"
        f"{get_column_letter(start_column + len(headers) - 1)}{row_cursor - 1}"
    )
    return row_cursor + 2


def _add_native_chart(worksheet, chart: Chart, theme: Theme, data_row: int,
                      data_column: int, anchor: str) -> int:
    """把 Chart 的数据写入工作表并创建 Office 原生图表，返回下一个可用行号。"""
    if not chart.categories or not chart.series:
        return data_row

    # 数据区：第 1 列类别 + 各系列一列
    cell = worksheet.cell(row=data_row, column=data_column, value="类别")
    cell.font = Font(name=theme.font_heading, size=10, bold=True, color=theme.primary)
    for series_index, (name, _) in enumerate(chart.series):
        cell = worksheet.cell(row=data_row, column=data_column + 1 + series_index, value=name)
        cell.font = Font(name=theme.font_heading, size=10, bold=True, color=theme.primary)

    for row_index, category in enumerate(chart.categories):
        worksheet.cell(row=data_row + 1 + row_index, column=data_column, value=category)
        for series_index, (_, values) in enumerate(chart.series):
            value = values[row_index] if row_index < len(values) else None
            worksheet.cell(row=data_row + 1 + row_index, column=data_column + 1 + series_index,
                           value=value)

    first_data_row = data_row + 1
    last_data_row = data_row + len(chart.categories)
    last_column = data_column + len(chart.series)

    if chart.kind == "pie":
        native = PieChart()
    elif chart.kind == "line":
        native = LineChart()
    else:
        native = BarChart()
        native.type = "bar" if chart.horizontal else "col"

    native.title = chart.title
    native.style = 10
    native.height = 8.5
    native.width = 17

    data = Reference(worksheet, min_col=data_column + 1, max_col=last_column,
                     min_row=data_row, max_row=last_data_row)
    categories = Reference(worksheet, min_col=data_column, min_row=first_data_row,
                           max_row=last_data_row)
    native.add_data(data, titles_from_data=True)
    native.set_categories(categories)

    if chart.unit and chart.kind != "pie":
        native.y_axis.title = chart.unit
    if chart.kind == "pie":
        native.dataLabels = DataLabelList()
        native.dataLabels.showPercent = True

    for index in range(len(chart.series)):
        color = theme.series_color(index)
        series = native.series[index]
        series.graphicalProperties.solidFill = color
        if chart.kind == "line":
            series.graphicalProperties.line.solidFill = color

    worksheet.add_chart(native, anchor)
    # 数据区宽：1 列类别 + N 列系列
    consumed = max(len(chart.categories) + 2, 12)
    return data_row + consumed


def render_xlsx(document: Document, out_path: Path) -> Path:
    """渲染为 .xlsx 文件。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    theme = document.theme
    workbook = Workbook()
    used: set = set()

    # ---- 概览 ----
    overview = workbook.active
    overview.title = _sheet_name("概览", used)
    overview.sheet_view.showGridLines = False
    overview.column_dimensions["A"].width = 22
    overview.column_dimensions["B"].width = 60

    row = 1
    overview.cell(row=row, column=1, value=document.title).font = Font(
        name=theme.font_heading, size=14, bold=True, color=theme.primary)
    row += 2
    for key, value in (
        ("文档编号", document.doc_id),
        ("副标题", document.subtitle),
        ("版本", document.version),
        ("密级", document.classification),
        ("编制", document.author),
        ("生成时间", document.generated_at),
        ("数据来源", document.source_commit),
    ):
        if not value:
            continue
        cell = overview.cell(row=row, column=1, value=key)
        cell.font = Font(name=theme.font_heading, size=10, bold=True, color=theme.muted)
        cell.fill = PatternFill("solid", fgColor=theme.surface)
        value_cell = overview.cell(row=row, column=2, value=value)
        value_cell.font = Font(name=theme.font_body, size=10, color=theme.text)
        value_cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    if document.summary:
        row += 1
        overview.cell(row=row, column=1, value="摘要").font = Font(
            name=theme.font_heading, size=11, bold=True, color=theme.primary)
        row += 1
        cell = overview.cell(row=row, column=1, value=document.summary)
        cell.font = Font(name=theme.font_body, size=10, color=theme.text)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        overview.merge_cells(start_row=row, start_column=1, end_row=row + 5, end_column=2)
        row += 7

    if document.highlights:
        overview.cell(row=row, column=1, value="关键结论").font = Font(
            name=theme.font_heading, size=11, bold=True, color=theme.primary)
        row += 1
        for item in document.highlights:
            cell = overview.cell(row=row, column=1, value=f"• {item}")
            cell.font = Font(name=theme.font_body, size=10, color=theme.text)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            overview.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 1

    # ---- 指标 ----
    metrics = [b for s in document.sections for b in s.blocks if isinstance(b, Metrics)]
    if metrics:
        row += 1
        overview.cell(row=row, column=1, value="核心指标").font = Font(
            name=theme.font_heading, size=11, bold=True, color=theme.primary)
        row += 1
        for key in ("指标", "数值", "说明"):
            cell = overview.cell(row=row, column=1 if key == "指标" else 2 if key == "数值" else 3,
                                 value=key)
            cell.font = Font(name=theme.font_heading, size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=theme.primary)
        overview.column_dimensions["C"].width = 34
        row += 1
        for block in metrics:
            for metric in block.items:
                overview.cell(row=row, column=1, value=metric.label).font = Font(
                    name=theme.font_body, size=10, color=theme.text)
                value_cell = overview.cell(row=row, column=2, value=metric.value)
                value_cell.font = Font(name=theme.font_body, size=10, bold=True,
                                       color=theme.status_color(metric.status))
                overview.cell(row=row, column=3, value=metric.note).font = Font(
                    name=theme.font_body, size=10, color=theme.muted)
                row += 1

    # ---- 每张表一个 Sheet ----
    for index, table in enumerate(document.all_tables(), start=1):
        title = table.caption or f"表 {index}"
        worksheet = workbook.create_sheet(_sheet_name(title, used))
        worksheet.sheet_view.showGridLines = False
        _write_table(worksheet, table, theme)

    # ---- 图表 Sheet（原生图表） ----
    charts = document.all_charts()
    if charts:
        worksheet = workbook.create_sheet(_sheet_name("图表", used))
        worksheet.sheet_view.showGridLines = False
        cursor = 1
        for chart in charts:
            cursor = _add_native_chart(worksheet, chart, theme, cursor, 1, f"K{cursor}")

    workbook.save(str(out_path))
    return out_path
