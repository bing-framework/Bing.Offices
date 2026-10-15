"""Markdown 渲染器。

适合：工程内部技术文档、发布说明、API 参考 —— 可入库、可 diff、可被
``Bing.Offices.Docs.Tests`` 之类的文档测试消费。

图表以 PNG 落到同目录 ``assets/<文件名>/`` 下并相对引用。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from ..charts import render_chart_png
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

_NOTE_LABEL = {
    "info": "说明",
    "warning": "注意",
    "danger": "警告",
    "success": "结论",
}


def _escape_cell(value: str) -> str:
    """转义表格单元格中的竖线与换行。"""
    return str(value).replace("|", "\\|").replace("\n", "<br>")


def _render_table(table: Table) -> List[str]:
    lines: List[str] = []
    if table.caption:
        lines.append(f"**{table.caption}**")
        lines.append("")
    headers = table.headers or []
    if not headers:
        return lines
    lines.append("| " + " | ".join(_escape_cell(h) for h in headers) + " |")
    lines.append("|" + "|".join(" --- " for _ in headers) + "|")
    for row in table.rows:
        cells = list(row) + [""] * (len(headers) - len(row))
        lines.append("| " + " | ".join(_escape_cell(c) for c in cells[: len(headers)]) + " |")
    lines.append("")
    return lines


def render_markdown(document: Document, out_path: Path) -> Path:
    """渲染为 Markdown 文件。"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    assets_dir = out_path.parent / "assets" / out_path.stem
    lines: List[str] = []

    # ---- 头部 ----
    lines.append(f"# {document.title}")
    lines.append("")
    if document.subtitle:
        lines.append(f"> {document.subtitle}")
        lines.append("")

    meta_pairs = [
        ("文档编号", document.doc_id),
        ("版本", document.version),
        ("生成时间", document.generated_at),
        ("数据来源", document.source_commit),
        ("编制", document.author),
        ("密级", document.classification),
    ]
    meta_pairs = [(k, v) for k, v in meta_pairs if v]
    if meta_pairs:
        lines.append("| 项 | 值 |")
        lines.append("| --- | --- |")
        for key, value in meta_pairs:
            lines.append(f"| {key} | {value} |")
        lines.append("")

    if document.delivery is not None:
        delivery = document.delivery
        pairs = [
            ("致（接收方）", delivery.client),
            ("项目名称", delivery.project),
            ("客户方对接人", delivery.contact),
            ("交付方负责人", delivery.owner),
            ("生效日期", delivery.effective_date),
            ("保密约定", delivery.confidential),
        ]
        pairs = [(k, v) for k, v in pairs if v]
        if pairs:
            lines.append("## 交付抬头")
            lines.append("")
            lines.append("| 项 | 值 |")
            lines.append("| --- | --- |")
            for key, value in pairs:
                lines.append(f"| {key} | {value} |")
            lines.append("")

    if document.summary:
        lines.append("## 摘要")
        lines.append("")
        lines.append(document.summary)
        lines.append("")

    if document.highlights:
        lines.append("## 关键结论")
        lines.append("")
        for item in document.highlights:
            lines.append(f"- {item}")
        lines.append("")

    # ---- 目录 ----
    if document.heading_count() > 2:
        lines.append("## 目录")
        lines.append("")
        for section in document.sections:
            if section.level == 1 and section.title:
                anchor = section.title.strip().lower().replace(" ", "-")
                lines.append(f"- [{section.title}](#{anchor})")
        lines.append("")

    # ---- 正文 ----
    chart_index = 0
    for section in document.sections:
        level = max(1, min(section.level, 6))
        if section.title:
            lines.append("#" * (level + 1) + f" {section.title}")
            lines.append("")

        for block in section.blocks:
            lines.extend(_render_block(block, document))

            if isinstance(block, Chart):
                chart_index += 1
                try:
                    png = render_chart_png(block, document.theme)
                    assets_dir.mkdir(parents=True, exist_ok=True)
                    image_path = assets_dir / f"chart-{chart_index}.png"
                    image_path.write_bytes(png)
                    relative = image_path.relative_to(out_path.parent).as_posix()
                    lines.append(f"![{block.title}]({relative})")
                    lines.append("")
                except Exception as error:  # 图表失败不应中断整篇文档
                    lines.append(f"> 图表「{block.title}」渲染失败：{error}")
                    lines.append("")

    # ---- 条款页 + 签署页（仅对外交付件） ----
    if document.delivery is not None:
        delivery = document.delivery
        lines.append("---")
        lines.append("")
        lines.append("## 条款与约定")
        lines.append("")
        for index, term in enumerate(delivery.terms, 1):
            lines.append(f"{index}. {term}")
        if delivery.confidential:
            lines.append("")
            lines.append(f"> **注意**：{delivery.confidential}")
        if delivery.warranty:
            lines.append("")
            lines.append(f"> **说明**：{delivery.warranty}")
        lines.append("")
        lines.append("## 签署页")
        lines.append("")
        lines.append("本文件经双方确认后生效，签署即代表接受本文件全部条款与约定。")
        lines.append("")
        roles = delivery.signatures or ["客户方负责人（签章）", "交付方负责人（签章）"]
        lines.append("| 签署角色 | 签字 / 盖章 | 日期 |")
        lines.append("| --- | --- | --- |")
        for role in roles:
            lines.append(f"| {role} |  |  |")
        lines.append("")

    lines.append("---")
    lines.append("")
    lines.append(f"*本文件由文档自动化流水线生成，请勿手工编辑。* {document.theme.footer_text}")
    lines.append("")

    text = "\n".join(lines)
    out_path.write_text(text, encoding="utf-8")
    return out_path


def _render_block(block: Block, document: Document) -> List[str]:
    if isinstance(block, Para):
        return [block.text, ""]

    if isinstance(block, Bullets):
        return [f"{i}. {item}" if block.ordered else f"- {item}" for i, item in enumerate(block.items, 1)] + [""]

    if isinstance(block, Table):
        return _render_table(block)

    if isinstance(block, Metrics):
        lines = ["| " + " | ".join(m.label for m in block.items) + " |",
                 "|" + "|".join(" --- " for _ in block.items) + "|",
                 "| " + " | ".join(f"**{m.value}**" for m in block.items) + " |"]
        notes = [m.note for m in block.items]
        if any(notes):
            lines.append("| " + " | ".join(n or "—" for n in notes) + " |")
        lines.append("")
        return lines

    if isinstance(block, Note):
        label = _NOTE_LABEL.get(block.level, "说明")
        return [f"> **{label}**：{block.text}", ""]

    if isinstance(block, Code):
        return [f"```{block.language}".rstrip(), block.text, "```", ""]

    if isinstance(block, PageBreak):
        return ["---", ""]

    if isinstance(block, Chart):
        # 图表标题的文本说明（图片在调用方追加）
        return [f"**图：{block.title}**", ""] if block.title else []

    return []
