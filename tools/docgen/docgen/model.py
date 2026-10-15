"""统一文档中间模型。

设计要点：**采集器只产出 Document，渲染器只消费 Document**。
新增一种文档类型 = 新采集器 + 复用已有渲染器；
新增一种输出格式 = 新渲染器 + 复用已有采集器。
两者互不感知，这是「可复用模板」原则的结构保证。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from .theme import DEFAULT_THEME, Theme


# --------------------------------------------------------------------------
# 块
# --------------------------------------------------------------------------
@dataclass
class Block:
    """所有内容块的基类。"""


@dataclass
class Para(Block):
    """普通段落。"""

    text: str = ""


@dataclass
class Bullets(Block):
    """列表。"""

    items: List[str] = field(default_factory=list)
    ordered: bool = False


@dataclass
class Table(Block):
    """表格。"""

    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    caption: str = ""
    # 状态列索引：该列单元格按 success/danger/warning 着色
    status_column: Optional[int] = None
    # 数值列索引：右对齐
    numeric_columns: Sequence[int] = ()


@dataclass
class Metric(Block):
    """单个指标卡。"""

    label: str = ""
    value: str = ""
    note: str = ""
    status: str = "neutral"  # neutral | success | danger | warning


@dataclass
class Metrics(Block):
    """指标卡组（横向排列）。"""

    items: List[Metric] = field(default_factory=list)


@dataclass
class Chart(Block):
    """图表。

    kind: bar | line | pie
    """

    title: str = ""
    kind: str = "bar"
    categories: List[str] = field(default_factory=list)
    series: List[Tuple[str, List[float]]] = field(default_factory=list)
    unit: str = ""
    horizontal: bool = False


@dataclass
class Note(Block):
    """提示块。level: info | warning | danger | success"""

    text: str = ""
    level: str = "info"


@dataclass
class Code(Block):
    """代码块。"""

    text: str = ""
    language: str = ""


@dataclass
class PageBreak(Block):
    """分页符。"""


# --------------------------------------------------------------------------
# 对外交付元信息
# --------------------------------------------------------------------------
@dataclass
class Delivery:
    """对外交付文档专属元信息。

    只有「交付类」文档会带上它；渲染器据此追加客户抬头、条款页与签章页。
    技术报告、API 参考等内部文档保持 delivery=None，不会多出这些版式。
    """

    client: str = ""            # 客户 / 接收方名称
    project: str = ""           # 项目或标的名称
    contact: str = ""           # 客户方对接人
    owner: str = ""             # 我方交付负责人
    effective_date: str = ""    # 生效日期
    confidential: str = ""      # 保密约定
    warranty: str = ""          # 质保与支持约定
    terms: List[str] = field(default_factory=list)        # 条款条目
    # 签章角色，例如 ["客户方负责人（签章）", "交付方负责人（签章）"]
    signatures: List[str] = field(default_factory=list)

    def is_configured(self) -> bool:
        return bool(self.client or self.project or self.terms or self.signatures)


# --------------------------------------------------------------------------
# 文档
# --------------------------------------------------------------------------
@dataclass
class Section:
    """章节。level=1 为一级标题。"""

    title: str = ""
    level: int = 1
    blocks: List[Block] = field(default_factory=list)

    def add(self, *blocks: Block) -> "Section":
        self.blocks.extend(b for b in blocks if b is not None)
        return self


@dataclass
class Document:
    """一份完整文档。"""

    title: str = ""
    subtitle: str = ""
    doc_id: str = ""
    version: str = ""
    author: str = ""
    department: str = ""
    generated_at: str = ""
    source_commit: str = ""
    classification: str = "内部资料"

    summary: str = ""
    highlights: List[str] = field(default_factory=list)

    sections: List[Section] = field(default_factory=list)

    # 对外交付版式；None 表示普通技术文档
    delivery: Optional[Delivery] = None

    theme: Theme = field(default_factory=lambda: DEFAULT_THEME)

    # 输出文件名（不含扩展名）
    filename: str = "document"

    # ------------------------------------------------------------------
    def add_section(self, title: str, level: int = 1) -> Section:
        section = Section(title=title, level=level)
        self.sections.append(section)
        return section

    def all_tables(self) -> List[Table]:
        return [b for s in self.sections for b in s.blocks if isinstance(b, Table)]

    def all_charts(self) -> List[Chart]:
        return [b for s in self.sections for b in s.blocks if isinstance(b, Chart)]

    def heading_count(self) -> int:
        return sum(1 for s in self.sections if s.title)

    def word_count(self) -> int:
        """粗略统计文本量，用于生成物自检。"""
        total = len(self.title) + len(self.summary)
        total += sum(len(h) for h in self.highlights)
        if self.delivery is not None:
            total += sum(len(t) for t in self.delivery.terms)
            total += sum(len(s) for s in self.delivery.signatures)
            total += len(self.delivery.client) + len(self.delivery.project)
        for section in self.sections:
            total += len(section.title)
            for block in section.blocks:
                if isinstance(block, Para):
                    total += len(block.text)
                elif isinstance(block, Bullets):
                    total += sum(len(i) for i in block.items)
                elif isinstance(block, Table):
                    total += len(block.caption)
                    total += sum(len(str(c)) for row in block.rows for c in row)
                    total += sum(len(h) for h in block.headers)
                elif isinstance(block, Note):
                    total += len(block.text)
                elif isinstance(block, Code):
                    total += len(block.text)
                elif isinstance(block, Metrics):
                    total += sum(len(m.label) + len(m.value) + len(m.note) for m in block.items)
        return total

    def validate(self) -> List[str]:
        """生成前自检，返回问题列表（空表示通过）。"""
        problems: List[str] = []
        if not self.title:
            problems.append("文档缺少 title")
        if not self.sections:
            problems.append("文档没有任何章节")
        for index, table in enumerate(self.all_tables(), start=1):
            if table.headers and table.rows:
                width = len(table.headers)
                for r, row in enumerate(table.rows, start=1):
                    if len(row) != width:
                        problems.append(
                            f"第 {index} 张表第 {r} 行有 {len(row)} 列，表头为 {width} 列"
                        )
        for index, chart in enumerate(self.all_charts(), start=1):
            for name, values in chart.series:
                if len(values) != len(chart.categories):
                    problems.append(
                        f"第 {index} 张图系列「{name}」有 {len(values)} 个值，"
                        f"但类别数为 {len(chart.categories)}"
                    )
        for section in self.sections:
            if section.level not in (1, 2, 3):
                problems.append(f"章节「{section.title}」层级 {section.level} 非法")
        if self.delivery is not None and not self.delivery.is_configured():
            problems.append("文档声明为对外交付件，但 delivery 元信息为空")
        return problems
