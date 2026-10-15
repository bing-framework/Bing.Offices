"""渲染层：把 Document 渲染为具体文件格式。"""

from .docx import render_docx
from .markdown import render_markdown
from .pdf import render_pdf
from .pptx import render_pptx
from .xlsx import render_xlsx

# 格式名 -> 渲染函数
RENDERERS = {
    "md": render_markdown,
    "docx": render_docx,
    "xlsx": render_xlsx,
    "pptx": render_pptx,
    "pdf": render_pdf,
}

__all__ = [
    "RENDERERS",
    "render_markdown",
    "render_docx",
    "render_xlsx",
    "render_pptx",
    "render_pdf",
]
