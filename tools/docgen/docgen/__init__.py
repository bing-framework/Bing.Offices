"""Bing.Offices 文档自动化工具包。

分层：
    collectors/  采集层 —— 从仓库原始产物抽取结构化数据
    model.py     中间模型 —— 全类型统一的 Document / Section / Block
    renderers/   渲染层 —— Markdown / DOCX / XLSX / PPTX
    reports.py   组装层 —— 把采集结果编排成具体文档
    cli.py       入口
"""

__version__ = "1.0.0"
