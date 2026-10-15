"""命令行入口。

用法::

    python -m docgen all                       # 生成全部文档
    python -m docgen tests --format docx,xlsx  # 只生成测试与基准报告
    python -m docgen release --format md
    python -m docgen deck  --format pptx
    python -m docgen api   --format md
    python -m docgen delivery --format docx,pdf   # 对外交付文档
    python -m docgen all --check               # 仅自检，不写文件（CI 用）
    python -m docgen gate                      # 质量门禁（注释覆盖率等）
    python -m docgen docs-check                # 文档腐化防护校验

产物位置：
    Markdown              -> docs/generated/
    DOCX / XLSX / PPTX/PDF -> artifacts/docs/
"""

from __future__ import annotations

import argparse
import sys
import traceback
from pathlib import Path
from typing import Dict, List

from .collectors import collect_api
from .gates import GateConfig, run_gates
from .model import Document
from .renderers import RENDERERS
from .repo import default_layout
from .reports import BUILDERS
from .samples import check_samples

TARGETS = ("tests", "release", "deck", "api", "delivery")
# 元目标：不产出文档，而是执行治理动作
META_TARGETS = ("gate", "docs-check")
ALL_TARGETS = (*TARGETS, *META_TARGETS)
FORMATS = ("md", "docx", "xlsx", "pptx", "pdf")
MARKDOWN_FORMATS = {"md"}


def _parse_formats(raw: str) -> List[str]:
    items = [item.strip().lower() for item in raw.split(",") if item.strip()]
    if "all" in items:
        return list(FORMATS)
    invalid = [item for item in items if item not in FORMATS]
    if invalid:
        raise argparse.ArgumentTypeError(
            f"不支持的格式：{', '.join(invalid)}；可选 {', '.join(FORMATS)}"
        )
    return items


def _output_path(document: Document, fmt: str, layout) -> Path:
    directory = layout.markdown_out if fmt in MARKDOWN_FORMATS else layout.binary_out
    return directory / f"{document.filename}.{fmt}"


def build_target(name: str, layout, check: bool = False) -> Document:
    builder = BUILDERS[name]
    # tests 报告在生成时会归档基准快照；--check 只读模式关闭归档副作用
    document = builder(layout, archive=not check) if name == "tests" else builder(layout)
    # 统一补齐生成时间，避免各 builder 遗漏
    if not document.generated_at:
        document.generated_at = _now()
    return document


def _now() -> str:
    import os

    if os.environ.get("DOCGEN_STABLE") == "1":
        return "由数据源确定"
    from datetime import datetime

    return datetime.now().strftime("%Y-%m-%d %H:%M")


# --------------------------------------------------------------------------
# 元目标：质量门禁 / 文档腐化校验
# --------------------------------------------------------------------------
def run_gate(args: argparse.Namespace) -> int:
    """执行质量门禁；任一门禁未通过即返回 1（CI 会因此失败）。"""
    layout = default_layout()
    config = GateConfig.load(layout)
    if args.min is not None:
        config.min_coverage = args.min

    print("=== 质量门禁 ===")
    results, coverage = run_gates(layout, config)
    for result in results:
        print(f"  {result.render()}")
        for problem in result.problems:
            print(f"       - {problem}")

    if coverage is not None and coverage.available:
        print(f"\n  分程序集明细（{coverage.tfm}，口径不含无参构造函数）：")
        for assembly in coverage.assemblies:
            print(f"      {assembly.name:<28} {assembly.pct:6.2f}%  "
                  f"{assembly.documented:>4}/{assembly.total:<4}  "
                  f"构造注释率 {assembly.ctor_pct:5.1f}%")

    failed = [r for r in results if not r.passed]
    print(f"\n  门禁 {len(results) - len(failed)} 项通过，{len(failed)} 项未通过")
    return 1 if failed else 0


def run_docs_check(args: argparse.Namespace) -> int:
    """文档腐化防护：校验生成文档中的示例与可编译源码一致。"""
    layout = default_layout()
    api = collect_api(layout.release_output_dir, layout.api_baseline)
    namespaces = list(api.by_namespace().keys())

    print("=== 文档腐化校验 ===")
    print(f"  示例宿主项目  : {layout.docs_tests_dir.name}")
    print(f"  扫描生成文档  : {layout.markdown_out.name}")
    result = check_samples(layout, required_namespaces=namespaces)

    print(f"  源码示例      : {len(result.source_samples)} 个")
    for sample in result.source_samples:
        print(f"      - {sample.name:<20} {sample.title}  ({sample.source}:{sample.line})")
    print(f"  文档内嵌示例  : {result.embedded_count} 处，"
          f"分布在 {len(result.doc_samples)} 个文件")

    if result.problems:
        print("  [未通过] 发现问题：")
        for problem in result.problems:
            print(f"       - {problem}")
        return 1
    print("  [通过] 文档示例与可编译源码完全一致")
    return 0


def run(args: argparse.Namespace) -> int:
    # 元目标互不兼容地先行处理，不参与文档生成循环
    meta = [t for t in args.target if t in META_TARGETS]
    if len(meta) == 1:
        if meta[0] == "gate":
            return run_gate(args)
        return run_docs_check(args)
    if meta:
        print("[错误] gate 与 docs-check 不能同时执行")
        return 2

    layout = default_layout()

    missing = layout.missing_inputs()
    if missing:
        print("[警告] 以下输入缺失，相关章节将降级：")
        for item in missing:
            print(f"        - {item}")

    targets: List[str] = list(TARGETS) if "all" in args.target else list(args.target)
    formats: List[str] = _parse_formats(args.format)

    if not args.check:
        layout.ensure_outputs()

    failures = 0
    produced: List[Path] = []

    for target in targets:
        print(f"\n=== 生成：{target} ===")
        try:
            document = build_target(target, layout, check=args.check)
        except Exception as error:
            failures += 1
            print(f"  [失败] 组装文档时出错：{error}")
            traceback.print_exc()
            continue

        problems = document.validate()
        if problems:
            failures += 1
            print("  [失败] 文档自检未通过：")
            for problem in problems:
                print(f"         - {problem}")
            continue

        print(f"  标题       : {document.title}")
        print(f"  章节数     : {len(document.sections)}")
        print(f"  表格 / 图表: {len(document.all_tables())} / {len(document.all_charts())}")
        print(f"  文本量     : 约 {document.word_count()} 字符")

        if args.check:
            print("  [自检通过] --check 模式，未写出文件")
            continue

        for fmt in formats:
            renderer = RENDERERS[fmt]
            path = _output_path(document, fmt, layout)
            try:
                renderer(document, path)
            except Exception as error:
                failures += 1
                print(f"  [失败] {fmt} 渲染出错：{error}")
                traceback.print_exc()
                continue
            size_kb = path.stat().st_size / 1024
            produced.append(path)
            print(f"  [完成] {fmt:<5} -> {path.relative_to(layout.root)}  ({size_kb:.1f} KB)")

    print("\n=== 汇总 ===")
    if produced:
        for path in produced:
            print(f"  产出 {path.relative_to(layout.root)}")
    print(f"  成功 {len(produced)} 个文件，失败 {failures} 项")
    return 1 if failures else 0


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docgen",
        description="Bing.Offices 文档自动化流水线",
    )
    parser.add_argument(
        "target",
        nargs="+",
        choices=(*ALL_TARGETS, "all"),
        help="要生成的文档类型：tests / release / deck / api / delivery / all；"
             "治理动作：gate（质量门禁） / docs-check（文档腐化校验）",
    )
    parser.add_argument(
        "-f", "--format",
        default="all",
        help="输出格式，逗号分隔：md,docx,xlsx,pptx,pdf,all（默认 all）",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="仅做数据采集与文档自检，不写出文件",
    )
    parser.add_argument(
        "--min",
        type=float,
        default=None,
        help="gate 目标下覆盖配置文件中的注释覆盖率总阈值（百分比）",
    )
    args = parser.parse_args(argv)

    try:
        return run(args)
    except Exception as error:  # 顶层兜底，避免 CI 里只看到半截堆栈
        print(f"[致命错误] {error}", file=sys.stderr)
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
