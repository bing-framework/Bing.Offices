"""组装层：把采集结果编排成具体文档。

每种文档类型只关心「取哪些数据、怎么组织章节」，
输出格式由渲染器决定 —— 因此同一份报告可以同时产出 DOCX / XLSX / Markdown / PPTX。
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from .collectors import (
    ApiInventory,
    BenchmarkSet,
    ChangeLog,
    TestRun,
    collect_api,
    collect_benchmarks,
    collect_changelog,
    collect_coverage,
    collect_test_runs,
)
from .collectors.archives import (
    compare_with_previous,
    snapshot_from_sets,
    write_snapshots,
)
from .collectors.benchmark import format_bytes, format_ms
from .collectors.changelog import CHANGE_TYPES
from .collectors.tests import format_duration, summarize
from .model import (
    Bullets,
    Chart,
    Code,
    Delivery,
    Document,
    Metric,
    Metrics,
    Note,
    Para,
    Section,
    Table,
)
from .repo import RepoLayout, default_layout
from .samples import collect_source_samples
from .theme import DEFAULT_THEME, Theme

GENERATED_NOTE = "本文件由文档自动化流水线生成，请勿手工编辑。"

# 稳定模式：CI 漂移检查时置 1，使 generated_at 输出确定值，
# 从而 `git diff` 只反映「数据/结构变化」，不因时间戳抖动而误报。
STABLE_MODE = __import__("os").environ.get("DOCGEN_STABLE") == "1"


def _now() -> str:
    if STABLE_MODE:
        return "由数据源确定"
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _status_of(success: bool) -> str:
    return "success" if success else "danger"


# ==========================================================================
# 1. 测试与基准报告
# ==========================================================================
def build_test_benchmark_report(layout: Optional[RepoLayout] = None,
                                theme: Theme = DEFAULT_THEME,
                                archive: bool = True) -> Document:
    layout = layout or default_layout()
    runs = collect_test_runs(layout.test_results_dir)
    benches = collect_benchmarks(layout.benchmark_dir)
    api = collect_api(layout.release_output_dir, layout.api_baseline)
    changelog = collect_changelog(layout.root, limit=1)

    stats = summarize(runs)
    total_bench_methods = sum(len(b.rows) for b in benches)

    document = Document(
        title="测试与基准测试报告",
        subtitle=f"{theme.brand_name} · 质量与性能证据汇总",
        doc_id="BO-QR-001",
        version=api.package_version or "—",
        author="文档自动化流水线",
        department=theme.organization,
        generated_at=_now(),
        source_commit=changelog.head_sha[:12] if changelog.head_sha else "",
        classification="内部资料",
        filename="bing-offices-test-benchmark-report",
        theme=theme,
    )

    document.summary = (
        f"本次汇总覆盖 {stats['runs']} 个测试运行文件、{stats['total']} 个测试用例，"
        f"执行 {stats['executed']} 个，通过 {stats['passed']} 个，失败 {stats['failed']} 个，"
        f"整体通过率 {stats['pass_rate']:.2f}%。"
        f"累计耗时 {format_duration(stats['duration_s'])}。"
        f"同时收录 {len(benches)} 组基准测试、{total_bench_methods} 条性能测量结果。"
    )

    document.highlights = [
        f"测试通过率 {stats['pass_rate']:.2f}%（{stats['passed']}/{stats['executed']}）",
        f"失败用例 {stats['failed']} 个" + ("，需要优先处理" if stats["failed"] else "，无阻塞问题"),
        f"基准测试覆盖 {len(benches)} 个场景类、{total_bench_methods} 个方法",
    ]
    if stats["failed"]:
        document.highlights.append(
            "存在未通过用例：详情见「失败用例明细」章节"
        )

    # ---- 1 测试执行概览 ----
    section = document.add_section("测试执行概览")
    section.add(Metrics(items=[
        Metric(label="测试用例总数", value=str(stats["total"])),
        Metric(label="通过", value=str(stats["passed"]), status="success"),
        Metric(label="失败", value=str(stats["failed"]),
               status="danger" if stats["failed"] else "success"),
        Metric(label="通过率", value=f"{stats['pass_rate']:.2f}%",
               status=_status_of(stats["failed"] == 0)),
        Metric(label="累计耗时", value=format_duration(stats["duration_s"])),
    ]))

    rows = []
    for run in sorted(runs, key=lambda r: (r.suite, r.target_framework)):
        rows.append([
            run.suite,
            run.target_framework or "—",
            str(run.total),
            str(run.passed),
            str(run.failed),
            str(run.skipped),
            f"{run.pass_rate:.2f}%",
            format_duration(run.duration_s),
            "通过" if run.is_success else "未通过",
        ])
    section.add(Table(
        caption="表 1  各测试运行结果明细",
        headers=["测试套件", "目标框架", "用例数", "通过", "失败", "跳过", "通过率", "耗时", "结论"],
        rows=rows,
        status_column=8,
        numeric_columns=(2, 3, 4, 5, 6, 7),
    ))

    # ---- 2 失败用例明细 ----
    section = document.add_section("失败用例明细")
    failures = stats["all_failures"]
    if failures:
        section.add(Note(
            text=f"共 {len(failures)} 个用例未通过，需在合入前修复。以下为错误摘要，"
                 f"完整堆栈请查阅对应 TRX 文件。",
            level="danger",
        ))
        failure_rows = []
        for case in failures:
            failure_rows.append([
                case.class_name.rsplit(".", 1)[-1] or case.class_name,
                case.short_name,
                case.outcome,
                format_duration(case.duration_s),
                case.first_error_line(),
            ])
        section.add(Table(
            caption="表 2  失败用例与错误摘要",
            headers=["测试类", "测试方法", "结果", "耗时", "错误摘要"],
            rows=failure_rows,
            status_column=2,
        ))
    else:
        section.add(Note(text="本次汇总未发现失败用例。", level="success"))

    # ---- 3 测试分层视图 ----
    section = document.add_section("测试分层视图")
    by_assembly: Dict[str, Dict[str, int]] = {}
    for run in runs:
        for case in run.cases:
            assembly = case.class_name.split(".")[2] if case.class_name.count(".") >= 2 else "其他"
            bucket = by_assembly.setdefault(assembly, {"total": 0, "passed": 0, "failed": 0})
            bucket["total"] += 1
            if case.is_failure:
                bucket["failed"] += 1
            elif case.outcome == "Passed":
                bucket["passed"] += 1

    if by_assembly:
        layer_rows = []
        for assembly, bucket in sorted(by_assembly.items(), key=lambda kv: -kv[1]["total"]):
            rate = bucket["passed"] / bucket["total"] * 100 if bucket["total"] else 0.0
            layer_rows.append([
                assembly,
                str(bucket["total"]),
                str(bucket["passed"]),
                str(bucket["failed"]),
                f"{rate:.2f}%",
                "健康" if bucket["failed"] == 0 else "需关注",
            ])
        section.add(Table(
            caption="表 3  按测试项目分布（同一用例在多个 TFM 下重复统计）",
            headers=["测试项目", "用例数", "通过", "失败", "通过率", "状态"],
            rows=layer_rows,
            status_column=5,
            numeric_columns=(1, 2, 3, 4),
        ))

    # ---- 3.5 代码覆盖率 ----
    coverage = collect_coverage(layout.test_results_dir)
    section = document.add_section("代码覆盖率")
    if coverage.is_available:
        section.add(Metrics(items=[
            Metric(label="行覆盖率", value=f"{coverage.line_pct:.1f}%",
                   status="success" if coverage.line_rate >= 0.8 else
                          "warning" if coverage.line_rate >= 0.5 else "danger"),
            Metric(label="分支覆盖率", value=f"{coverage.branch_pct:.1f}%"),
            Metric(label="覆盖行数", value=f"{coverage.lines_covered}/{coverage.lines_valid}"),
            Metric(label="覆盖类数", value=str(coverage.class_count)),
        ]))
        package_rows = []
        for package in coverage.packages[:15]:
            package_rows.append([
                package.name,
                str(package.class_count),
                f"{package.line_pct:.1f}%",
                f"{sum(c.lines_covered for c in package.classes)}/"
                f"{sum(c.lines_valid for c in package.classes)}",
            ])
        section.add(Table(
            caption="表  各命名空间行覆盖率（前 15）",
            headers=["命名空间", "类数", "行覆盖率", "覆盖/有效行"],
            rows=package_rows,
            numeric_columns=(1, 2, 3),
        ))
        uncovered = coverage.uncovered_classes(threshold=0.5)
        if uncovered:
            section.add(Note(
                text=f"存在 {len(uncovered)} 个类行覆盖率低于 50%，建议补充测试。",
                level="warning",
            ))
            section.add(Table(
                caption="表  覆盖率低于 50% 的类（按覆盖率升序，前 20）",
                headers=["类", "命名空间", "行覆盖率", "覆盖/有效行"],
                rows=[
                    [cls.short_name, cls.namespace, f"{cls.line_pct:.1f}%",
                     f"{cls.lines_covered}/{cls.lines_valid}"]
                    for cls in uncovered[:20]
                ],
                numeric_columns=(2, 3),
            ))
    else:
        section.add(Note(
            text="未找到 Cobertura 覆盖率数据（coverage.cobertura.xml）。"
                 "CI 尚未启用覆盖率采集，详见阶段二路线图。",
            level="info",
        ))

    # ---- 4 基准测试结果 ----
    section = document.add_section("基准测试结果")
    timed = [(b, r) for b in benches for r in b.rows if r.mean_ms is not None]
    if timed:
        fastest = min(timed, key=lambda item: item[1].mean_ms)
        slowest = max(timed, key=lambda item: item[1].mean_ms)
        heaviest = max(
            ((b, r) for b, r in timed if r.allocated_bytes is not None),
            key=lambda item: item[1].allocated_bytes,
            default=None,
        )
        section.add(Metrics(items=[
            Metric(label="基准场景类", value=str(len(benches))),
            Metric(label="测量方法数", value=str(total_bench_methods)),
            Metric(label="最快", value=format_ms(fastest[1].mean_ms),
                   note=f"{fastest[0].display_name} / {fastest[1].method}"),
            Metric(label="最慢", value=format_ms(slowest[1].mean_ms), status="warning",
                   note=f"{slowest[0].display_name} / {slowest[1].method}"),
            *((
                Metric(label="最高分配", value=format_bytes(heaviest[1].allocated_bytes),
                       note=f"{heaviest[0].display_name} / {heaviest[1].method}"),
            ) if heaviest else ()),
        ]))

        for bench in benches:
            timed_rows = [r for r in bench.rows if r.mean_ms is not None]
            if not timed_rows:
                continue
            section.add(Para(text=f"{bench.display_name}（运行时 {bench.runtime or '未标注'}）"))

            table_rows = []
            for row in sorted(timed_rows, key=lambda r: r.mean_ms or 0):
                relative = row.relative_error
                table_rows.append([
                    row.method,
                    row.job,
                    format_ms(row.mean_ms),
                    f"{relative:.2f}%" if relative is not None else "—",
                    format_bytes(row.allocated_bytes),
                    f"{row.gen0:.4f}",
                    row.row_count,
                ])
            section.add(Table(
                caption=f"表  基准类 {bench.display_name}",
                headers=["方法", "作业", "平均耗时", "相对误差", "分配内存", "Gen0", "数据行数"],
                rows=table_rows,
                numeric_columns=(2, 3, 4, 5, 6),
            ))
            section.add(Chart(
                title=f"{bench.display_name} · 平均耗时（毫秒，越小越好）",
                kind="bar",
                horizontal=True,
                categories=[r.method for r in sorted(timed_rows, key=lambda r: r.mean_ms or 0)],
                series=[("平均耗时", [
                    round(r.mean_ms or 0.0, 4)
                    for r in sorted(timed_rows, key=lambda r: r.mean_ms or 0)
                ])],
                unit="毫秒 (ms)",
            ))
    else:
        section.add(Note(text="未找到 BenchmarkDotNet 结果文件。", level="warning"))

    # ---- 4.5 基准趋势（与历史归档对比） ----
    snapshots = snapshot_from_sets(benches, layout.root)
    history_dir = layout.benchmark_history_dir
    trends = compare_with_previous(snapshots, history_dir)

    section = document.add_section("基准趋势对比")
    if trends:
        regressions = [
            p for points in trends.values() for p in points if p.verdict == "regression"
        ]
        improvements = [
            p for points in trends.values() for p in points if p.verdict == "improvement"
        ]
        section.add(Metrics(items=[
            Metric(label="对比基准类", value=str(len(trends))),
            Metric(label="性能回归", value=str(len(regressions)),
                   status="danger" if regressions else "success"),
            Metric(label="性能优化", value=str(len(improvements)), status="success"),
        ]))
        trend_rows: List[List[str]] = []
        for name, points in sorted(trends.items()):
            display = name.rsplit(".", 1)[-1]
            for point in sorted(points, key=lambda p: -p.delta_ratio):
                trend_rows.append([
                    display,
                    point.method,
                    format_ms(point.previous_ms),
                    format_ms(point.current_ms),
                    f"{point.delta_ratio:+.1f}%",
                    {"regression": "回归", "improvement": "优化", "stable": "持平"}[point.verdict],
                ])
        section.add(Table(
            caption="表  方法级性能变化（相对上一次归档，阈值 ±10%）",
            headers=["基准类", "方法", "上次耗时", "本次耗时", "变化率", "判定"],
            rows=trend_rows,
            status_column=5,
            numeric_columns=(2, 3, 4),
        ))
        if regressions:
            section.add(Note(
                text=f"发现 {len(regressions)} 处性能回归（变化率 > +10%），"
                     f"建议核查相关改动。",
                level="danger",
            ))
    else:
        section.add(Note(
            text="尚未建立基准历史归档。本次结果将作为基线写入 "
                 "artifacts/benchmark-history/，后续生成即可做趋势对比。",
            level="info",
        ))

    # 归档本次结果（追加写，不覆盖历史）；--check 只读模式不归档
    written = write_snapshots(snapshots, history_dir) if archive else []

    # ---- 5 结论与风险 ----
    section = document.add_section("结论与风险")
    conclusions = [
        f"测试通过率 {stats['pass_rate']:.2f}%，"
        + ("质量门禁通过。" if stats["failed"] == 0 else "存在失败用例，质量门禁未通过。"),
        f"基准数据来自 {len(benches)} 个场景类；NPOI 管线为 DOM 模型，"
        f"耗时可接受但内存峰值需按部署容量规划。",
        "异步 API 仅异步化外围 IO，DOM 阶段仍占用调用线程，高并发场景需限制并发度。",
        f"本次基准结果已归档 {len(written)} 份快照至 {history_dir.name}，供后续趋势对比。",
    ]
    section.add(Bullets(items=conclusions))
    if stats["failed"]:
        section.add(Note(
            text="本报告存在未通过用例，不应作为发布依据；请先修复后重新生成。",
            level="danger",
        ))
    else:
        section.add(Note(text="全部用例通过，可作为发布质量证据归档。", level="success"))

    return document


# ==========================================================================
# 2. 发布说明
# ==========================================================================
def build_release_notes(layout: Optional[RepoLayout] = None,
                        theme: Theme = DEFAULT_THEME,
                        limit: int = 120) -> Document:
    layout = layout or default_layout()
    changelog = collect_changelog(layout.root, limit=limit)
    api = collect_api(layout.release_output_dir, layout.api_baseline)

    document = Document(
        title="发布说明与变更日志",
        subtitle=f"{theme.brand_name} 版本变更记录",
        doc_id="BO-RN-001",
        version=api.package_version or "—",
        author="文档自动化流水线",
        department=theme.organization,
        generated_at=_now(),
        source_commit=changelog.head_sha[:12] if changelog.head_sha else "",
        classification="公开",
        filename="bing-offices-release-notes",
        theme=theme,
    )

    grouped = changelog.by_type()
    user_facing = changelog.by_type(user_facing_only=True)
    type_counts = changelog.type_counts()
    breaking = changelog.breaking_changes

    top_types = "、".join(f"{t} {c}" for t, c in list(type_counts.items())[:3])
    document.summary = (
        f"本版本基于 {changelog.branch or 'HEAD'} 分支共 {len(changelog.entries)} 次提交，"
        f"其中面向使用者的变更 {sum(len(v) for v in user_facing.values())} 项"
        f"（{top_types}）。"
        f"公共 API 快照 Schema 为 {api.schema or '—'}，"
        f"覆盖 {len(api.assemblies)} 个程序集、合计 {api.total_members} 个公共成员。"
    )
    document.highlights = [
        f"提交总数 {len(changelog.entries)}，参与变更范围 {len(changelog.scope_counts)} 个",
        f"提交规范符合率 {changelog.conventional_ratio:.1f}%"
        + ("（低于 80%，建议统一提交规范以提升变更日志质量）"
           if changelog.conventional_ratio < 80 else ""),
        f"破坏性变更 {len(breaking)} 项",
        f"最近标签 {changelog.latest_tag or '无'}"
        + (f"（共 {changelog.tag_count} 个标签）" if changelog.tag_count else ""),
    ]

    # ---- 1 版本概览 ----
    section = document.add_section("版本概览")
    section.add(Metrics(items=[
        Metric(label="版本号", value=api.package_version or "—"),
        Metric(label="提交数", value=str(len(changelog.entries))),
        Metric(label="新功能", value=str(len(grouped.get("feat", []))), status="success"),
        Metric(label="缺陷修复", value=str(len(grouped.get("fix", []))), status="success"),
        Metric(label="破坏性变更", value=str(len(breaking)),
               status="danger" if breaking else "success"),
        Metric(
            label="提交规范符合率",
            value=f"{changelog.conventional_ratio:.1f}%",
            status="success" if changelog.conventional_ratio >= 80 else "warning",
            note=f"{changelog.conventional_count}/{len(changelog.entries)} 条可自动归类",
        ),
    ]))
    section.add(Table(
        caption="表 1  变更类型分布",
        headers=["类型", "数量", "是否计入发布说明"],
        rows=[
            [
                CHANGE_TYPES.get(type_key, ("其他", True))[0],
                str(count),
                "是" if CHANGE_TYPES.get(type_key, ("其他", True))[1] else "否",
            ]
            for type_key, count in type_counts.items()
        ],
        numeric_columns=(1,),
    ))

    # ---- 2 新增功能 ----
    section = document.add_section("新增功能")
    features = grouped.get("feat", [])
    if features:
        section.add(Bullets(items=[
            f"{entry.scope or 'general'}：{entry.subject}（{entry.short_sha}，{entry.date}）"
            for entry in features
        ]))
    else:
        section.add(Para(text="本次范围内没有新增功能。"))

    # ---- 3 缺陷修复 ----
    section = document.add_section("缺陷修复")
    fixes = grouped.get("fix", [])
    if fixes:
        section.add(Bullets(items=[
            f"{entry.scope or 'general'}：{entry.subject}（{entry.short_sha}，{entry.date}）"
            for entry in fixes
        ]))
    else:
        section.add(Para(text="本次范围内没有缺陷修复。"))

    # ---- 4 性能与回滚 ----
    section = document.add_section("性能优化与其他变更")
    other_rows = []
    for type_name in ("perf", "refactor", "docs", "test", "build", "ci", "chore", "style", "other"):
        items = grouped.get(type_name, [])
        if not items:
            continue
        label = CHANGE_TYPES.get(type_name, ("其他", True))[0]
        other_rows.append([label, str(len(items)), items[0].subject[:70]])
    if other_rows:
        section.add(Table(
            caption="表 2  非面向使用者的变更",
            headers=["类型", "数量", "最近一条"],
            rows=other_rows,
            numeric_columns=(1,),
        ))
    if breaking:
        section.add(Note(
            text=f"存在 {len(breaking)} 项破坏性变更，升级前请查阅迁移说明并评估影响面。",
            level="danger",
        ))

    # ---- 5 公共 API 变更 ----
    section = document.add_section("公共 API 快照")
    if api.assemblies:
        section.add(Table(
            caption="表 3  公共 API 成员统计",
            headers=["程序集", "目标框架", "公共成员数", "快照哈希（前 12 位）", "XML 注释覆盖率"],
            rows=[
                [
                    assembly.name,
                    assembly.tfm,
                    str(assembly.member_count),
                    assembly.hash_short or "—",
                    f"{assembly.coverage:.1f}%",
                ]
                for assembly in api.assemblies
            ],
            numeric_columns=(2, 4),
        ))
        section.add(Para(
            text=f"快照 Schema：{api.schema}；生成器版本：{api.generator_version}；"
                 f"基线提交：{api.baseline_commit}；审批：{api.approved_by}（{api.approved_at}）。"
        ))
    else:
        section.add(Note(text="未找到 build/api-snapshot-baseline.json，无法给出 API 统计。",
                         level="warning"))

    # ---- 6 变更分布图 ----
    section = document.add_section("变更分布")
    if type_counts:
        labels = [CHANGE_TYPES.get(k, ("其他", True))[0] for k in type_counts]
        section.add(Chart(
            title="提交类型分布",
            kind="pie",
            categories=labels,
            series=[("提交数", [float(v) for v in type_counts.values()])],
        ))
    scopes = changelog.scope_counts
    if scopes:
        top_scopes = dict(list(scopes.items())[:8])
        section.add(Chart(
            title="变更范围（scope）分布 · 前 8",
            kind="bar",
            categories=list(top_scopes.keys()),
            series=[("提交数", [float(v) for v in top_scopes.values()])],
            unit="提交数",
        ))

    # ---- 7 升级提示 ----
    section = document.add_section("升级提示")
    if changelog.conventional_ratio < 90:
        section.add(Note(
            text=f"当前提交规范符合率仅 {changelog.conventional_ratio:.1f}%"
                 f"（{changelog.conventional_count}/{len(changelog.entries)}）。"
                 f"不符合规范的提交会落入「其他」类，只能整条列出、无法按范围聚合。"
                 f"建议统一采用 Conventional Commits（本仓库已是「emoji + 类型(范围): 描述」风格，"
                 f"继续沿用即可）。",
            level="warning",
        ))
    section.add(Bullets(items=[
        "Runtime 支持：Bing.Offices.Npoi 提供 net6.0 / net8.0 资产；"
        "Abstractions 与 Core 保持 netstandard2.0。",
        "Async API 不承诺 DOM 阶段异步：NPOI 的 WorkbookFactory.Create、DOM 操作与 "
        "workbook.Write 均为同步阶段。",
        "模板流必须同步可读，async-only 模板流不在支持的输入合同内。",
        "文件原子提交保证失败或取消不会截断已有目标文件；直接写调用方 Stream 无回滚保证。",
    ]))
    section.add(Note(text=GENERATED_NOTE, level="info"))

    return document


# ==========================================================================
# 3. 汇报材料
# ==========================================================================
def build_status_deck(layout: Optional[RepoLayout] = None,
                      theme: Theme = DEFAULT_THEME) -> Document:
    layout = layout or default_layout()
    runs = collect_test_runs(layout.test_results_dir)
    benches = collect_benchmarks(layout.benchmark_dir)
    api = collect_api(layout.release_output_dir, layout.api_baseline)
    changelog = collect_changelog(layout.root, limit=60)

    stats = summarize(runs)
    grouped = changelog.by_type()
    total_bench = sum(len(b.rows) for b in benches)

    document = Document(
        title=f"{theme.brand_name} 项目进展汇报",
        subtitle="质量、性能与交付节奏",
        doc_id="BO-DECK-001",
        version=api.package_version or "—",
        author="文档自动化流水线",
        generated_at=_now(),
        source_commit=changelog.head_sha[:12] if changelog.head_sha else "",
        classification="内部资料",
        filename="bing-offices-status-deck",
        theme=theme,
    )

    document.summary = (
        f"测试通过率 {stats['pass_rate']:.2f}%（{stats['passed']}/{stats['executed']}），"
        f"基准场景 {len(benches)} 组 / {total_bench} 个方法，"
        f"近期提交 {len(changelog.entries)} 次，公共成员 {api.total_members} 个。"
    )
    document.highlights = [
        f"质量：通过率 {stats['pass_rate']:.2f}%，失败 {stats['failed']} 个",
        f"性能：{len(benches)} 组基准覆盖导入、导出、CSV、映射与失败工作簿路径",
        f"交付：新增功能 {len(grouped.get('feat', []))} 项，修复 {len(grouped.get('fix', []))} 项",
        f"API 面：{len(api.assemblies)} 个程序集 / {api.total_members} 个公共成员",
    ]

    # ---- 1 项目概览 ----
    section = document.add_section("项目概览")
    section.add(Metrics(items=[
        Metric(label="包版本", value=api.package_version or "—"),
        Metric(label="公共成员", value=str(api.total_members)),
        Metric(label="测试通过率", value=f"{stats['pass_rate']:.2f}%",
               status=_status_of(stats["failed"] == 0)),
        Metric(label="基准方法", value=str(total_bench)),
    ]))

    # ---- 2 质量指标 ----
    section = document.add_section("质量指标")
    rows = []
    for run in sorted(runs, key=lambda r: (r.suite, r.target_framework)):
        rows.append([
            f"{run.suite} · {run.target_framework or '—'}",
            str(run.total), str(run.passed), str(run.failed),
            f"{run.pass_rate:.2f}%",
            "通过" if run.is_success else "未通过",
        ])
    section.add(Table(
        caption="各测试套件执行结果",
        headers=["套件", "用例数", "通过", "失败", "通过率", "结论"],
        rows=rows,
        status_column=5,
        numeric_columns=(1, 2, 3, 4),
    ))
    if runs:
        section.add(Chart(
            title="各套件通过 / 失败用例数",
            kind="bar",
            categories=[f"{r.suite}\n{r.target_framework or ''}".strip() for r in runs],
            series=[
                ("通过", [float(r.passed) for r in runs]),
                ("失败", [float(r.failed) for r in runs]),
            ],
            unit="用例数",
        ))

    # ---- 3 性能表现 ----
    section = document.add_section("性能表现")
    timed = [(b, r) for b in benches for r in b.rows if r.mean_ms is not None]
    if timed:
        top = sorted(timed, key=lambda item: item[1].mean_ms, reverse=True)[:8]
        section.add(Chart(
            title="耗时最高的 8 个基准方法（毫秒）",
            kind="bar",
            horizontal=True,
            categories=[f"{b.display_name} / {r.method}" for b, r in top],
            series=[("平均耗时", [round(r.mean_ms or 0.0, 3) for _, r in top])],
            unit="毫秒 (ms)",
        ))
        lightest = sorted([t for t in timed if t[1].allocated_bytes is not None],
                          key=lambda item: item[1].allocated_bytes)[:8]
        if lightest:
            section.add(Chart(
                title="分配内存最低的 8 个基准方法（字节）",
                kind="bar",
                horizontal=True,
                categories=[f"{b.display_name} / {r.method}" for b, r in lightest],
                series=[("分配字节", [round(r.allocated_bytes or 0.0, 1) for _, r in lightest])],
                unit="字节 (B)",
            ))
    else:
        section.add(Note(text="未找到基准测试结果。", level="warning"))

    # ---- 4 近期交付 ----
    section = document.add_section("近期交付")
    recent = [e for e in changelog.entries if e.is_user_facing][:10]
    if recent:
        section.add(Bullets(items=[
            f"{e.type_cn}｜{e.scope or 'general'}：{e.subject}（{e.short_sha}）" for e in recent
        ]))
    else:
        section.add(Para(text="近期没有面向使用者的交付项。"))

    # ---- 5 风险与下一步 ----
    section = document.add_section("风险与下一步")
    risks = [
        "NPOI 为 DOM 模型，MaxInputBytes 不覆盖解压后峰值内存，"
        "不受信任文件需额外设置进程内存与 CPU 配额。",
        "未被映射的图片不进入图片限制器扫描范围，存在残留攻击面。",
        "计划缓存无 TTL 且为进程级共享，动态生成大量映射时需关注缓存抖动。",
        "AtomicFileCommitter 的 Exists → Replace/Move 存在 TOCTOU 竞态，"
        "同一路径应避免并发写入。",
    ]
    section.add(Bullets(items=risks))
    next_steps = [
        "把文档生成接入 CI，随构建产出测试与基准报告并作为 Artifact 归档。",
        "将 XML 注释覆盖率纳入质量门禁，低于阈值时阻断合入。",
        "对基准结果建立历史趋势基线，识别性能回归。",
    ]
    section.add(Para(text="下一步计划："))
    section.add(Bullets(items=next_steps, ordered=True))

    return document


# ==========================================================================
# 4. 工程内部技术文档：API 参考
# ==========================================================================
def build_api_reference(layout: Optional[RepoLayout] = None,
                        theme: Theme = DEFAULT_THEME,
                        max_types_per_namespace: int = 60) -> Document:
    layout = layout or default_layout()
    api = collect_api(layout.release_output_dir, layout.api_baseline)
    changelog = collect_changelog(layout.root, limit=1)

    document = Document(
        title="公共 API 参考",
        subtitle=f"{theme.brand_name} · 从源码注释与 API 快照自动生成",
        doc_id="BO-APIDOC-001",
        version=api.package_version or "—",
        author="文档自动化流水线",
        department=theme.organization,
        generated_at=_now(),
        source_commit=changelog.head_sha[:12] if changelog.head_sha else "",
        classification="内部资料",
        filename="bing-offices-api-reference",
        theme=theme,
    )

    public_types = api.public_types()
    document.summary = (
        f"本文档由 output/release 下的 XML 文档注释与 build/api-snapshot-baseline.json "
        f"公共 API 快照合并生成，覆盖 {len(api.assemblies)} 个程序集、"
        f"{api.total_members} 个公共成员、{len(public_types)} 个已记录公共类型。"
        f"注释与代码同源，代码变更后重新生成即可保持同步。"
    )
    document.highlights = [
        f"程序集 {', '.join(a.name.replace('Bing.Offices.', '') for a in api.assemblies)}",
        f"公共成员合计 {api.total_members} 个",
        f"已记录公共类型 {len(public_types)} 个，命名空间 {len(api.by_namespace())} 个",
    ]

    # ---- 1 程序集清单 ----
    section = document.add_section("程序集清单")
    if api.assemblies:
        section.add(Table(
            caption="表 1  程序集与注释覆盖",
            headers=["程序集", "目标框架", "公共成员数", "已写注释成员", "快照哈希（前 12 位）"],
            rows=[
                [a.name, a.tfm, str(a.member_count), str(len(a.documented)), a.hash_short or "—"]
                for a in api.assemblies
            ],
            numeric_columns=(2, 3),
        ))
        section.add(Metrics(items=[
            Metric(label="程序集", value=str(len(api.assemblies))),
            Metric(label="公共成员", value=str(api.total_members)),
            Metric(label="已记录类型", value=str(len(public_types))),
            Metric(label="命名空间", value=str(len(api.by_namespace()))),
        ]))
    else:
        section.add(Note(text="未找到 API 快照或 XML 注释产物。", level="warning"))

    # ---- 2 成员种类分布 ----
    section = document.add_section("成员种类分布")
    kinds = api.snapshot_kinds
    if kinds:
        section.add(Table(
            caption="表 2  公共成员按种类统计（快照口径）",
            headers=["成员种类", "数量"],
            rows=[[k, str(v)] for k, v in sorted(kinds.items(), key=lambda kv: -kv[1])],
            numeric_columns=(1,),
        ))
        section.add(Chart(
            title="公共成员种类分布",
            kind="pie",
            categories=list(kinds.keys())[:8],
            series=[("数量", [float(v) for v in list(kinds.values())[:8]])],
        ))

    # ---- 3 命名空间与类型 ----
    section = document.add_section("命名空间与类型清单")
    grouped = api.by_namespace()
    for namespace, types in grouped.items():
        section.add(Para(text=f"{namespace}（{len(types)} 个类型）"))
        shown = types[:max_types_per_namespace]
        section.add(Table(
            caption="",
            headers=["类型", "说明"],
            rows=[[item.short_name, item.summary or "（未写注释）"] for item in shown],
        ))
        if len(types) > len(shown):
            section.add(Note(
                text=f"该命名空间另有 {len(types) - len(shown)} 个类型未在此列出。",
                level="info",
            ))

    # ---- 3.5 典型用法（来自可编译源码，防止文档腐化） ----
    samples = collect_source_samples(layout.docs_tests_dir)
    if samples:
        section = document.add_section("典型用法（可编译示例）")
        section.add(Note(
            text="以下示例全部取自 "
                 f"{layout.docs_tests_dir.name} 中带 // docgen:sample 标记的方法，"
                 "与项目一起编译。API 变更导致示例失效时构建即失败，"
                 "因此文档不会在无人察觉的情况下腐化。",
            level="info",
        ))
        section.add(Table(
            caption="表  示例索引",
            headers=["示例", "标题", "出处"],
            rows=[[s.name, s.title, f"{s.source}:{s.line}"] for s in samples],
        ))
        for sample in samples:
            section.add(Para(text=f"{sample.title}（{sample.name}）"))
            section.add(Code(
                text=f"// docgen:sample name={sample.name} title={sample.title}\n{sample.code}",
                language="csharp",
            ))

    # ---- 4 使用约束 ----
    section = document.add_section("使用约束与运行时边界")
    section.add(Bullets(items=[
        "NPOI 类型（IWorkbook / ISheet / ICell）不出现在任何公共契约中，"
        "Provider 只能看到 provider-neutral 的映射计划视图。",
        "Async API 只异步化外围文件 / Stream IO；DOM 构建、映射、校验与序列化仍在调用线程执行。",
        "生产代码禁止使用 Task.Run / .Result / .Wait 伪装异步。",
        "模板流必须在 WorkbookFactory.Create 阶段同步可读。",
        "Excel 图表仅支持 XLSX（依赖 XSSFWorkbook），.xls 请求图表会抛 NotSupportedException。",
    ]))
    section.add(Note(
        text="本文档为自动生成产物。修改 API 后请重新运行文档流水线，"
             "并在 CI 中校验生成的公共 API 与 XML 注释覆盖率。",
        level="info",
    ))

    return document


# ==========================================================================
# 5. 对外交付文档
# ==========================================================================
_DEFAULT_TERMS = [
    "本文件所载测试数据与基准数据来自交付方内部质量流水线，"
    "数据对应的代码版本以「数据来源」所记提交为准。",
    "本文件描述的 API 以对应 NuGet 包版本的公共契约为准，"
    "未在本文档列出的成员不在支持范围内。",
    "Bing.Offices.Npoi 提供 net6.0 / net8.0 资产；"
    "Bing.Offices.Abstractions 与 Bing.Offices.Core 为 netstandard2.0。",
    "异步 API 仅异步化外围文件 / Stream IO；DOM 构建、映射、校验与序列化阶段"
    "仍在调用线程执行，不承诺降低调用线程占用。",
    "未经交付方书面许可，不得向第三方披露本文件的全部或部分内容。",
]

_DEFAULT_CONFIDENTIAL = (
    "本文件含未公开的架构与性能数据，仅限接收方项目相关人员内部使用；"
    "对外披露需取得交付方书面同意。"
)
_DEFAULT_WARRANTY = (
    "本文件为交付说明，不构成性能承诺；实际表现取决于运行环境、数据规模与调用方式。"
    "质保与支持范围以双方签署的合同为准。"
)


def load_delivery_profile(layout: RepoLayout) -> Delivery:
    """读取对外交付档案；缺失时回落到内置默认条款，保证流水线永远可运行。"""
    import json

    data = {}
    path = layout.delivery_profile
    if path.exists():
        try:
            with path.open("r", encoding="utf-8-sig") as handle:
                data = json.load(handle)
        except (ValueError, OSError):
            data = {}

    return Delivery(
        client=str(data.get("client", "") or "（未填写接收方）"),
        project=str(data.get("project", "") or "Bing.Offices 组件交付"),
        contact=str(data.get("contact", "") or ""),
        owner=str(data.get("owner", "") or "交付方项目负责人"),
        effective_date=str(data.get("effectiveDate", "") or ""),
        confidential=str(data.get("confidential", "") or _DEFAULT_CONFIDENTIAL),
        warranty=str(data.get("warranty", "") or _DEFAULT_WARRANTY),
        terms=[str(item) for item in data.get("terms", [])] or list(_DEFAULT_TERMS),
        signatures=[str(item) for item in data.get("signatures", [])]
        or ["客户方负责人（签章）", "交付方负责人（签章）"],
    )


def build_delivery_brief(layout: Optional[RepoLayout] = None,
                         theme: Theme = DEFAULT_THEME) -> Document:
    """对外交付文档：客户抬头 + 质量证据 + 条款页 + 签章页。"""
    layout = layout or default_layout()
    runs = collect_test_runs(layout.test_results_dir)
    benches = collect_benchmarks(layout.benchmark_dir)
    api = collect_api(layout.release_output_dir, layout.api_baseline)
    changelog = collect_changelog(layout.root, limit=40)

    stats = summarize(runs)
    total_bench = sum(len(b.rows) for b in benches)
    delivery = load_delivery_profile(layout)

    document = Document(
        title="Bing.Offices 交付说明与质量证据",
        subtitle=f"{theme.brand_name} {api.package_version or ''} 组件交付件".strip(),
        doc_id="BO-DEL-001",
        version=api.package_version or "—",
        author=delivery.owner or "文档自动化流水线",
        department=theme.organization,
        generated_at=_now(),
        source_commit=changelog.head_sha[:12] if changelog.head_sha else "",
        classification="对外 · 受控",
        filename="bing-offices-delivery-brief",
        theme=theme,
        delivery=delivery,
    )

    document.summary = (
        f"本文件说明 {theme.brand_name}（版本 {api.package_version or '—'}）的交付内容、"
        f"质量证据与使用约束。交付范围涵盖 {len(api.assemblies)} 个程序集、"
        f"{api.total_members} 个公共成员；质量证据基于 {stats['total']} 个测试用例"
        f"（通过率 {stats['pass_rate']:.2f}%）与 {total_bench} 条基准测量。"
    )
    document.highlights = [
        f"交付版本 {api.package_version or '—'}，程序集 {len(api.assemblies)} 个",
        f"质量：{stats['total']} 个测试用例，通过率 {stats['pass_rate']:.2f}%，失败 {stats['failed']} 个",
        f"性能：{len(benches)} 组基准场景、{total_bench} 条测量记录",
        "NPOI 类型不进入任何公共契约，Provider 可替换",
    ]

    # ---- 1 交付概览 ----
    section = document.add_section("交付概览")
    section.add(Metrics(items=[
        Metric(label="交付版本", value=api.package_version or "—"),
        Metric(label="程序集", value=str(len(api.assemblies))),
        Metric(label="公共成员", value=str(api.total_members)),
        Metric(label="测试通过率", value=f"{stats['pass_rate']:.2f}%",
               status=_status_of(stats["failed"] == 0)),
        Metric(label="基准测量", value=str(total_bench)),
    ]))

    # ---- 2 交付内容清单 ----
    section = document.add_section("交付内容清单")
    if api.assemblies:
        section.add(Table(
            caption="表 1  交付程序集清单",
            headers=["程序集", "目标框架", "公共成员数", "说明"],
            rows=[
                [
                    assembly.name,
                    assembly.tfm,
                    str(assembly.member_count),
                    {
                        "Bing.Offices.Abstractions": "契约层：Provider 中立的类型与映射计划视图",
                        "Bing.Offices.Core": "引擎层：映射编译、校验、CSV 管线与原子提交",
                        "Bing.Offices.Npoi": "适配层：基于 NPOI 的 XLSX / XLS 实现",
                    }.get(assembly.name, "—"),
                ]
                for assembly in api.assemblies
            ],
            numeric_columns=(2,),
        ))
    else:
        section.add(Note(text="未找到 API 快照，无法列出交付清单。", level="warning"))

    section.add(Bullets(items=[
        "NuGet 包：Bing.Offices.Abstractions / Bing.Offices.Core / Bing.Offices.Npoi",
        "公共 API 快照：build/api-snapshot-baseline.json（用于变更审批与契约比对）",
        "XML 文档注释：随包发布，IDE 内可直接查看中文说明",
        "自动生成文档：API 参考、发布说明、测试与基准报告",
    ]))

    # ---- 3 质量证据 ----
    section = document.add_section("质量证据")
    rows = []
    for run in sorted(runs, key=lambda r: (r.suite, r.target_framework)):
        rows.append([
            run.suite,
            run.target_framework or "—",
            str(run.total),
            str(run.passed),
            str(run.failed),
            f"{run.pass_rate:.2f}%",
            "通过" if run.is_success else "未通过",
        ])
    if rows:
        section.add(Table(
            caption="表 2  测试执行结果（各测试套件）",
            headers=["测试套件", "目标框架", "用例数", "通过", "失败", "通过率", "结论"],
            rows=rows,
            status_column=6,
            numeric_columns=(2, 3, 4, 5),
        ))
    else:
        section.add(Note(text="未找到测试结果（TestResults/*.trx）。", level="warning"))

    if stats["failed"]:
        section.add(Note(
            text=f"当前存在 {stats['failed']} 个未通过用例。"
                 f"本文件在用例全部通过后才可作为正式交付依据。",
            level="danger",
        ))
    else:
        section.add(Note(text="全部用例通过，可作为交付质量证据。", level="success"))

    # ---- 4 性能证据 ----
    section = document.add_section("性能证据")
    timed = [(b, r) for b in benches for r in b.rows if r.mean_ms is not None]
    if timed:
        fastest = min(timed, key=lambda item: item[1].mean_ms)
        slowest = max(timed, key=lambda item: item[1].mean_ms)
        section.add(Metrics(items=[
            Metric(label="基准场景", value=str(len(benches))),
            Metric(label="测量方法", value=str(total_bench)),
            Metric(label="最快", value=format_ms(fastest[1].mean_ms),
                   note=f"{fastest[0].display_name} / {fastest[1].method}"),
            Metric(label="最慢", value=format_ms(slowest[1].mean_ms), status="warning",
                   note=f"{slowest[0].display_name} / {slowest[1].method}"),
        ]))
        top = sorted(timed, key=lambda item: item[1].mean_ms, reverse=True)[:10]
        section.add(Table(
            caption="表 3  耗时最高的 10 个基准方法",
            headers=["基准场景", "方法", "平均耗时", "分配内存"],
            rows=[
                [b.display_name, r.method, format_ms(r.mean_ms), format_bytes(r.allocated_bytes)]
                for b, r in top
            ],
            numeric_columns=(2, 3),
        ))
    else:
        section.add(Note(text="未找到 BenchmarkDotNet 结果文件。", level="warning"))
    section.add(Note(
        text="基准数据在本机环境测得，仅供量级参考；生产环境请以实际数据规模复测。",
        level="info",
    ))

    # ---- 5 使用约束 ----
    section = document.add_section("使用约束与运行环境")
    section.add(Bullets(items=[
        "NPOI 类型（IWorkbook / ISheet / ICell）不出现在任何公共契约中，Provider 可替换。",
        "Async API 只异步化外围文件 / Stream IO；DOM 阶段仍在调用线程执行。",
        "模板流必须在读取阶段同步可读，async-only 模板流不在支持的输入合同内。",
        "Excel 图表仅支持 XLSX（依赖 XSSFWorkbook），.xls 请求图表会抛 NotSupportedException。",
        "文件原子提交保证失败或取消不会截断已有目标文件；"
        "直接写调用方 Stream 无回滚保证。",
    ]))

    # ---- 6 近期变更 ----
    section = document.add_section("近期变更")
    recent = [e for e in changelog.entries if e.is_user_facing][:12]
    if recent:
        section.add(Bullets(items=[
            f"{e.type_cn}｜{e.scope or 'general'}：{e.subject}（{e.short_sha}）" for e in recent
        ]))
    else:
        section.add(Para(text="近期没有面向使用者的变更。"))
    section.add(Note(text=GENERATED_NOTE, level="info"))

    return document


BUILDERS = {
    "tests": build_test_benchmark_report,
    "release": build_release_notes,
    "deck": build_status_deck,
    "api": build_api_reference,
    "delivery": build_delivery_brief,
}
