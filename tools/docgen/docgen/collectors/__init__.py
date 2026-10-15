"""采集层：把仓库原始产物转换为结构化数据。

约定：采集器 **只返回纯数据对象**，不构造 Document，也不接触渲染。
"""

from .api import ApiInventory, ApiMemberDoc, collect_api
from .archives import (
    BenchmarkSnapshot,
    TrendPoint,
    compare_with_previous,
    load_history,
    snapshot_from_sets,
    write_snapshots,
)
from .benchmark import BenchmarkRow, BenchmarkSet, collect_benchmarks
from .changelog import ChangeEntry, ChangeLog, collect_changelog
from .coverage import ClassCoverage, CoverageInventory, PackageCoverage, collect_coverage
from .tests import TestCase, TestRun, collect_test_runs

__all__ = [
    "ApiInventory",
    "ApiMemberDoc",
    "collect_api",
    "BenchmarkSnapshot",
    "TrendPoint",
    "compare_with_previous",
    "load_history",
    "snapshot_from_sets",
    "write_snapshots",
    "BenchmarkRow",
    "BenchmarkSet",
    "collect_benchmarks",
    "ChangeEntry",
    "ChangeLog",
    "collect_changelog",
    "ClassCoverage",
    "CoverageInventory",
    "PackageCoverage",
    "collect_coverage",
    "TestCase",
    "TestRun",
    "collect_test_runs",
]
