"""基准测试采集器：解析 BenchmarkDotNet 宽表 CSV。

数据源：BenchmarkDotNet.Artifacts/results/*-report.csv
产出：BenchmarkSet 列表，每个基准类一组，含 Mean / 误差 / GC 分配。
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

_SIZE_UNITS = {
    "B": 1.0,
    "KB": 1024.0,
    "MB": 1024.0**2,
    "GB": 1024.0**3,
}

_TIME_UNITS = {
    "ns": 1e-6,
    "us": 1e-3,
    "µs": 1e-3,
    "ms": 1.0,
    "s": 1000.0,
}

_VALUE_RE = re.compile(r"^([\d.]+)\s*([A-Za-zµ]+)$")


def parse_time_ms(value: str) -> Optional[float]:
    """'16.438 ms' -> 16.438（毫秒）"""
    if not value:
        return None
    match = _VALUE_RE.match(value.strip())
    if not match:
        return None
    number, unit = float(match.group(1)), match.group(2)
    factor = _TIME_UNITS.get(unit)
    return number * factor if factor is not None else None


def parse_bytes(value: str) -> Optional[float]:
    """'2.22 MB' -> 字节数"""
    if not value:
        return None
    match = _VALUE_RE.match(value.strip())
    if not match:
        return None
    number, unit = float(match.group(1)), match.group(2)
    factor = _SIZE_UNITS.get(unit)
    return number * factor if factor is not None else None


def format_bytes(value: Optional[float]) -> str:
    if value is None:
        return "—"
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.2f} {unit}" if unit != "B" else f"{value:.0f} B"
        value /= 1024
    return f"{value:.2f} GB"


def format_ms(value: Optional[float]) -> str:
    if value is None:
        return "—"
    if value < 1:
        return f"{value * 1000:.1f} µs"
    return f"{value:.3f} ms"


@dataclass
class BenchmarkRow:
    """单条基准结果。"""

    method: str
    job: str = ""
    runtime: str = ""
    mean_ms: Optional[float] = None
    error_ms: Optional[float] = None
    allocated_bytes: Optional[float] = None
    gen0: float = 0.0
    row_count: str = ""

    @property
    def relative_error(self) -> Optional[float]:
        if self.mean_ms and self.error_ms is not None and self.mean_ms > 0:
            return self.error_ms / self.mean_ms * 100.0
        return None


@dataclass
class BenchmarkSet:
    """一个基准类的全部结果。"""

    name: str
    runtime: str = ""
    source: str = ""
    rows: List[BenchmarkRow] = field(default_factory=list)

    @property
    def display_name(self) -> str:
        """Bing.Offices.Benchmarks.CsvPipelineBenchmarks -> CsvPipeline"""
        name = self.name.rsplit(".", 1)[-1]
        return re.sub(r"Benchmarks$", "", name) or name

    @property
    def fastest(self) -> Optional[BenchmarkRow]:
        timed = [r for r in self.rows if r.mean_ms is not None]
        return min(timed, key=lambda r: r.mean_ms) if timed else None

    @property
    def slowest(self) -> Optional[BenchmarkRow]:
        timed = [r for r in self.rows if r.mean_ms is not None]
        return max(timed, key=lambda r: r.mean_ms) if timed else None


def _read_bdn_csv(path: Path) -> BenchmarkSet:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        try:
            header = next(reader)
        except StopIteration:
            return BenchmarkSet(name=path.stem)

        # 宽表 CSV 可能出现重复列名，统一取**首个**出现位置，避免歧义
        index: Dict[str, int] = {}
        for position, column in enumerate(header):
            index.setdefault(column.strip(), position)

        def cell(row: List[str], column: str) -> str:
            position = index.get(column)
            if position is None or position >= len(row):
                return ""
            return (row[position] or "").strip()

        result = BenchmarkSet(name=path.stem, source=path.name)
        for row in reader:
            if not row or not row[0].strip():
                continue
            entry = BenchmarkRow(
                method=cell(row, "Method"),
                job=cell(row, "Job"),
                runtime=cell(row, "Runtime"),
                mean_ms=parse_time_ms(cell(row, "Mean")),
                error_ms=parse_time_ms(cell(row, "Error")),
                allocated_bytes=parse_bytes(cell(row, "Allocated")),
                row_count=cell(row, "RowCount"),
            )
            gen0 = cell(row, "Gen0")
            if gen0:
                try:
                    entry.gen0 = float(gen0)
                except ValueError:
                    entry.gen0 = 0.0
            if not result.runtime and entry.runtime:
                result.runtime = entry.runtime
            result.rows.append(entry)
        return result


def collect_benchmarks(directory: Path) -> List[BenchmarkSet]:
    """解析目录下全部 BDN CSV。

    文件名形如 ``Bing.Offices.Benchmarks.CsvPipelineBenchmarks-report.csv``，
    去掉 ``-report`` 后缀即为基准类全名。同名多次运行只保留排序最后的一份。
    """
    latest: Dict[str, BenchmarkSet] = {}
    for path in sorted(directory.glob("*-report.csv")):
        parsed = _read_bdn_csv(path)
        parsed.name = re.sub(r"-report$", "", path.stem)
        latest[parsed.name] = parsed
    return [latest[name] for name in sorted(latest)]
