"""测试结果采集器：解析 VSTest TRX。

数据源：TestResults/*.trx
产出：TestRun 列表（含逐用例明细与失败原因）
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional

TRX_NS = "{http://microsoft.com/schemas/VisualStudio/TeamTest/2010}"

_DURATION_RE = re.compile(r"^(?:(\d+)\.)?(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?$")

_PASSED = {"Passed"}
_FAILED = {"Failed"}
_SKIPPED = {"NotExecuted", "Skipped", "Inconclusive", "Aborted"}


def parse_duration(value: str) -> float:
    """'00:00:00.0607335' 或 '1.02:03:04.5' -> 秒。"""
    if not value:
        return 0.0
    match = _DURATION_RE.match(value.strip())
    if not match:
        return 0.0
    days = int(match.group(1) or 0)
    hours, minutes, seconds = int(match.group(2)), int(match.group(3)), int(match.group(4))
    fraction = float("0." + match.group(5)) if match.group(5) else 0.0
    return days * 86400 + hours * 3600 + minutes * 60 + seconds + fraction


def format_duration(seconds: float) -> str:
    """秒 -> 人类可读。"""
    if seconds < 1:
        return f"{seconds * 1000:.0f} ms"
    if seconds < 60:
        return f"{seconds:.2f} s"
    minutes, rest = divmod(seconds, 60)
    if minutes < 60:
        return f"{int(minutes)} m {rest:.1f} s"
    hours, minutes = divmod(minutes, 60)
    return f"{int(hours)} h {int(minutes)} m"


@dataclass
class TestCase:
    """单个测试用例。"""

    name: str
    full_name: str
    class_name: str
    outcome: str
    duration_s: float = 0.0
    message: str = ""
    stack_trace: str = ""

    @property
    def short_name(self) -> str:
        return self.name.rsplit(".", 1)[-1]

    @property
    def is_failure(self) -> bool:
        return self.outcome in _FAILED

    def first_error_line(self, limit: int = 160) -> str:
        """取错误消息的首行，压平空白。"""
        source = (self.message or self.stack_trace or "").strip()
        if not source:
            return ""
        line = source.splitlines()[0]
        line = re.sub(r"\s+", " ", line).strip()
        return line[:limit] + ("…" if len(line) > limit else "")


@dataclass
class TestRun:
    """一个 TRX 文件对应的测试运行。"""

    source: str
    total: int = 0
    passed: int = 0
    failed: int = 0
    skipped: int = 0
    duration_s: float = 0.0
    created: str = ""
    cases: List[TestCase] = field(default_factory=list)

    @property
    def executed(self) -> int:
        return self.total - self.skipped

    @property
    def pass_rate(self) -> float:
        return (self.passed / self.executed * 100.0) if self.executed else 0.0

    @property
    def is_success(self) -> bool:
        return self.failed == 0

    @property
    def failures(self) -> List[TestCase]:
        return [c for c in self.cases if c.is_failure]

    @property
    def target_framework(self) -> str:
        """从文件名推断目标框架，如 round2-existing-net8.trx -> net8。"""
        stem = Path(self.source).stem
        match = re.search(r"(net\d+(?:\.\d+)?)", stem)
        return match.group(1) if match else ""

    @property
    def suite(self) -> str:
        """套件名：round2-existing / round2-integration / round2-docs。"""
        stem = Path(self.source).stem
        return re.sub(r"-net\d+(?:\.\d+)?$", "", stem)


def _class_name(test_name: str) -> str:
    parts = test_name.split(".")
    return ".".join(parts[:-1]) if len(parts) > 1 else ""


def collect_test_runs(directory: Path, pattern: str = "*.trx") -> List[TestRun]:
    """解析目录下全部 TRX。"""
    runs: List[TestRun] = []
    for path in sorted(directory.glob(pattern)):
        runs.append(_parse_trx(path))
    return runs


def _parse_trx(path: Path) -> TestRun:
    tree = ET.parse(path)
    root = tree.getroot()

    times = root.find(f"{TRX_NS}Times")
    created = times.get("creation", "") if times is not None else ""

    run = TestRun(source=path.name, created=created)

    for element in root.iter(f"{TRX_NS}UnitTestResult"):
        outcome = element.get("outcome", "")
        full_name = element.get("testName", "")
        case = TestCase(
            name=full_name,
            full_name=full_name,
            class_name=_class_name(full_name),
            outcome=outcome,
            duration_s=parse_duration(element.get("duration", "")),
        )

        error_info = element.find(f"{TRX_NS}Output/{TRX_NS}ErrorInfo")
        if error_info is not None:
            message = error_info.find(f"{TRX_NS}Message")
            stack = error_info.find(f"{TRX_NS}StackTrace")
            case.message = (message.text or "") if message is not None else ""
            case.stack_trace = (stack.text or "") if stack is not None else ""

        run.cases.append(case)
        run.duration_s += case.duration_s

        if outcome in _PASSED:
            run.passed += 1
        elif outcome in _FAILED:
            run.failed += 1
        elif outcome in _SKIPPED:
            run.skipped += 1

    run.total = len(run.cases)
    return run


def summarize(runs: List[TestRun]) -> dict:
    """跨 TRX 汇总。"""
    total = sum(r.total for r in runs)
    passed = sum(r.passed for r in runs)
    failed = sum(r.failed for r in runs)
    skipped = sum(r.skipped for r in runs)
    executed = total - skipped
    return {
        "runs": len(runs),
        "total": total,
        "passed": passed,
        "failed": failed,
        "skipped": skipped,
        "executed": executed,
        "pass_rate": (passed / executed * 100.0) if executed else 0.0,
        "duration_s": sum(r.duration_s for r in runs),
        "has_failure": failed > 0,
        "all_failures": [c for r in runs for c in r.failures],
    }
