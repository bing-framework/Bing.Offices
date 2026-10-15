"""代码覆盖率采集器：解析 Cobertura XML。

数据源：``dotnet test --collect:"XPlat Code Coverage"`` 产出的 ``coverage.cobertura.xml``。
该文件通常位于 ``TestResults/<guid>/coverage.cobertura.xml``。

产出：按程序集/命名空间/类/方法的逐级覆盖率，
用于在测试报告中新增「覆盖率」章节与未覆盖类型清单。

注意：本项目 CI 当前**尚未**产出 Cobertura（见阶段二路线图），
因此本采集器在找不到文件时返回空 inventory，报告章节相应降级。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

_REPORT_ATTRS = ("line-rate", "branch-rate", "lines-covered", "lines-valid",
                 "branches-covered", "branches-valid")


def _cobertura_files(directory: Path) -> List[Path]:
    """在 TestResults 下递归寻找 coverage.cobertura.xml。"""
    return sorted(directory.rglob("coverage.cobertura.xml"))


def _rate(value: str) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


@dataclass
class ClassCoverage:
    name: str
    filename: str
    line_rate: float = 0.0
    branch_rate: float = 0.0
    lines_covered: int = 0
    lines_valid: int = 0
    branches_covered: int = 0
    branches_valid: int = 0

    @property
    def line_pct(self) -> float:
        return self.line_rate * 100.0

    @property
    def short_name(self) -> str:
        return self.name.rsplit(".", 1)[-1]

    @property
    def namespace(self) -> str:
        parts = self.name.split(".")
        return ".".join(parts[:-1]) if len(parts) > 1 else "(全局)"


@dataclass
class PackageCoverage:
    name: str
    classes: List[ClassCoverage] = field(default_factory=list)

    @property
    def line_rate(self) -> float:
        valid = sum(c.lines_valid for c in self.classes)
        covered = sum(c.lines_covered for c in self.classes)
        return (covered / valid) if valid else 0.0

    @property
    def line_pct(self) -> float:
        return self.line_rate * 100.0

    @property
    def class_count(self) -> int:
        return len(self.classes)


@dataclass
class CoverageInventory:
    """一次覆盖率采集的汇总。"""

    source: str = ""
    timestamp: str = ""
    line_rate: float = 0.0
    branch_rate: float = 0.0
    lines_covered: int = 0
    lines_valid: int = 0
    branches_covered: int = 0
    branches_valid: int = 0
    packages: List[PackageCoverage] = field(default_factory=list)

    @property
    def line_pct(self) -> float:
        return self.line_rate * 100.0

    @property
    def branch_pct(self) -> float:
        return self.branch_rate * 100.0

    @property
    def class_count(self) -> int:
        return sum(p.class_count for p in self.packages)

    @property
    def is_available(self) -> bool:
        return self.lines_valid > 0

    def uncovered_classes(self, threshold: float = 0.5) -> List[ClassCoverage]:
        """覆盖率低于阈值（默认 < 50%）的类，按覆盖率升序。"""
        result = []
        for package in self.packages:
            for cls in package.classes:
                if cls.lines_valid > 0 and cls.line_rate < threshold:
                    result.append(cls)
        return sorted(result, key=lambda c: c.line_rate)


def collect_coverage(test_results_dir: Path) -> CoverageInventory:
    """解析 TestResults 下第一个可用的 Cobertura 文件。"""
    inventory = CoverageInventory()

    files = _cobertura_files(test_results_dir)
    if not files:
        return inventory

    path = files[0]
    inventory.source = str(path)
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError:
        return inventory

    inventory.timestamp = root.get("timestamp", "")
    for attr in _REPORT_ATTRS:
        value = root.get(attr)
        if attr == "line-rate":
            inventory.line_rate = _rate(value or "0")
        elif attr == "branch-rate":
            inventory.branch_rate = _rate(value or "0")
        elif attr == "lines-covered":
            inventory.lines_covered = int(float(value or 0))
        elif attr == "lines-valid":
            inventory.lines_valid = int(float(value or 0))
        elif attr == "branches-covered":
            inventory.branches_covered = int(float(value or 0))
        elif attr == "branches-valid":
            inventory.branches_valid = int(float(value or 0))

    for package_elem in root.iter("package"):
        package_name = package_elem.get("name", "(未命名)")
        package = PackageCoverage(name=package_name)
        for cls_elem in package_elem.iter("class"):
            package.classes.append(ClassCoverage(
                name=cls_elem.get("name", ""),
                filename=cls_elem.get("filename", ""),
                line_rate=_rate(cls_elem.get("line-rate", "0")),
                branch_rate=_rate(cls_elem.get("branch-rate", "0")),
                lines_covered=int(float(cls_elem.get("lines-covered", 0))),
                lines_valid=int(float(cls_elem.get("lines-valid", 0))),
                branches_covered=int(float(cls_elem.get("branches-covered", 0))),
                branches_valid=int(float(cls_elem.get("branches-valid", 0))),
            ))
        if package.classes:
            inventory.packages.append(package)

    # 按行覆盖率降序，便于报告呈现
    inventory.packages.sort(key=lambda p: -p.line_rate)
    return inventory
