"""质量门禁（gates）。

阶段三的「注释覆盖率门禁」落地在这里：把 XML 文档注释覆盖率算成**可执行判定**，
低于阈值就让流水线退出非 0，从而在 CI 上阻断合入。

口径说明（重要）：
    XML 文档文件里**只包含写了注释的成员**，完全没写注释的成员不会出现，
    因此「XML 条目中有 summary 的比例」永远接近 100%，不能作为覆盖率。
    真正的分母必须来自公共 API 快照（build/api-snapshot-baseline.json），
    它按反射枚举了全部对外可见成员。本模块采用：

        覆盖率 = 有 summary 的对外可见成员数 / 快照记录的对外可见成员数

快照行格式（示例）::

    type|Bing.Offices.Attributes.ColumnNameAttribute|kind=class|visibility=public|...
    method|Bing.Offices.Configurations.ExcelMappingConfigurationLoader.FromJsonDocument|...
    property|Bing.Offices.Attributes.ColumnNameAttribute.Name|type=System.String|...
    field|Bing.Offices.Dates.ExcelDateOffsetPolicy.RequireExplicitOffset|...
    constructor|Bing.Offices.Attributes.ColumnNameAttribute|visibility=public|...

与 XML doc comment ID 的映射：type→T:、method→M:、property→P:、field→F:、
constructor→M:<Type>.#ctor。
"""

from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .collectors.changelog import collect_changelog
from .repo import RepoLayout

# 快照成员种类 -> XML doc comment ID 前缀
_KIND_PREFIX = {
    "type": "T:",
    "method": "M:",
    "property": "P:",
    "field": "F:",
    "event": "E:",
    "constructor": "M:",
    "operator": "M:",
}

# 泛型回记号：`1 / ``2 等，两侧写法不同，统一去掉
_GENERIC_RE = re.compile(r"`+\d+")

# 阈值取「当前实测值减去约 5 个百分点」：现在必定通过，
# 但任何一次实质性的注释回退都会被拦下。实测基线见 docs/doc-automation/README.md。
DEFAULT_MIN_COVERAGE = 90.0
DEFAULT_MIN_PER_ASSEMBLY = 85.0
# 提交规范符合率实测约 77%，先按 70% 设闸，历史提交规范后再上调到 80%。
DEFAULT_MIN_CONVENTIONAL = 70.0


# --------------------------------------------------------------------------
# 结果模型
# --------------------------------------------------------------------------
@dataclass
class GateResult:
    """一项门禁的判定结果。"""

    name: str
    passed: bool
    value: float
    threshold: float
    unit: str = "%"
    detail: str = ""
    problems: List[str] = field(default_factory=list)

    def render(self) -> str:
        flag = "通过" if self.passed else "未通过"
        text = f"[{flag}] {self.name}：{self.value:.2f}{self.unit}（阈值 {self.threshold:.2f}{self.unit}）"
        if self.detail:
            text += f"  {self.detail}"
        return text


@dataclass
class AssemblyCoverage:
    """单个程序集的注释覆盖率。"""

    name: str
    tfm: str
    total: int = 0        # 参与统计的对外可见成员数（不含无参构造）
    documented: int = 0   # 其中有 summary 的成员数
    unmatched: int = 0    # 快照里存在、但 XML 中根本没出现的成员（即完全没写注释）
    ctor_total: int = 0       # 构造函数数（参考值，不计入主指标）
    ctor_documented: int = 0

    @property
    def pct(self) -> float:
        return self.documented / self.total * 100.0 if self.total else 0.0

    @property
    def ctor_pct(self) -> float:
        return self.ctor_documented / self.ctor_total * 100.0 if self.ctor_total else 0.0


@dataclass
class CommentCoverage:
    """全仓库注释覆盖率快照。"""

    tfm: str = "net8.0"
    assemblies: List[AssemblyCoverage] = field(default_factory=list)
    available: bool = False
    reason: str = ""

    @property
    def total(self) -> int:
        return sum(a.total for a in self.assemblies)

    @property
    def documented(self) -> int:
        return sum(a.documented for a in self.assemblies)

    @property
    def pct(self) -> float:
        return self.documented / self.total * 100.0 if self.total else 0.0

    @property
    def ctor_total(self) -> int:
        return sum(a.ctor_total for a in self.assemblies)

    @property
    def ctor_documented(self) -> int:
        return sum(a.ctor_documented for a in self.assemblies)

    @property
    def ctor_pct(self) -> float:
        return self.ctor_documented / self.ctor_total * 100.0 if self.ctor_total else 0.0


# --------------------------------------------------------------------------
# 采集与匹配
# --------------------------------------------------------------------------
def _normalize(text: str) -> str:
    return _GENERIC_RE.sub("", text or "")


def _snapshot_keys(lines: List[str], include_protected: bool,
                   exclude_constructors: bool = True) -> Tuple[List[str], List[str]]:
    """把快照行转成 XML doc comment ID 形态的键。

    返回 ``(参与统计的键, 构造函数键)``。构造函数默认**不计入统计口径**：
    绝大多数类型的无参构造是编译器生成的，无法为其编写 XML 注释，
    把它们放进分母只会让覆盖率被系统性低估（实测会把 95% 拉低到 92%）。
    构造覆盖率作为参考值单独输出。
    """
    keys: List[str] = []
    ctor_keys: List[str] = []
    for line in lines:
        parts = line.split("|")
        if len(parts) < 2:
            continue
        kind = parts[0].strip()
        name = _normalize(parts[1].strip())
        prefix = _KIND_PREFIX.get(kind)
        if prefix is None:
            continue
        visibility = ""
        for part in parts[2:]:
            if part.startswith("visibility="):
                visibility = part.split("=", 1)[1].strip()
                break
        allowed = ("public", "protected") if include_protected else ("public",)
        if visibility and visibility not in allowed:
            continue
        if kind == "constructor":
            ctor_keys.append(f"M:{name}.#ctor")
            if not exclude_constructors:
                keys.append(f"M:{name}.#ctor")
        else:
            keys.append(f"{prefix}{name}")
    return keys, ctor_keys


def _documented_keys(xml_path: Path) -> Dict[str, bool]:
    """XML 文档文件 -> {规范化成员键: 是否写了 summary}。"""
    result: Dict[str, bool] = {}
    try:
        root = ET.parse(xml_path).getroot()
    except (ET.ParseError, OSError):
        return result
    members = root.find("members")
    if members is None:
        return result

    for member in members.iter("member"):
        member_id = _normalize(member.get("name", ""))
        if not member_id:
            continue
        # 方法带参数列表：M:X.Y(System.String) -> M:X.Y
        key = member_id.split("(", 1)[0]
        summary_node = member.find("summary")
        summary = ""
        if summary_node is not None:
            summary = " ".join("".join(summary_node.itertext()).split())
        # 同一键可能被多个重载命中，任一有 summary 即视为已注释
        result[key] = result.get(key, False) or bool(summary)
    return result


def collect_comment_coverage(layout: RepoLayout, tfm: str = "net8.0",
                             include_protected: bool = True) -> CommentCoverage:
    """计算注释覆盖率。"""
    coverage = CommentCoverage(tfm=tfm)

    baseline = layout.api_baseline
    if not baseline.exists():
        coverage.reason = f"缺少 API 快照：{baseline.name}"
        return coverage

    try:
        with baseline.open("r", encoding="utf-8-sig") as handle:
            data = json.load(handle)
    except (ValueError, OSError) as error:
        coverage.reason = f"API 快照解析失败：{error}"
        return coverage

    assemblies = data.get("assemblies", {}).get(tfm, {})
    if not assemblies:
        coverage.reason = f"API 快照中没有 {tfm} 的程序集记录"
        return coverage

    xml_dir = layout.release_output_dir / tfm
    for name, meta in sorted(assemblies.items()):
        keys, ctor_keys = _snapshot_keys(meta.get("lines", []), include_protected)
        entry = AssemblyCoverage(name=name, tfm=tfm, total=len(keys),
                                 ctor_total=len(ctor_keys))
        xml_path = xml_dir / f"{name}.xml"
        if xml_path.exists():
            documented = _documented_keys(xml_path)
            entry.documented = sum(1 for key in keys if documented.get(key))
            entry.ctor_documented = sum(1 for key in ctor_keys if documented.get(key))
        entry.unmatched = len(keys) - entry.documented
        coverage.assemblies.append(entry)

    coverage.available = bool(coverage.assemblies) and coverage.total > 0
    return coverage


# --------------------------------------------------------------------------
# 阈值配置
# --------------------------------------------------------------------------
@dataclass
class GateConfig:
    """门禁阈值。"""

    min_coverage: float = DEFAULT_MIN_COVERAGE
    min_per_assembly: float = DEFAULT_MIN_PER_ASSEMBLY
    min_conventional: float = DEFAULT_MIN_CONVENTIONAL

    @classmethod
    def load(cls, layout: RepoLayout) -> "GateConfig":
        path = layout.gates_config
        if not path.exists():
            return cls()
        try:
            with path.open("r", encoding="utf-8-sig") as handle:
                data = json.load(handle)
        except (ValueError, OSError):
            return cls()
        comment = data.get("commentCoverage", {}) or {}
        commits = data.get("conventionalCommits", {}) or {}
        return cls(
            min_coverage=float(comment.get("min", DEFAULT_MIN_COVERAGE)),
            min_per_assembly=float(comment.get("minPerAssembly", DEFAULT_MIN_PER_ASSEMBLY)),
            min_conventional=float(commits.get("min", DEFAULT_MIN_CONVENTIONAL)),
        )


# --------------------------------------------------------------------------
# 门禁判定
# --------------------------------------------------------------------------
def gate_comment_coverage(layout: RepoLayout, config: Optional[GateConfig] = None,
                          tfm: str = "net8.0") -> Tuple[GateResult, CommentCoverage]:
    config = config or GateConfig.load(layout)
    coverage = collect_comment_coverage(layout, tfm=tfm)

    if not coverage.available:
        return GateResult(
            name="XML 注释覆盖率",
            passed=False,
            value=0.0,
            threshold=config.min_coverage,
            detail=coverage.reason or "无可用数据",
            problems=[coverage.reason or "未能采集到公共 API 快照或 XML 注释产物"],
        ), coverage

    problems = []
    for assembly in coverage.assemblies:
        if assembly.pct < config.min_per_assembly:
            problems.append(
                f"{assembly.name}：{assembly.pct:.2f}%"
                f"（{assembly.documented}/{assembly.total}，阈值 {config.min_per_assembly:.2f}%）"
            )

    passed = coverage.pct >= config.min_coverage and not problems
    return GateResult(
        name="XML 注释覆盖率",
        passed=passed,
        value=coverage.pct,
        threshold=config.min_coverage,
        detail=f"{coverage.documented}/{coverage.total} 个对外可见成员已写 summary，"
               f"覆盖 {len(coverage.assemblies)} 个程序集（{tfm}）；"
               f"构造函数注释覆盖率 {coverage.ctor_pct:.1f}% 仅供参考，不计入判定",
        problems=problems,
    ), coverage


def gate_conventional_commits(layout: RepoLayout,
                              config: Optional[GateConfig] = None) -> GateResult:
    config = config or GateConfig.load(layout)
    changelog = collect_changelog(layout.root, limit=200)
    ratio = changelog.conventional_ratio
    total = len(changelog.entries)
    return GateResult(
        name="提交规范符合率",
        passed=ratio >= config.min_conventional,
        value=ratio,
        threshold=config.min_conventional,
        detail=f"{changelog.conventional_count}/{total} 条提交可自动归类（Conventional Commits）",
        problems=[] if ratio >= config.min_conventional else [
            f"符合率 {ratio:.2f}% 低于阈值 {config.min_conventional:.2f}%，"
            f"会导致发布说明出现大量「其他」类条目"
        ],
    )


def run_gates(layout: RepoLayout, config: Optional[GateConfig] = None
              ) -> Tuple[List[GateResult], Optional[CommentCoverage]]:
    """执行全部门禁，返回结果列表与注释覆盖率明细（供报告复用）。"""
    config = config or GateConfig.load(layout)
    coverage_result, coverage = gate_comment_coverage(layout, config)
    commits_result = gate_conventional_commits(layout, config)
    return [coverage_result, commits_result], coverage
