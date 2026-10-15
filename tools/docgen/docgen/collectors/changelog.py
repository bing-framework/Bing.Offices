"""变更日志采集器：解析 git 历史（Conventional Commits）。

数据源：``git log`` / ``git tag``
产出：按类型分组的变更条目，用于自动生成 CHANGELOG 与发布说明。

设计要点：合并提交与机器人提交被过滤；未遵循 Conventional Commits 的提交
归入「其他」，不静默丢弃 —— 遗漏比分类错误更危险。
"""

from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

# type -> (中文名, 是否计入面向用户的发布说明)
CHANGE_TYPES: Dict[str, tuple] = {
    "feat": ("新功能", True),
    "fix": ("缺陷修复", True),
    "perf": ("性能优化", True),
    "refactor": ("重构", False),
    "docs": ("文档", False),
    "test": ("测试", False),
    "build": ("构建", False),
    "ci": ("持续集成", False),
    "chore": ("杂项", False),
    "style": ("代码风格", False),
    "revert": ("回滚", True),
}

_HEADER_RE = re.compile(
    r"^(?P<type>[a-zA-Z]+)(?:\((?P<scope>[^)]*)\))?(?P<breaking>!)?:\s*(?P<subject>.+)$"
)

# 本仓库使用「emoji + Conventional Commits」风格（如 "📝 docs(offices): ..."），
# 因此匹配前必须先剥离前导的 emoji / 装饰符号，否则整条提交会被误判为「其他」。
_LEADING_DECOR_RE = re.compile(r"^[^0-9A-Za-z\u4e00-\u9fff]+")

_SEPARATOR = "\x1f"
_RECORD = "\x1e"


@dataclass
class ChangeEntry:
    """一条提交。"""

    sha: str
    date: str
    subject: str
    raw_type: str = ""
    scope: str = ""
    breaking: bool = False

    @property
    def type_name(self) -> str:
        return self.raw_type if self.raw_type in CHANGE_TYPES else "other"

    @property
    def type_cn(self) -> str:
        return CHANGE_TYPES.get(self.raw_type, ("其他", True))[0]

    @property
    def is_user_facing(self) -> bool:
        return CHANGE_TYPES.get(self.raw_type, ("其他", True))[1]

    @property
    def short_sha(self) -> str:
        return self.sha[:8]


@dataclass
class ChangeLog:
    """一段提交历史。"""

    entries: List[ChangeEntry] = field(default_factory=list)
    head_sha: str = ""
    head_subject: str = ""
    branch: str = ""
    latest_tag: str = ""
    tag_count: int = 0

    def by_type(self, user_facing_only: bool = False) -> Dict[str, List[ChangeEntry]]:
        grouped: Dict[str, List[ChangeEntry]] = {}
        for entry in self.entries:
            if user_facing_only and not entry.is_user_facing:
                continue
            grouped.setdefault(entry.type_name, []).append(entry)
        return grouped

    @property
    def breaking_changes(self) -> List[ChangeEntry]:
        return [e for e in self.entries if e.breaking]

    def type_counts(self) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for entry in self.entries:
            counts[entry.type_name] = counts.get(entry.type_name, 0) + 1
        return dict(sorted(counts.items(), key=lambda kv: -kv[1]))

    @property
    def scope_counts(self) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for entry in self.entries:
            if entry.scope:
                counts[entry.scope] = counts.get(entry.scope, 0) + 1
        return dict(sorted(counts.items(), key=lambda kv: -kv[1]))

    @property
    def conventional_count(self) -> int:
        """符合 Conventional Commits 的提交数。"""
        return sum(1 for e in self.entries if e.raw_type in CHANGE_TYPES)

    @property
    def conventional_ratio(self) -> float:
        """提交规范符合率（%）。它是变更日志自动化的前提条件。"""
        if not self.entries:
            return 0.0
        return self.conventional_count / len(self.entries) * 100.0


def _run_git(root: Path, args: List[str]) -> str:
    try:
        completed = subprocess.run(
            ["git", *args],
            cwd=str(root),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except FileNotFoundError as error:  # pragma: no cover - 环境缺 git
        raise RuntimeError("未找到 git 可执行文件，无法采集变更日志") from error
    if completed.returncode != 0:
        return ""
    return completed.stdout


def collect_changelog(root: Path, limit: int = 200, since_tag: Optional[str] = None) -> ChangeLog:
    """采集 git 历史。"""
    changelog = ChangeLog()

    if not (root / ".git").exists():
        return changelog

    changelog.branch = _run_git(root, ["rev-parse", "--abbrev-ref", "HEAD"]).strip()

    tags = [t for t in _run_git(root, ["tag", "--sort=-creatordate"]).splitlines() if t.strip()]
    changelog.tag_count = len(tags)
    changelog.latest_tag = tags[0] if tags else ""

    revision = f"{since_tag}..HEAD" if since_tag else "HEAD"
    pretty = f"%H{_SEPARATOR}%ad{_SEPARATOR}%s{_RECORD}"
    output = _run_git(
        root,
        ["log", revision, f"--pretty=format:{pretty}", "--date=short", f"-n{limit}"],
    )

    for record in output.split(_RECORD):
        record = record.strip("\n\r")
        if not record:
            continue
        parts = record.split(_SEPARATOR)
        if len(parts) < 3:
            continue
        sha, date, subject = parts[0], parts[1], _SEPARATOR.join(parts[2:])
        sha, subject = sha.strip(), subject.strip()
        if not sha:
            continue

        entry = ChangeEntry(sha=sha, date=date.strip(), subject=subject)
        match = _HEADER_RE.match(_LEADING_DECOR_RE.sub("", subject))
        if match:
            entry.raw_type = match.group("type").lower()
            entry.scope = (match.group("scope") or "").strip()
            entry.breaking = bool(match.group("breaking"))
            entry.subject = match.group("subject").strip()
        changelog.entries.append(entry)

    if changelog.entries:
        changelog.head_sha = changelog.entries[0].sha
        changelog.head_subject = changelog.entries[0].subject

    return changelog
