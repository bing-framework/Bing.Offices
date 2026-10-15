"""基准历史归档与趋势对比。

问题：BenchmarkDotNet 的 ``-report.csv`` 每次运行会**覆盖**，历史信息只在
``.log`` 文件名的时间戳里。要回答「这次改动是否引入性能回归」，必须
**主动归档**每次的基准结果快照。

设计：
  * ``archives.py`` 每次运行把当期基准结果（附 commit / 时间戳 / 机器指纹）
    追加为 ``artifacts/benchmark-history/<基准类>/<commit>-<utc时间>.json``；
  * 报告渲染时读取该目录，与「上一次快照」做方法级对比，
    超过阈值的方法标记为回归 / 优化。

这是「从现在开始积累」的机制，首份归档即基线。
"""

from __future__ import annotations

import json
import subprocess
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .benchmark import BenchmarkRow, BenchmarkSet, collect_benchmarks

HISTORY_DIRNAME = "benchmark-history"


def _git_info(root: Path) -> Tuple[str, str]:
    """(短 commit, 完整 commit)。"""
    def run(args):
        try:
            return subprocess.run(
                ["git", *args], cwd=str(root), capture_output=True, text=True,
                encoding="utf-8", errors="replace", check=False,
            ).stdout.strip()
        except FileNotFoundError:
            return ""
    return run(["rev-parse", "--short", "HEAD"]), run(["rev-parse", "HEAD"])


def _now_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


@dataclass
class BenchmarkSnapshot:
    """一次运行的一组基准结果快照。"""

    name: str
    runtime: str
    commit: str
    captured_at: str
    rows: List[Dict[str, object]] = field(default_factory=list)

    @property
    def display_name(self) -> str:
        return self.name.rsplit(".", 1)[-1]


@dataclass
class TrendPoint:
    """同一方法在两次快照间的对比。"""

    method: str
    previous_ms: float
    current_ms: float

    @property
    def delta_ratio(self) -> float:
        """变化率，正数表示变慢。"""
        if self.previous_ms <= 0:
            return 0.0
        return (self.current_ms - self.previous_ms) / self.previous_ms * 100.0

    @property
    def verdict(self) -> str:
        """回归 / 优化 / 持平。"""
        if self.delta_ratio > 10.0:
            return "regression"
        if self.delta_ratio < -10.0:
            return "improvement"
        return "stable"


# --------------------------------------------------------------------------
# 归档
# --------------------------------------------------------------------------
def snapshot_from_sets(sets: List[BenchmarkSet], root: Path) -> List[BenchmarkSnapshot]:
    short_commit, full_commit = _git_info(root)
    captured = _now_utc()
    snapshots: List[BenchmarkSnapshot] = []
    for bench in sets:
        rows = []
        for row in bench.rows:
            if row.mean_ms is None:
                continue
            rows.append({
                "method": row.method,
                "job": row.job,
                "mean_ms": row.mean_ms,
                "allocated_bytes": row.allocated_bytes,
                "gen0": row.gen0,
            })
        snapshots.append(BenchmarkSnapshot(
            name=bench.name, runtime=bench.runtime, commit=full_commit,
            captured_at=captured, rows=rows,
        ))
    return snapshots


def write_snapshots(snapshots: List[BenchmarkSnapshot], history_dir: Path) -> List[Path]:
    """把快照写入历史目录，返回写入的文件路径。

    幂等约定：同一基准类 + 同一 commit 的归档**覆盖**（不无限累积），
    因此重复运行不会产生重复快照；commit 变化时才新增一个历史点。
    """
    written: List[Path] = []
    for snapshot in snapshots:
        target_dir = history_dir / snapshot.name
        target_dir.mkdir(parents=True, exist_ok=True)
        short = snapshot.commit[:12] if snapshot.commit else "unknown"

        # 删除同 commit 的旧快照（若有），实现「同 commit 覆盖」
        for existing in target_dir.glob(f"*-{short}.json"):
            existing.unlink()

        path = target_dir / f"{snapshot.captured_at}-{short}.json"
        payload = {
            "name": snapshot.name,
            "runtime": snapshot.runtime,
            "commit": snapshot.commit,
            "captured_at": snapshot.captured_at,
            "rows": snapshot.rows,
        }
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        written.append(path)
    return written


# --------------------------------------------------------------------------
# 趋势对比
# --------------------------------------------------------------------------
def _load_snapshot(path: Path) -> Optional[BenchmarkSnapshot]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    return BenchmarkSnapshot(
        name=data.get("name", ""),
        runtime=data.get("runtime", ""),
        commit=data.get("commit", ""),
        captured_at=data.get("captured_at", ""),
        rows=data.get("rows", []),
    )


def load_history(history_dir: Path) -> Dict[str, List[BenchmarkSnapshot]]:
    """按基准类名加载全部历史快照，按时间排序。"""
    grouped: Dict[str, List[BenchmarkSnapshot]] = {}
    if not history_dir.exists():
        return grouped
    for path in sorted(history_dir.glob("*/*.json")):
        snapshot = _load_snapshot(path)
        if snapshot is None:
            continue
        grouped.setdefault(snapshot.name, []).append(snapshot)
    for name in grouped:
        grouped[name].sort(key=lambda s: s.captured_at)
    return grouped


def compare_with_previous(current: List[BenchmarkSnapshot],
                          history_dir: Path) -> Dict[str, List[TrendPoint]]:
    """把当前快照与「每个基准类的上一次归档」做方法级对比。"""
    grouped = load_history(history_dir)
    result: Dict[str, List[TrendPoint]] = {}
    for snapshot in current:
        previous_list = grouped.get(snapshot.name, [])
        if not previous_list:
            continue
        previous = previous_list[-1]
        prev_by_method = {row["method"]: row.get("mean_ms") for row in previous.rows}
        points: List[TrendPoint] = []
        for row in snapshot.rows:
            method = row.get("method", "")
            previous_ms = prev_by_method.get(method)
            if previous_ms is None:
                continue
            points.append(TrendPoint(
                method=method,
                previous_ms=float(previous_ms),
                current_ms=float(row.get("mean_ms", 0.0)),
            ))
        if points:
            result[snapshot.name] = points
    return result
