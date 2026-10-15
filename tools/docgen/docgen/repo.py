"""仓库路径解析。

所有采集器通过 RepoLayout 取路径，禁止在采集器里拼接相对路径字符串。
这样从任意工作目录运行都能定位正确的位置。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


def find_repo_root(start: Path) -> Path:
    """从 start 向上寻找含 Bing.Offices.sln 的目录。"""
    current = start.resolve()
    for candidate in (current, *current.parents):
        if (candidate / "Bing.Offices.sln").exists():
            return candidate
    raise FileNotFoundError(f"未能从 {start} 向上找到 Bing.Offices.sln")


@dataclass(frozen=True)
class RepoLayout:
    """仓库内与文档生成相关的路径。"""

    root: Path

    # ---- 输入 ----
    @property
    def test_results_dir(self) -> Path:
        return self.root / "TestResults"

    @property
    def benchmark_dir(self) -> Path:
        return self.root / "BenchmarkDotNet.Artifacts" / "results"

    @property
    def release_output_dir(self) -> Path:
        return self.root / "output" / "release"

    @property
    def api_baseline(self) -> Path:
        return self.root / "build" / "api-snapshot-baseline.json"

    @property
    def delivery_profile(self) -> Path:
        """对外交付档案（客户抬头 / 条款 / 签章角色），缺失时使用内置默认条款。"""
        return self.root / "build" / "delivery-profile.json"

    @property
    def gates_config(self) -> Path:
        """质量门禁阈值；缺失时使用 gates.GateConfig 的内置默认值。"""
        return self.root / "build" / "docgen-gates.json"

    @property
    def benchmark_history_dir(self) -> Path:
        # 基准历史快照：构建产物，不入库（artifacts/ 被整体 gitignore）。
        # 本地保留；如需跨环境复用，应单独归档。
        return self.root / "artifacts" / "benchmark-history"

    @property
    def src_dir(self) -> Path:
        return self.root / "src"

    @property
    def docs_dir(self) -> Path:
        return self.root / "docs"

    @property
    def docs_tests_dir(self) -> Path:
        """文档示例代码的宿主测试项目：示例在这里必须可编译，腐化即编译失败。"""
        return self.root / "tests" / "Bing.Offices.Docs.Tests"

    # ---- 输出 ----
    @property
    def markdown_out(self) -> Path:
        return self.docs_dir / "generated"

    @property
    def binary_out(self) -> Path:
        return self.root / "artifacts" / "docs"

    def ensure_outputs(self) -> None:
        self.markdown_out.mkdir(parents=True, exist_ok=True)
        self.binary_out.mkdir(parents=True, exist_ok=True)

    # ---- 校验 ----
    def missing_inputs(self) -> list[str]:
        """返回缺失的输入项，便于给出可操作的报错。"""
        checks = {
            "TestResults 目录": self.test_results_dir,
            "BenchmarkDotNet.Artifacts/results 目录": self.benchmark_dir,
            "output/release 目录": self.release_output_dir,
            "build/api-snapshot-baseline.json": self.api_baseline,
        }
        return [name for name, path in checks.items() if not path.exists()]


def default_layout() -> RepoLayout:
    """以本文件位置为锚点解析仓库根目录。"""
    return RepoLayout(root=find_repo_root(Path(__file__).resolve().parent))
