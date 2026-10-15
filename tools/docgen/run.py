"""自举式运行入口 —— 一条命令即可生成全部文档。

    python tools/docgen/run.py all
    python tools/docgen/run.py tests --format docx,xlsx
    python tools/docgen/run.py all --check

首次运行会自动创建虚拟环境并安装依赖（放在仓库内的 tools/docgen/.venv，
避免污染系统 Python，也不会被 git 跟踪）。
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
VENV = HERE / ".venv"
REQUIREMENTS = HERE / "requirements.txt"
STAMP = VENV / ".installed"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def _deps_available(python: Path) -> bool:
    """当前解释器是否已装齐依赖（避免重复建 venv，也加速 CI）。"""
    import importlib.util

    for module in ("docx", "openpyxl", "pptx", "matplotlib"):
        if importlib.util.find_spec(module) is None:
            return False
    return True


def ensure_venv() -> Path:
    python = venv_python()
    if python.exists() and STAMP.exists():
        return python

    # 若当前解释器已装齐依赖，直接用，不建 venv（CI 里 pip install -r 之后的情况）
    if _deps_available(Path(sys.executable)):
        return Path(sys.executable)

    print("[docgen] 首次运行，正在准备虚拟环境 …")
    subprocess.run([sys.executable, "-m", "venv", str(VENV)], check=True)
    subprocess.run([str(python), "-m", "pip", "install", "--quiet", "--upgrade", "pip"],
                   check=True)
    subprocess.run([str(python), "-m", "pip", "install", "--quiet", "-r", str(REQUIREMENTS)],
                   check=True)
    STAMP.write_text("ok", encoding="utf-8")
    print("[docgen] 环境就绪。")
    return python


def main() -> int:
    python = ensure_venv()
    env = dict(os.environ)
    # 让子进程能 import docgen 包
    env["PYTHONPATH"] = str(HERE) + os.pathsep + env.get("PYTHONPATH", "")
    return subprocess.run(
        [str(python), "-m", "docgen", *sys.argv[1:]],
        cwd=str(HERE),
        env=env,
    ).returncode


if __name__ == "__main__":
    raise SystemExit(main())
