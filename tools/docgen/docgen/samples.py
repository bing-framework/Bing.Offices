"""文档腐化防护：让文档里的示例代码**必须来自可编译的源码**。

腐化的根因是「文档与代码两条独立维护的时间线」：代码改了、文档没改，
文档就静默失真，而且没有任何信号。本模块把两条时间线合并成一条：

    唯一事实来源 = tests/Bing.Offices.Docs.Tests 中带 docgen:sample 标记的
                  可编译方法（改 API 会直接编译失败）
    文档中的示例 = 从上述源码抽取后 **生成** 出来的 markdown 代码块

于是「文档腐化」从「没人发现」变成两种必然被捕获的状态之一：
  * 改了 API 但没改示例  -> 编译失败（CI 必红）
  * 改了示例但没重新生成文档 -> docs-check 报「与源码不一致」（CI 必红）

用法::

    python -m docgen docs-check
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# // docgen:sample name=ExportWorkbook title=导出一个工作簿
_META_RE = re.compile(
    r"^\s*//\s*docgen:sample\s+name=(?P<name>[^\s]+)"
    r"(?:\s+title=(?P<title>.+?))?\s*$"
)
_FENCE_RE = re.compile(r"^```(?P<info>.*)$")
_USING_RE = re.compile(r"^\s*using\s+(?P<ns>[A-Za-z_][\w.]*)\s*;")

_SKIP_DIRS = {"bin", "obj", "Generated"}


@dataclass
class Sample:
    """一个可编译的文档示例。"""

    name: str
    title: str
    code: str
    source: str = ""   # 来源文件（相对仓库根）
    line: int = 0      # 源码中的行号（1 基）

    @property
    def fingerprint(self) -> str:
        """用于比对的代码指纹：忽略空行与行尾空白。"""
        return "\n".join(line.rstrip() for line in self.code.splitlines() if line.strip())


@dataclass
class SampleCheckResult:
    """docs-check 的结果。"""

    source_samples: List[Sample] = field(default_factory=list)
    doc_samples: Dict[str, List[Sample]] = field(default_factory=dict)  # md 文件 -> 示例
    problems: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.problems

    @property
    def embedded_count(self) -> int:
        return sum(len(v) for v in self.doc_samples.values())


# --------------------------------------------------------------------------
# 源码侧：从可编译的测试代码里抽取示例
# --------------------------------------------------------------------------
def _dedent(lines: List[str]) -> str:
    indents = [len(line) - len(line.lstrip()) for line in lines if line.strip()]
    shift = min(indents) if indents else 0
    return "\n".join(line[shift:] if len(line) >= shift else line.lstrip() for line in lines)


def _extract_method(lines: List[str], start: int) -> Optional[Tuple[int, int]]:
    """从 start 行开始，找到一个用大括号包裹的整体，返回 (首行, 末行)。"""
    begin = None
    for index in range(start, len(lines)):
        if "{" in lines[index]:
            begin = index
            break
    if begin is None:
        return None
    depth = 0
    for index in range(begin, len(lines)):
        depth += lines[index].count("{") - lines[index].count("}")
        if depth <= 0:
            return begin, index
    return None


def collect_source_samples(cs_dir: Path) -> List[Sample]:
    """扫描目录（默认 tests/Bing.Offices.Docs.Tests）下的 docgen:sample 标记。"""
    samples: List[Sample] = []
    if not cs_dir.exists():
        return samples

    for path in sorted(cs_dir.rglob("*.cs")):
        if any(part in _SKIP_DIRS for part in path.parts):
            continue
        try:
            text = path.read_text(encoding="utf-8-sig")
        except OSError:
            continue
        lines = text.splitlines()

        index = 0
        while index < len(lines):
            match = _META_RE.match(lines[index])
            if not match:
                index += 1
                continue
            # 跳过标记行之后的空行，定位方法签名
            cursor = index + 1
            while cursor < len(lines) and not lines[cursor].strip():
                cursor += 1
            span = _extract_method(lines, cursor)
            if span is None:
                index += 1
                continue
            begin, end = span
            samples.append(Sample(
                name=match.group("name"),
                title=(match.group("title") or match.group("name")).strip(),
                code=_dedent(lines[begin:end + 1]),
                source=path.name,
                line=begin + 1,
            ))
            index = end + 1
    return samples


# --------------------------------------------------------------------------
# 文档侧：从 markdown 中抽取 csharp 示例
# --------------------------------------------------------------------------
def collect_doc_samples(md_path: Path) -> List[Sample]:
    """抽取 markdown 中带 docgen:sample 元信息的 csharp 代码块。"""
    samples: List[Sample] = []
    try:
        lines = md_path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return samples

    index = 0
    while index < len(lines):
        match = _FENCE_RE.match(lines[index])
        if not match or not match.group("info").strip().lower().startswith(("csharp", "c#", "cs")):
            index += 1
            continue
        body: List[str] = []
        cursor = index + 1
        while cursor < len(lines) and not lines[cursor].startswith("```"):
            body.append(lines[cursor])
            cursor += 1

        # 元信息行只用于建立「文档示例 -> 源码示例」的关联，
        # 比对指纹时要剔除，否则永远与源码（不含该行）不相等
        meta = None
        meta_index = -1
        for offset, line in enumerate(body):
            match = _META_RE.match(line)
            if match:
                meta, meta_index = match, offset
                break
        if meta:
            samples.append(Sample(
                name=meta.group("name"),
                title=(meta.group("title") or meta.group("name")).strip(),
                code="\n".join(line for offset, line in enumerate(body) if offset != meta_index),
            ))
        index = cursor + 1
    return samples


def collect_doc_samples_from_dir(md_dir: Path) -> Dict[str, List[Sample]]:
    result: Dict[str, List[Sample]] = {}
    if not md_dir.exists():
        return result
    for path in sorted(md_dir.glob("*.md")):
        found = collect_doc_samples(path)
        if found:
            result[path.name] = found
    return result


# --------------------------------------------------------------------------
# 校验
# --------------------------------------------------------------------------
def _referenced_namespaces(code: str) -> List[str]:
    result = []
    for line in code.splitlines():
        match = _USING_RE.match(line)
        if match:
            result.append(match.group("ns"))
    return result


def check_samples(layout, required_namespaces: Optional[List[str]] = None) -> SampleCheckResult:
    """双向一致性校验：源码示例 <-> 文档示例。"""
    result = SampleCheckResult()
    source = collect_source_samples(layout.docs_tests_dir)
    result.source_samples = source
    docs = collect_doc_samples_from_dir(layout.markdown_out)
    result.doc_samples = docs

    if not source:
        result.problems.append(
            f"在 {layout.docs_tests_dir.name} 中没有找到任何 // docgen:sample 标记的示例，"
            f"文档腐化防护未生效"
        )
        return result

    by_name: Dict[str, Sample] = {}
    for sample in source:
        if sample.name in by_name:
            result.problems.append(f"示例名重复：{sample.name}（{sample.source}:{sample.line}）")
        by_name[sample.name] = sample

    embedded: Dict[str, Sample] = {}
    for file_name, items in docs.items():
        for sample in items:
            origin = by_name.get(sample.name)
            if origin is None:
                result.problems.append(
                    f"{file_name}：文档示例「{sample.name}」在 {layout.docs_tests_dir.name} "
                    f"中不存在（文档示例只能来自可编译源码）"
                )
                continue
            if sample.fingerprint != origin.fingerprint:
                result.problems.append(
                    f"{file_name}：文档示例「{sample.name}」与源码 {origin.source}:{origin.line} "
                    f"不一致，请重新生成文档"
                )
            embedded.setdefault(sample.name, sample)

    for name, sample in by_name.items():
        if name not in embedded:
            result.problems.append(
                f"源码示例「{name}」（{sample.source}:{sample.line}）未出现在任何生成文档中"
            )

    # 命名空间存在性校验：示例引用的 Bing.Offices 命名空间必须是真实存在的公共命名空间
    known = set(required_namespaces or [])
    for file_name, items in docs.items():
        for sample in items:
            for namespace in _referenced_namespaces(sample.code):
                if namespace.startswith("Bing.Offices") and known and namespace not in known:
                    result.problems.append(
                        f"{file_name}：示例「{sample.name}」引用了不存在的命名空间 {namespace}"
                    )
    return result
